import { parse } from "acorn";

/**
 * Kodetilstandens sandkasse (ADR-0017): statisk kontrol og omskrivning af et Office.js-script,
 * før add-in'et kører det.
 *
 * Scriptet er modeloutput og kan være påvirket af prompt injection i dokumentet. Det må derfor kun
 * kunne nå de objekter, det får (`context` og værtens objekt, fx `Excel`) og en håndfuld sikre indbyggede objekter –
 * aldrig `window`, `fetch`, `localStorage` (indstillinger og API-nøgler) eller DOM'en.
 *
 * Kontrollen er kapabilitetsbaseret:
 *  - Hvert navn, scriptet bruger, skal være erklæret i scriptet selv (med korrekt blok-scope), være
 *    en parameter eller stå på listen over sikre globale objekter. Alt andet afvises – også
 *    globale objekter, vi ikke kender endnu.
 *  - Egenskaber, der kan føre ud af sandkassen (`constructor`, `prototype`, `__proto__`,
 *    `getPrototypeOf` …) og interne Office.js-egenskaber (`_…`, `m_…`) er spærret. Ved beregnede
 *    navne (`obj[x]`) tjekkes navnet, når scriptet kører (`__key`).
 *  - `this`, klasser, `import()`, `new.target`/`import.meta` og `with` er ikke tilladt.
 *  - Løkker får en tæller (`__tick`), så en uendelig løkke stopper i stedet for at fryse Office.
 *
 * Det er et forsvar i dybden, ikke en garanti: brugeren ser og godkender scriptet, og politikken
 * kan slå kodetilstanden fra.
 */

/** Globale objekter, scriptet må bruge. Ingen af dem giver adgang til netværk, DOM eller lager. */
export const SCRIPT_GLOBALS: ReadonlySet<string> = new Set([
  "Math",
  "JSON",
  "Number",
  "String",
  "Boolean",
  "Array",
  "Object",
  "Date",
  "Map",
  "Set",
  "Promise",
  "Error",
  "RangeError",
  "TypeError",
  "isNaN",
  "isFinite",
  "parseInt",
  "parseFloat",
  "undefined",
  "NaN",
  "Infinity",
  "console",
]);

/** Værtens globale objekt i scriptet: `Excel`, `Word` eller `PowerPoint` (kun det, der hører til værten). */
export const SCRIPT_HOSTS = { excel: "Excel", word: "Word", powerpoint: "PowerPoint" } as const;
export type ScriptHost = (typeof SCRIPT_HOSTS)[keyof typeof SCRIPT_HOSTS];

/** Egenskaber, der kan føre fra et tilladt objekt ud til funktioner eller det globale objekt. */
const BANNED_MEMBERS = new Set([
  "constructor",
  "prototype",
  "__proto__",
  "__defineGetter__",
  "__defineSetter__",
  "__lookupGetter__",
  "__lookupSetter__",
  "getPrototypeOf",
  "setPrototypeOf",
  "getOwnPropertyDescriptor",
  "getOwnPropertyDescriptors",
  "defineProperty",
  "defineProperties",
  "caller",
  "callee",
  "arguments",
  "prepareStackTrace",
  "captureStackTrace",
  "stackTraceLimit",
  "getThis",
  "getFunction",
  "ownerDocument",
  "defaultView",
]);

/** Spærret egenskabsnavn? Også interne Office.js-felter (`_…`, `$…`, `m_…`). */
export function isBannedKey(key: unknown): boolean {
  const k = String(key);
  return BANNED_MEMBERS.has(k) || /^[_$]/.test(k) || /^m_/.test(k);
}

/** Maks. antal gennemløb af alle løkker i ét script. */
export const MAX_SCRIPT_ITERATIONS = 1_000_000;

export type ScriptCheck = { ok: true; code: string } | { ok: false; error: string };

type Node = { type: string; start: number; end: number; [k: string]: unknown };

const FORBIDDEN_NODES: Record<string, string> = {
  ThisExpression: "this",
  ClassDeclaration: "klasser",
  ClassExpression: "klasser",
  ImportExpression: "import()",
  MetaProperty: "new.target/import.meta",
  WithStatement: "with",
  Super: "super",
  ImportDeclaration: "import",
  ExportNamedDeclaration: "export",
  ExportDefaultDeclaration: "export",
  ExportAllDeclaration: "export",
};

/**
 * Tjekker scriptet og returnerer den omskrevne kode (med `__key` og `__tick`), som add-in'et
 * kører som krop i en async-funktion med parametrene `context`, værtens objekt (`Excel`, `Word`
 * eller `PowerPoint`), `__key` og `__tick`.
 */
export function checkScript(source: string, host: ScriptHost = "Excel"): ScriptCheck {
  let ast: Node;
  try {
    ast = parse(source, {
      ecmaVersion: 2022,
      sourceType: "script",
      allowAwaitOutsideFunction: true,
      allowReturnOutsideFunction: true,
    }) as unknown as Node;
  } catch (e) {
    return { ok: false, error: `Syntaksfejl: ${(e as Error).message}` };
  }

  const inserts: { at: number; text: string }[] = [];
  const problems: string[] = [];
  const problem = (msg: string) => problems.length < 5 && !problems.includes(msg) && problems.push(msg);

  type Scope = { names: Set<string>; parent?: Scope; fn: boolean };
  // Parametrene, scriptet kaldes med: `context` og værtens eget objekt.
  const top: Scope = { names: new Set(["context", host]), fn: true };
  const resolves = (name: string, scope: Scope | undefined): boolean => {
    for (let s = scope; s; s = s.parent) if (s.names.has(name)) return true;
    return SCRIPT_GLOBALS.has(name);
  };

  /** Navne i et mønster (const {a, b: [c]} = …). Standardværdier og beregnede nøgler gennemgås. */
  const patternNames = (p: Node | null | undefined, out: string[], scope: Scope) => {
    if (!p) return;
    switch (p.type) {
      case "Identifier":
        out.push(p.name as string);
        break;
      case "ObjectPattern":
        for (const prop of p.properties as Node[]) {
          if (prop.type === "RestElement") patternNames(prop.argument as Node, out, scope);
          else {
            checkKey(prop, scope);
            patternNames(prop.value as Node, out, scope);
          }
        }
        break;
      case "ArrayPattern":
        for (const el of p.elements as (Node | null)[]) patternNames(el, out, scope);
        break;
      case "RestElement":
        patternNames(p.argument as Node, out, scope);
        break;
      case "AssignmentPattern":
        patternNames(p.left as Node, out, scope);
        visit(p.right as Node, scope);
        break;
      case "MemberExpression":
        // Kun i tildelinger ([a.b] = …), ikke i erklæringer.
        visit(p, scope);
        break;
    }
  };

  /** Nøglen i en objekt-egenskab: spærrede navne afvises, beregnede nøgler tjekkes ved kørsel. */
  const checkKey = (prop: Node, scope: Scope) => {
    const key = prop.key as Node;
    if (prop.computed) {
      if (key.type === "Literal") {
        if (isBannedKey(key.value)) problem(`egenskaben "${String(key.value)}"`);
      } else wrapKey(key);
      visit(key, scope);
    } else {
      const name = key.type === "Identifier" ? (key.name as string) : String(key.value);
      if (isBannedKey(name)) problem(`egenskaben "${name}"`);
    }
  };

  const wrapKey = (expr: Node) => {
    inserts.push({ at: expr.start, text: "__key(" }, { at: expr.end, text: ")" });
  };

  /** Erklæringer direkte i en blok (let/const/function) – og med var, i hele funktionen. */
  const hoist = (body: Node[], scope: Scope, fnScope: Scope) => {
    const varNames = (n: Node | null | undefined) => {
      if (!n || typeof n !== "object") return;
      if (n.type === "VariableDeclaration" && n.kind === "var") {
        for (const d of n.declarations as Node[]) {
          const names: string[] = [];
          patternNames(d.id as Node, names, fnScope);
          names.forEach((x) => fnScope.names.add(x));
        }
      }
      if (/Function/.test(n.type)) return; // var i indre funktioner hører til dem
      for (const v of Object.values(n)) {
        if (Array.isArray(v)) v.forEach((c) => varNames(c as Node));
        else if (v && typeof v === "object" && "type" in (v as object)) varNames(v as Node);
      }
    };
    for (const st of body) {
      if (st.type === "VariableDeclaration" && st.kind !== "var") {
        for (const d of st.declarations as Node[]) {
          const names: string[] = [];
          patternNames(d.id as Node, names, scope);
          names.forEach((x) => scope.names.add(x));
        }
      } else if (st.type === "FunctionDeclaration" && st.id) scope.names.add((st.id as Node).name as string);
      varNames(st);
    }
  };

  const tick = (body: Node) => {
    if (body.type === "BlockStatement") inserts.push({ at: body.start + 1, text: "__tick();" });
    else inserts.push({ at: body.start, text: "{__tick();" }, { at: body.end, text: "}" });
  };

  const fnScope = (fn: Node, scope: Scope) => {
    const inner: Scope = { names: new Set(), parent: scope, fn: true };
    if (fn.type === "FunctionExpression" && fn.id) inner.names.add((fn.id as Node).name as string);
    const params: string[] = [];
    for (const p of fn.params as Node[]) patternNames(p, params, inner);
    params.forEach((x) => inner.names.add(x));
    const body = fn.body as Node;
    if (body.type === "BlockStatement") {
      hoist(body.body as Node[], inner, inner);
      for (const st of body.body as Node[]) visit(st, inner);
    } else visit(body, inner);
  };

  function visit(n: Node | null | undefined, scope: Scope): void {
    if (!n || typeof n !== "object") return;
    if (FORBIDDEN_NODES[n.type]) {
      problem(FORBIDDEN_NODES[n.type]);
      return;
    }
    switch (n.type) {
      case "Identifier": {
        const name = n.name as string;
        if (name.startsWith("__")) problem(`navnet "${name}"`);
        else if (!resolves(name, scope)) problem(`"${name}" findes ikke i sandkassen`);
        return;
      }
      case "MemberExpression": {
        visit(n.object as Node, scope);
        const prop = n.property as Node;
        if (n.computed) {
          if (prop.type === "Literal") {
            if (isBannedKey(prop.value)) problem(`egenskaben "${String(prop.value)}"`);
          } else wrapKey(prop);
          visit(prop, scope);
        } else if (prop.type === "PrivateIdentifier" || isBannedKey(prop.name)) problem(`egenskaben "${String(prop.name)}"`);
        return;
      }
      case "ObjectExpression":
        for (const p of n.properties as Node[]) {
          if (p.type === "SpreadElement") visit(p.argument as Node, scope);
          else {
            checkKey(p, scope);
            visit(p.value as Node, scope);
          }
        }
        return;
      case "FunctionDeclaration":
      case "FunctionExpression":
      case "ArrowFunctionExpression":
        fnScope(n, scope);
        return;
      case "BlockStatement": {
        const inner: Scope = { names: new Set(), parent: scope, fn: false };
        hoist(n.body as Node[], inner, nearestFn(scope));
        for (const st of n.body as Node[]) visit(st, inner);
        return;
      }
      case "VariableDeclaration":
        for (const d of n.declarations as Node[]) {
          const names: string[] = [];
          patternNames(d.id as Node, names, scope);
          for (const x of names) if (x.startsWith("__")) problem(`navnet "${x}"`);
          visit(d.init as Node, scope);
        }
        return;
      case "ForStatement":
      case "ForInStatement":
      case "ForOfStatement": {
        const inner: Scope = { names: new Set(), parent: scope, fn: false };
        const init = (n.init ?? n.left) as Node | null;
        if (init?.type === "VariableDeclaration") hoist([init], inner, nearestFn(scope));
        visit(init, inner);
        visit(n.test as Node, inner);
        visit(n.update as Node, inner);
        visit(n.right as Node, inner);
        tick(n.body as Node);
        visit(n.body as Node, inner);
        return;
      }
      case "WhileStatement":
      case "DoWhileStatement":
        visit(n.test as Node, scope);
        tick(n.body as Node);
        visit(n.body as Node, scope);
        return;
      case "CatchClause": {
        const inner: Scope = { names: new Set(), parent: scope, fn: false };
        const names: string[] = [];
        patternNames(n.param as Node, names, inner);
        names.forEach((x) => inner.names.add(x));
        visit(n.body as Node, inner);
        return;
      }
      case "AssignmentExpression":
        if ((n.left as Node).type === "Identifier") visit(n.left as Node, scope);
        else {
          // Destrukturering til eksisterende variabler: hvert navn skal findes (ellers kunne et
          // globalt objekt blive overskrevet).
          const names: string[] = [];
          patternNames(n.left as Node, names, scope);
          for (const x of names) if (x.startsWith("__") || !resolves(x, scope)) problem(`"${x}" findes ikke i sandkassen`);
        }
        visit(n.right as Node, scope);
        return;
      case "LabeledStatement":
        visit(n.body as Node, scope);
        return;
      case "BreakStatement":
      case "ContinueStatement":
        return;
      case "Property":
      case "PropertyDefinition":
      case "MethodDefinition":
        return; // håndteres af ObjectExpression; klasser er afvist
    }
    for (const [k, v] of Object.entries(n)) {
      if (k === "type" || k === "start" || k === "end") continue;
      if (Array.isArray(v)) v.forEach((c) => visit(c as Node, scope));
      else if (v && typeof v === "object" && "type" in (v as object)) visit(v as Node, scope);
    }
  }
  const nearestFn = (scope: Scope): Scope => {
    let s: Scope | undefined = scope;
    while (s && !s.fn) s = s.parent;
    return s ?? top;
  };

  const program = ast.body as Node[];
  hoist(program, top, top);
  for (const st of program) visit(st, top);
  if (problems.length) return { ok: false, error: `Ikke tilladt i sandkassen: ${problems.join(", ")}` };

  let out = source;
  for (const ins of [...inserts].sort((a, b) => b.at - a.at)) out = out.slice(0, ins.at) + ins.text + out.slice(ins.at);
  return { ok: true, code: out };
}
