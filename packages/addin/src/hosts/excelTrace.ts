import { rangeBox, type TraceKind } from "@miai/shared";
import { splitAddress } from "../excelAddress";
import { colName } from "./excelSheet";

/**
 * Formelsporing (#76) som ren logik: gennemløb af Excels afhængighedsgraf med loft over dybde og
 * antal celler, beskæring og den kompakte liste, modellen får. Uden Office.js, så det kan testes
 * direkte; læsningen i arket ligger bag `TraceReader` (excelTraceReader.ts).
 *
 * Sporingen sker i koden og ikke i modellen: en lille model kan ikke selv følge henvisninger
 * gennem flere ark, men den kan læse en kort liste med adresse, formel, værdi og fejl.
 */

export interface TraceLimits {
  /** Højst så mange celler i listen. */
  maxCells: number;
  /** Højst så mange niveauer fra startcellen. */
  maxDepth: number;
  /** Højst så mange tegn i den færdige tekst. */
  maxChars: number;
}

/** Find fejlkilden og agentens sporinger. Teksten skal kunne være i en kontekst på 16.000 tokens sammen med resten. */
export const TRACE_LIMITS: TraceLimits = { maxCells: 200, maxDepth: 25, maxChars: 9_000 };
/** Forklar formel: kun de nærmeste celler, så forklaringen kan sige, hvor tallene kommer fra. */
export const EXPLAIN_LIMITS: TraceLimits = { maxCells: 40, maxDepth: 2, maxChars: 3_000 };
/** I et område med flere celler end dette vises fejlcellerne og kun så mange andre; resten tælles. */
export const AREA_SAMPLE = 12;
/** Af et meget stort område (fx en hel kolonne) undersøges højst så mange celler af den brugte del. */
export const MAX_AREA_READ = 2_000;

/** Et rektangulært område på ét ark, som Excel angiver det: uden $ og med store bogstaver, fx "B2:B13" eller "B:B". */
export interface TraceArea {
  sheet: string;
  range: string;
}

/** Én celle som start- eller mellemstation. */
export interface CellRef {
  sheet: string;
  /** Uden arknavn, fx "B5". */
  address: string;
}

/** Cellerne i et område, som de blev læst i arket. */
export interface AreaCells {
  /** Den læste bloks øverste venstre celle (0-baseret, som Office.js' rowIndex/columnIndex). */
  row: number;
  col: number;
  /** Formler som Office.js giver dem: "=…" for en formel, ellers værdien; "" for en tom celle. */
  formulas: unknown[][];
  values: unknown[][];
  /** Excel.RangeValueType pr. celle ("Error" for en fejlværdi). Uden den genkendes fejl på teksten. */
  types?: string[][];
  /** Antal celler i den brugte del af området, når den er større end det læste. */
  total?: number;
}

/**
 * Læsningen i arket. `direct` giver for hver celle de områder, den regnes ud fra (eller de områder,
 * der bruger den): [] når der ingen er, undefined når Excel ikke kunne svare. `read` henter
 * cellerne i områderne (undefined = kunne ikke læses). `whole` er hele grafen fra startcellen i ét
 * kald, hvor værten kan (ExcelApi 1.14 for præcedenter, 1.15 for afhængige).
 */
export interface TraceReader {
  direct(cells: CellRef[]): Promise<(TraceArea[] | undefined)[]>;
  read(areas: TraceArea[]): Promise<(AreaCells | undefined)[]>;
  whole?(cell: CellRef): Promise<TraceArea[] | undefined>;
}

export interface TraceCell extends CellRef {
  /** 0 = startcellen, 1 = de celler, dens formel bruger (eller som bruger den), osv. */
  depth: number;
  formula?: string;
  value: unknown;
  /** Fejlværdien, fx "#DIV/0!". */
  error?: string;
  empty?: boolean;
  /** Fejlen opstår her: cellen har en fejl, og ingen af de celler, dens formel bruger, har en. */
  origin?: boolean;
}

/** Et område, hvor ikke alle celler står i listen. */
export interface AreaNote extends TraceArea {
  depth: number;
  /** Celler i (den brugte del af) området. */
  cells: number;
  listed: number;
  /** Ikke viste celler uden fejl: formler og faste værdier. */
  formulas: number;
  constants: number;
  empty: number;
  /** Celler, der ikke blev undersøgt, fordi området er for stort. */
  unread: number;
}

export interface TraceGraph {
  kind: TraceKind;
  /** Celler og områdenoter i den rækkefølge, de blev fundet; den første er startcellen. */
  items: (({ type: "cell" } & TraceCell) | ({ type: "area" } & AreaNote))[];
  /** Startcellen har en fejl, så kun celler med fejl er fulgt videre (og det, kildens formel bruger). */
  followedErrors: boolean;
  /** Fundne celler, der ikke kom med, fordi loftet over antal celler var nået – og hvor mange af dem der har en fejl. */
  omitted: number;
  omittedErrors: number;
  /** Celler, der ikke blev sporet videre, fordi loftet over dybde eller antal celler var nået. */
  notFollowed: string[];
  /** Celler og områder, Excel ikke kunne svare på. */
  failed: string[];
  /** Hele grafens områder ifølge Excel – kun hentet, når grafen er beskåret, og værten kan. */
  whole?: TraceArea[];
}

/** Excels fejlværdier, som de står i en celles værdi. */
const ERROR_VALUE = /^#(?:DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|EXTERNAL!|PYTHON!|GETTING_DATA)$/;
const BOX = /^[A-Z]{1,3}[1-9]\d{0,6}(?::[A-Z]{1,3}[1-9]\d{0,6})?$/;

/** Områdets hjørner, når det er en celle eller et almindeligt område ("B2:D4") – ikke hele kolonner eller rækker ("B:B"). */
export const areaBox = (range: string) => (BOX.test(range) ? rangeBox(range) : undefined);

/** Arknavnet, som Excel skriver det i en adresse: i apostroffer, når det har andet end bogstaver, tal, _ og punktum. */
const quoteSheet = (name: string) => (/^[\p{L}\p{N}_.]+$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`);
/** "Data!B4" eller "'Mit ark'!B4" – med arknavn, så adressen er entydig og kan klikkes i svaret (#22). */
export const refText = (ref: { sheet: string; address: string }) => `${quoteSheet(ref.sheet)}!${ref.address}`;

const cellKey = (sheet: string, address: string) => `${sheet.toLowerCase()}!${address}`;

/** Den første celle i en adresse fra modellen eller markeringen: "Ark1!$B$2:C5" → { sheet: "Ark1", address: "B2" }. */
export function startCell(address: string, activeSheet: string): CellRef {
  const { sheet, range } = address.includes("!") ? splitAddress(address) : { sheet: activeSheet, range: address };
  return { sheet: sheet || activeSheet, address: range.replace(/\$/g, "").split(":")[0].toUpperCase() };
}

/**
 * Adresserne fra Excels WorkbookRangeAreas → områder. Hver tekst kan rumme flere områder adskilt
 * af komma (eller semikolon), hvert med sit arknavn: "Data!B2:B13, 'Mit ark'!$C$4". Skilletegn i
 * et arknavn i apostroffer deler ikke. Uden arknavn gælder `sheet`.
 */
export function parseAreas(addresses: string[], sheet: string): TraceArea[] {
  const out: TraceArea[] = [];
  const add = (text: string) => {
    const t = text.trim();
    if (!t) return;
    const s = t.includes("!") ? splitAddress(t) : { sheet, range: t };
    out.push({ sheet: s.sheet || sheet, range: s.range.replace(/\$/g, "").toUpperCase() });
  };
  for (const text of addresses) {
    let part = "";
    let quoted = false;
    for (const ch of text) {
      if (ch === "'") quoted = !quoted;
      if ((ch === "," || ch === ";") && !quoted) {
        add(part);
        part = "";
      } else part += ch;
    }
    add(part);
  }
  return out;
}

/** Cellerne i et læst område som celler i grafen (uden dybde). */
function cellsOf(area: TraceArea, data: AreaCells): Omit<TraceCell, "depth">[] {
  const out: Omit<TraceCell, "depth">[] = [];
  data.formulas.forEach((row, r) =>
    row.forEach((f, c) => {
      const value = data.values[r]?.[c] ?? "";
      const type = data.types?.[r]?.[c];
      const isError = type !== undefined ? type === "Error" : typeof value === "string" && ERROR_VALUE.test(value);
      const formula = typeof f === "string" && f.startsWith("=") ? f : undefined;
      out.push({
        sheet: area.sheet,
        address: `${colName(data.col + c)}${data.row + r + 1}`,
        value,
        ...(formula ? { formula } : {}),
        ...(isError ? { error: String(value) } : {}),
        ...(!formula && (value === "" || value === null || value === undefined) ? { empty: true } : {}),
      });
    }),
  );
  return out;
}

/**
 * Gennemløber grafen bredde først fra startcellen: niveau for niveau spørges Excel om de direkte
 * præcedenter (eller afhængige) for de celler, der skal følges, og cellerne i de nye områder læses.
 *
 * - Celler ses kun én gang, så en cirkulær henvisning stopper af sig selv.
 * - Har startcellen en fejl, følges kun celler med fejl: en celle uden fejl kan ikke være der,
 *   fejlen kommer fra. Den celle, hvis formel kun bruger fejlfri celler, er kilden (`origin`), og
 *   det, dens formel bruger, følges ét niveau mere, så man kan se hvorfor (fx at nævneren er 0).
 * - I et stort område vises fejlcellerne og de første AREA_SAMPLE andre; resten tælles i en note.
 * - Er loftet over celler nået, beholdes fejl før formler før faste værdier, og resten tælles.
 */
export async function traceGraph(start: CellRef, kind: TraceKind, reader: TraceReader, limits: TraceLimits = TRACE_LIMITS): Promise<TraceGraph> {
  const [first] = await reader.read([{ sheet: start.sheet, range: start.address }]);
  const root = first && cellsOf({ sheet: start.sheet, range: start.address }, first)[0];
  // Beskeden går til modellen (agentens læsning), derfor på engelsk som resten af det, den får.
  if (!root) throw new Error(`the cell ${refText(start)} could not be read`);
  const graph: TraceGraph = {
    kind,
    items: [{ type: "cell", ...root, depth: 0 }],
    followedErrors: kind === "precedents" && !!root.error,
    omitted: 0,
    omittedErrors: 0,
    notFollowed: [],
    failed: [],
  };
  const precedents = kind === "precedents";
  if (precedents && root.error && !root.formula) (graph.items[0] as TraceCell).origin = true;
  const seen = new Set([cellKey(root.sheet, root.address)]);
  /** Områder, der er læst: om de har en fejl, og om hele området er undersøgt. */
  const areas = new Map<string, { error: boolean; complete: boolean }>();
  let listed = 1;
  // Afhængige: enhver celle kan have nogen. Præcedenter: kun en formel har nogen.
  let frontier: TraceCell[] = !precedents || root.formula ? [graph.items[0] as TraceCell] : [];

  for (let depth = 1; frontier.length; depth++) {
    if (depth > limits.maxDepth || listed >= limits.maxCells) {
      graph.notFollowed = frontier.map(refText);
      break;
    }
    const links = await reader.direct(frontier);
    const fresh: TraceArea[] = [];
    const key = (a: TraceArea) => cellKey(a.sheet, a.range);
    for (const list of links) for (const a of list ?? []) if (!areas.has(key(a)) && !fresh.some((x) => key(x) === key(a))) fresh.push(a);
    const data = fresh.length ? await reader.read(fresh) : [];

    type Found = Omit<TraceCell, "depth"> & { area: string };
    const found: (Found | ({ note: string } & AreaNote))[] = [];
    fresh.forEach((area, i) => {
      const d = data[i];
      if (!d) {
        areas.set(key(area), { error: false, complete: false });
        graph.failed.push(refText({ sheet: area.sheet, address: area.range }));
        return;
      }
      const all = cellsOf(area, d);
      const unread = Math.max(0, (d.total ?? all.length) - all.length);
      areas.set(key(area), { error: all.some((c) => c.error), complete: unread === 0 });
      const news = all.filter((c) => !seen.has(cellKey(c.sheet, c.address)));
      // Et lille område vises helt, også de tomme celler (en tom nævner giver #DIV/0!).
      const small = all.length <= AREA_SAMPLE && unread === 0;
      let others = 0;
      const shown = news.filter((c) => small || c.error || (!c.empty && ++others <= AREA_SAMPLE));
      const hidden = news.filter((c) => !shown.includes(c));
      if (hidden.length || unread)
        found.push({
          note: key(area),
          ...area,
          depth,
          cells: d.total ?? all.length,
          listed: shown.length,
          formulas: hidden.filter((c) => c.formula).length,
          constants: hidden.filter((c) => !c.formula && !c.empty).length,
          empty: hidden.filter((c) => c.empty).length,
          unread,
        });
      for (const c of shown) {
        seen.add(cellKey(c.sheet, c.address));
        found.push({ ...c, area: key(area) });
      }
    });

    // Kilden: en fejlcelle, hvis formel kun bruger celler uden fejl (og hvor alle er undersøgt).
    const originAreas = new Set<string>();
    frontier.forEach((cell, i) => {
      const list = links[i];
      if (!list) {
        graph.failed.push(refText(cell));
        return;
      }
      if (!precedents || !cell.error) return;
      const known = list.map((a) => areas.get(key(a)));
      if (known.every((a) => a && a.complete && !a.error)) {
        cell.origin = true;
        for (const a of list) originAreas.add(key(a));
      }
    });

    // Loftet over antal celler: fejl først, så formler, så faste værdier – i øvrigt i den fundne rækkefølge.
    const cells = found.filter((x): x is Found => !("note" in x));
    const rank = (c: Found) => (c.error ? 0 : c.formula ? 1 : c.empty ? 3 : 2);
    const room = limits.maxCells - listed;
    const kept = new Set(
      cells
        .map((c, i) => ({ c, i }))
        .sort((a, b) => rank(a.c) - rank(b.c) || a.i - b.i)
        .slice(0, room)
        .map((x) => x.c),
    );
    const dropped = cells.filter((c) => !kept.has(c));
    graph.omitted += dropped.length;
    graph.omittedErrors += dropped.filter((c) => c.error).length;

    const next: TraceCell[] = [];
    for (const x of found) {
      if ("note" in x) {
        // "Vist" er det, der faktisk kom med i listen – også når loftet over celler har skåret i området.
        const { note, ...area } = x;
        graph.items.push({ type: "area", ...area, listed: cells.filter((c) => c.area === note && kept.has(c)).length });
        continue;
      }
      if (!kept.has(x)) continue;
      const { area, ...rest } = x;
      const cell: TraceCell = { ...rest, depth, ...(precedents && rest.error && !rest.formula ? { origin: true } : {}) };
      graph.items.push({ type: "cell", ...cell });
      listed++;
      // items rummer en kopi; den celle, der følges videre, skal være den samme, så `origin` kan sættes på den.
      const item = graph.items[graph.items.length - 1] as TraceCell;
      const follow = precedents ? !!item.formula && (!graph.followedErrors || !!item.error || originAreas.has(area)) : !item.empty;
      if (follow) next.push(item);
    }
    frontier = next;
  }

  const cut = graph.omitted > 0 || graph.notFollowed.length > 0;
  if (cut && reader.whole) graph.whole = await reader.whole(start).catch(() => undefined);
  return graph;
}

const clip = (s: string, max: number) => (s.length > max ? `${s.slice(0, max)}…` : s);

/** En værdi, som den står i listen: tal som tal, tekst i anførselstegn, tom celle som "(empty)". */
function valueText(c: Pick<TraceCell, "value" | "empty" | "error">): string {
  if (c.error) return c.error;
  if (c.empty) return "(empty)";
  if (typeof c.value === "number") return String(c.value);
  if (typeof c.value === "boolean") return c.value ? "TRUE" : "FALSE";
  return `"${clip(String(c.value).replace(/\s+/g, " "), 80)}"`;
}

const plural = (n: number, one: string, many: string) => `${n} ${n === 1 ? one : many}`;

/** Antal celler i områderne, når alle er almindelige områder (hele kolonner har intet fast antal). */
function countCells(list: TraceArea[]): number | undefined {
  let n = 0;
  for (const a of list) {
    const b = areaBox(a.range);
    if (!b) return undefined;
    n += (b.r1 - b.r0 + 1) * (b.c1 - b.c0 + 1);
  }
  return n;
}

/**
 * Grafen som den kompakte tekst, modellen får: én linje pr. celle med niveau, adresse (med
 * arknavn), formel og værdi, og fejl og kilden markeret. Alt, der er udeladt, står med ordet
 * TRUNCATED, så modellen ikke tager listen for hele grafen.
 */
export function renderTrace(g: TraceGraph, limits: TraceLimits = TRACE_LIMITS): string {
  const start = g.items[0] as TraceCell;
  const precedents = g.kind === "precedents";
  const head = precedents
    ? `Precedents of ${refText(start)}: the cells its formula is calculated from, traced by the add-in in Excel's dependency graph. L0 is the cell itself, L1 the cells its formula uses, L2 the cells those use, and so on.`
    : `Dependents of ${refText(start)}: the cells whose formulas use it, traced by the add-in in Excel's dependency graph. L0 is the cell itself, L1 the cells that use it directly, L2 the cells that use those, and so on.`;
  const lines = [
    head,
    "Each line: address: formula → value, or address: value for a cell without a formula. A reference without a sheet name inside a formula is on the same sheet as that formula's cell.",
  ];
  const tail: string[] = [];
  const cells = g.items.filter((x) => x.type === "cell");
  if (cells.length === 1 && !g.failed.length && !g.notFollowed.length)
    tail.push(
      !precedents
        ? "No formula uses this cell."
        : start.formula
          ? "The cell's formula uses no other cells."
          : "The cell holds a value, not a formula, so nothing is traced.",
    );
  if (g.followedErrors) {
    tail.push(
      cells.some((c) => c.origin)
        ? "Only cells with an error value were traced further, and the cells used by the ORIGIN formula: a cell without an error is not where the error comes from."
        : "Only cells with an error value were traced further. The add-in could not determine in which cell the error starts.",
    );
  }
  if (g.omitted)
    tail.push(
      `TRUNCATED: the list stops at ${limits.maxCells} cells. ${plural(g.omitted, "more cell was", "more cells were")} found but ${g.omitted === 1 ? "is" : "are"} not listed${g.omittedErrors ? ` (${g.omittedErrors} of them with an error)` : ""}.`,
    );
  if (g.notFollowed.length)
    tail.push(
      `TRUNCATED: the trace stopped here (limit: ${limits.maxDepth} levels, ${limits.maxCells} cells). Not traced further: ${g.notFollowed.slice(0, 8).join(", ")}${g.notFollowed.length > 8 ? ` and ${g.notFollowed.length - 8} more` : ""}. ${precedents ? "The cells they are calculated from" : "The cells that use them"} are not in the list.`,
    );
  if (g.failed.length)
    tail.push(`Excel could not trace or read: ${g.failed.slice(0, 8).join(", ")}${g.failed.length > 8 ? ` and ${g.failed.length - 8} more` : ""}.`);
  if (g.whole?.length) {
    const n = countCells(g.whole);
    const names = g.whole.map((a) => refText({ sheet: a.sheet, address: a.range }));
    tail.push(
      `According to Excel, the whole graph of ${g.kind} covers${n === undefined ? "" : ` ${plural(n, "cell", "cells")} in`}: ${names.slice(0, 12).join(", ")}${names.length > 12 ? ` and ${names.length - 12} more ranges` : ""}.`,
    );
  }

  const room = limits.maxChars - lines.join("\n").length - tail.join("\n").length - 120;
  let used = 0;
  let cutLines = 0;
  for (const x of g.items) {
    let line: string;
    if (x.type === "area") {
      const rest = [
        x.formulas ? plural(x.formulas, "formula", "formulas") : "",
        x.constants ? plural(x.constants, "value", "values") : "",
        x.empty ? plural(x.empty, "empty cell", "empty cells") : "",
      ].filter(Boolean);
      line =
        `L${x.depth} ${refText({ sheet: x.sheet, address: x.range })}: a range of ${plural(x.cells, "cell", "cells")}; ${x.listed} listed below` +
        (rest.length ? `, not listed (none of them has an error): ${rest.join(", ")}` : "") +
        (x.unread ? `. TRUNCATED: ${plural(x.unread, "cell", "cells")} of the range were not examined` : "") +
        ".";
    } else {
      const marks = [x.error ? "error" : "", x.origin ? "ORIGIN: the error starts in this cell" : ""].filter(Boolean);
      line =
        `L${x.depth} ${refText(x)}: ` +
        (x.formula ? `${clip(x.formula, 300)} → ${valueText(x)}` : valueText(x)) +
        (marks.length ? ` [${marks.join("; ")}]` : "");
    }
    if (used + line.length + 1 > room) {
      cutLines++;
      continue;
    }
    used += line.length + 1;
    lines.push(line);
  }
  if (cutLines) tail.push(`TRUNCATED: ${plural(cutLines, "line was", "lines were")} left out to keep the list short.`);
  return [...lines, ...tail].join("\n");
}
