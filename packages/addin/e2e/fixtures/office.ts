/**
 * Simuleret Office.js til browsertests (#10).
 *
 * `installFakeOffice` serialiseres med toString() og køres i siden i stedet for Microsofts
 * office.js. Den skal derfor være selvstændig (ingen lukninger over modulets variabler).
 *
 * Alle objekter er "strenge": læser adapterne en egenskab, som simuleringen ikke har,
 * kastes "Simuleret Office.js mangler: …". Så opdages det med det samme, når adapterne
 * begynder at bruge nye API'er, i stedet for at testen fejler et uventet sted.
 *
 * Alt, hvad adapterne gør ved dokumentet, registreres i `window.__miai.calls`.
 */

export type HostName = "Word" | "Excel" | "PowerPoint" | "Outlook";

export interface FakeOfficeOptions {
  host: HostName;
  /** Højeste understøttede version pr. requirement set, fx { WordApi: "1.4" }. Manglende = ikke understøttet. */
  apiSets?: Record<string, string>;
  /** Markeret tekst (Word, PowerPoint). */
  selection?: string;
  /** Word: hele dokumentets tekst (bruges, når intet er markeret). Standard: markeringen. */
  document?: string;
  /**
   * Word med dokumentmodel: afsnit med typografi og formatering (agent-tilstand). Når den er
   * sat, kan dokumentets tilstand efter handlinger læses med window.__miai.doc().
   */
  paragraphs?: {
    text: string;
    style?: string;
    list?: boolean;
    table?: boolean;
    bold?: boolean;
    marks?: { start: number; end: number; bold?: boolean; italic?: boolean; link?: string }[];
  }[];
  /** Excel: arkets værdier (række 1 = A1). Markeringen er hele området. Formater og tabeller er valgfri. */
  excel?: {
    sheet: string;
    values: (string | number)[][];
    /** Fx { "B2:B4": { numberFormat: "#,##0 kr" }, "A1:C1": { bold: true } }. */
    formats?: Record<string, { numberFormat?: string; bold?: boolean; fill?: string }>;
    /** Excel-tabeller, fx [{ name: "Salg", range: "A1:C4" }]. */
    tables?: { name: string; range: string }[];
    /** Flere ark i projektmappen (#21); arket ovenfor er det aktive. */
    sheets?: {
      sheet: string;
      values?: (string | number)[][];
      formats?: Record<string, { numberFormat?: string; bold?: boolean; fill?: string }>;
      tables?: { name: string; range: string }[];
    }[];
  };
  displayLanguage?: string;
  /** Outlook: mailen i læsetilstand (read) eller et udkast i skrivetilstand (compose). */
  mail?: { mode: "read" | "compose"; subject: string; from?: { name: string; email: string }; body: string; selection?: string; html?: boolean };
  /** PowerPoint med præsentationsmodel (agent-tilstand): slides med layout og figurer. */
  slides?: {
    layout: string;
    shapes: {
      name: string;
      type?: string;
      ph?: string;
      text: string;
      style?: Record<string, unknown>;
      marks?: { start: number; end: number; bold?: boolean; italic?: boolean }[];
    }[];
  }[];
  /** Præsentationens layouts og deres pladsholdere (til nye slides). */
  layouts?: { name: string; placeholders: string[] }[];
  /** Office' tema. Kan skiftes undervejs med window.__miai.setTheme("dark"). */
  theme?: "light" | "dark";
}

export interface OfficeCall {
  kind: string;
  [key: string]: unknown;
}

export function installFakeOffice(o: FakeOfficeOptions): void {
  const calls: OfficeCall[] = [];
  const w = window as unknown as Record<string, unknown>;
  // Office' temafarver (lyst: hvid; mørkt: som Office' "Sort"/"Mørkegrå").
  const THEMES = {
    light: {
      bodyBackgroundColor: "#ffffff",
      bodyForegroundColor: "#262626",
      controlBackgroundColor: "#ffffff",
      controlForegroundColor: "#262626",
      isDarkTheme: false,
    },
    dark: {
      bodyBackgroundColor: "#262626",
      bodyForegroundColor: "#ffffff",
      controlBackgroundColor: "#333333",
      controlForegroundColor: "#ffffff",
      isDarkTheme: true,
    },
  };
  const officeTheme = { ...THEMES[o.theme ?? "light"] };
  const themeHandlers: ((e: unknown) => void)[] = [];
  w.__miai = {
    calls,
    setTheme(name: "light" | "dark") {
      Object.assign(officeTheme, THEMES[name]);
      for (const h of themeHandlers) h({ type: "officeThemeChanged", officeTheme });
    },
  };

  const PASS = new Set(["then", "toJSON", "constructor", "$$typeof", "nodeType", "asymmetricMatch"]);
  function strict<T extends object>(name: string, obj: T): T {
    return new Proxy(obj, {
      get(t, p, r) {
        if (typeof p === "symbol" || p in t || PASS.has(p)) return Reflect.get(t, p, r);
        throw new Error(`Simuleret Office.js mangler: ${name}.${p}`);
      },
    });
  }
  const record = (c: OfficeCall) => {
    calls.push(c);
    return c;
  };
  const sync = async () => {};

  function versionAtLeast(have: string | undefined, want: string): boolean {
    if (!have) return false;
    const a = have.split(".").map(Number);
    const b = want.split(".").map(Number);
    for (let i = 0; i < Math.max(a.length, b.length); i++) {
      if ((a[i] ?? 0) !== (b[i] ?? 0)) return (a[i] ?? 0) > (b[i] ?? 0);
    }
    return true;
  }

  // --- Outlook: mailen i læse- eller skrivetilstand (#7) ------------------------
  // Ikke en "strict" proxy: adapteren afgør tilstanden med typeof item.displayReplyForm.
  const ok = (cb: (r: unknown) => void, value?: unknown) => cb({ status: "succeeded", value });
  let mailItem: Record<string, unknown> | undefined;
  if (o.host === "Outlook") {
    const m = o.mail ?? { mode: "read", subject: "", body: "" };
    let body = m.body;
    let sel = m.selection ?? "";
    mailItem =
      m.mode === "read"
        ? {
            itemType: "message",
            subject: m.subject,
            from: m.from ? { displayName: m.from.name, emailAddress: m.from.email } : undefined,
            body: { getAsync: (_t: string, cb: (r: unknown) => void) => ok(cb, body) },
            displayReplyForm: (f: { htmlBody: string }) => record({ kind: "replyForm", all: false, html: f.htmlBody }),
            displayReplyAllForm: (f: { htmlBody: string }) => record({ kind: "replyForm", all: true, html: f.htmlBody }),
          }
        : {
            itemType: "message",
            displayReplyForm: undefined,
            subject: { getAsync: (cb: (r: unknown) => void) => ok(cb, m.subject) },
            getSelectedDataAsync: (type: string, cb: (r: unknown) => void) => ok(cb, { data: type === "html" ? `<b>${sel}</b>` : sel }),
            body: {
              getAsync: (_t: string, cb: (r: unknown) => void) => ok(cb, body),
              getTypeAsync: (cb: (r: unknown) => void) => ok(cb, m.html === false ? "text" : "html"),
              setAsync: (v: string, opts: { coercionType: string }, cb: (r: unknown) => void) => {
                record({ kind: "setBody", value: v, coercionType: opts.coercionType });
                body = v;
                ok(cb);
              },
              setSelectedDataAsync: (v: string, opts: { coercionType: string }, cb: (r: unknown) => void) => {
                record({ kind: "setSelectedData", text: v, coercionType: opts.coercionType });
                sel = "";
                ok(cb);
              },
            },
          };
  }

  // --- Fælles Office-objekt -------------------------------------------------
  let selection = o.selection ?? "";
  w.Office = strict("Office", {
    HostType: { Word: "Word", Excel: "Excel", PowerPoint: "PowerPoint", Outlook: "Outlook" },
    CoercionType: { Text: "text", Html: "html" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    GoToType: { Index: "Index" },
    EventType: { OfficeThemeChanged: "officeThemeChanged" },
    context: strict("Office.context", {
      displayLanguage: o.displayLanguage ?? "da-DK",
      officeTheme,
      requirements: { isSetSupported: (name: string, v = "1.1") => versionAtLeast(o.apiSets?.[name], v) },
      mailbox: { item: mailItem },
      document: strict("Office.context.document", {
        goToByIdAsync: (id: number, type: string, cb: (r: unknown) => void) => {
          record({ kind: "goTo", id, type });
          cb({ status: "succeeded" });
        },
        getSelectedDataAsync: (_t: string, cb: (r: unknown) => void) => cb({ status: "succeeded", value: selection }),
        addHandlerAsync: (type: string, handler: (e: unknown) => void, cb?: (r: unknown) => void) => {
          if (type === "officeThemeChanged") themeHandlers.push(handler);
          cb?.({ status: "succeeded" });
        },
        setSelectedDataAsync: (text: string, _opts: unknown, cb: (r: unknown) => void) => {
          record({ kind: "setSelectedData", text });
          selection = text;
          cb({ status: "succeeded" });
        },
      }),
    }),
    onReady(cb: (info: { host: string }) => void) {
      setTimeout(() => cb({ host: o.host }), 0);
      return Promise.resolve({ host: o.host });
    },
  });

  // --- Word med dokumentmodel (agent-tilstand, #6) ---------------------------------
  if (o.host === "Word" && o.paragraphs) {
    installWordModel();
  }

  /**
   * Word med afsnit, typografi, inline-formatering, kommentarer og sporede ændringer.
   * Efterligner Words egne regler, også dem agenten skal arbejde uden om:
   *  - et nyt afsnit arver naboafsnittets typografi, liste, tabel og direkte formatering;
   *  - erstattet tekst får formateringen fra første tegn i det erstattede.
   * Dokumentets tilstand kan læses med window.__miai.doc().
   */
  function installWordModel() {
    interface Mark {
      start: number;
      end: number;
      bold?: boolean;
      italic?: boolean;
      link?: string;
    }
    interface P {
      text: string;
      style: string;
      list: boolean;
      table: boolean;
      /** Direkte formatering på hele afsnittet (arves ved indsættelse). */
      bold: boolean;
      marks: Mark[];
    }
    const listStyles = new Set(["ListBullet", "ListNumber"]);
    let paras: P[] = o.paragraphs!.map((p) => ({
      text: p.text,
      style: p.style ?? "Normal",
      list: !!p.list || listStyles.has(p.style ?? ""),
      table: !!p.table,
      bold: !!p.bold,
      marks: (p.marks ?? []).map((m) => ({ ...m })),
    }));
    const clone = (ps: P[]) => ps.map((p) => ({ ...p, marks: p.marks.map((m) => ({ ...m })) }));
    const state = { tracking: "Off" };
    const tracked: { date: string; before: P[] }[] = [];
    /** Den tidligste afviste ændring – tilstanden før den gendannes ved næste sync. */
    let restoreIdx: number | undefined;
    const comments: { id: string; on: string; text: string }[] = [];
    /** Kaldes før hver ændring: med sporede ændringer registreres en ændring, der kan afvises. */
    const change = (kind: string, detail: Record<string, unknown>) => {
      record({ kind, ...detail });
      if (state.tracking !== "Off") tracked.push({ date: new Date().toISOString(), before: clone(paras) });
    };
    (w.__miai as Record<string, unknown>).doc = () => clone(paras).map((p) => ({ ...p, marks: p.marks }));
    (w.__miai as Record<string, unknown>).comments = () => comments;
    /** Brugeren redigerer et afsnit (fx mellem plan og udførelse). */
    (w.__miai as Record<string, unknown>).editParagraph = (i: number, text: string) => {
      paras[i].text = text;
    };

    /** Erstat [start,end) i et afsnit; markeringer følger Words regel om første tegn. */
    function replaceIn(p: P, start: number, end: number, rep: string) {
      const delta = rep.length - (end - start);
      const first = p.marks.find((m) => m.start <= start && m.end > start);
      p.marks = p.marks
        .filter((m) => m.end <= start || m.start >= end || m === first)
        .map((m) => {
          if (m === first) return { ...m, end: m.end >= end ? m.end + delta : start + rep.length };
          return m.start >= end ? { ...m, start: m.start + delta, end: m.end + delta } : m;
        })
        .filter((m) => m.end > m.start);
      p.text = p.text.slice(0, start) + rep + p.text.slice(end);
    }

    function paragraphObj(p: P): Record<string, unknown> {
      const obj: Record<string, unknown> = {
        load() {},
        get text() {
          return p.text;
        },
        get style() {
          return p.style;
        },
        get styleBuiltIn() {
          return p.style;
        },
        set styleBuiltIn(v: string) {
          change("style", { text: p.text, style: v });
          p.style = v;
          if (listStyles.has(v)) p.list = true;
        },
        get isListItem() {
          return p.list;
        },
        get tableNestingLevel() {
          return p.table ? 1 : 0;
        },
        detachFromList() {
          change("detachFromList", { text: p.text });
          p.list = false;
        },
        insertParagraph(text: string, loc: string) {
          // Som Word: det nye afsnit arver naboens typografi, liste, tabel og direkte formatering.
          const np: P = { text, style: p.style, list: p.list, table: p.table, bold: p.bold, marks: [] };
          change("insertParagraph", { text, loc, anchor: p.text, inheritedStyle: p.style });
          const i = paras.indexOf(p);
          paras.splice(loc === "Before" ? i : i + 1, 0, np);
          return strict("Word.Paragraph", paragraphObj(np));
        },
        delete() {
          change("deleteParagraph", { text: p.text });
          paras = paras.filter((x) => x !== p);
        },
        select() {
          record({ kind: "select", paragraph: paras.indexOf(p) + 1, text: p.text });
        },
        getRange: () => wholeParagraph(p),
        search: (t: string) => search([p], t),
      };
      return obj;
    }

    function wholeParagraph(p: P): unknown {
      return strict("Word.Range(paragraph)", {
        font: strict("Word.Font", {
          set(f: { bold?: boolean; italic?: boolean }) {
            change("font", { text: p.text, ...f });
            if (f.bold === false) {
              p.bold = false;
              p.marks = p.marks.filter((m) => !m.bold);
            }
            if (f.italic === false) p.marks = p.marks.filter((m) => !m.italic);
          },
        }),
        getOoxml: () => ({ value: JSON.stringify(p) }),
        insertOoxml(xml: string, loc: string) {
          const np = JSON.parse(xml) as P;
          change("insertOoxml", { text: np.text, loc, anchor: p.text });
          const i = paras.indexOf(p);
          paras.splice(loc === "Before" ? i : i + 1, 0, np);
        },
      });
    }

    /** Et søgeresultat: tekst i ét afsnit fra `start`. */
    function hit(p: P, start: number, len: number): unknown {
      const get = () => p.text.slice(start, start + len);
      return strict("Word.Range(search)", {
        load() {},
        select() {
          record({ kind: "select", text: get() });
        },
        get text() {
          return get();
        },
        insertText(rep: string, loc: string) {
          change("replace", { from: get(), to: rep, loc });
          replaceIn(p, start, start + len, rep);
          len = rep.length;
        },
        delete() {
          change("delete", { text: get() });
          replaceIn(p, start, start + len, "");
          len = 0;
        },
        insertComment(text: string) {
          const c = { id: `c${comments.length + 1}`, on: get(), text };
          record({ kind: "comment", on: c.on, text });
          comments.push(c);
          return strict("Word.Comment", { id: c.id, load() {} });
        },
        search: (t: string) => {
          const items: unknown[] = [];
          for (let i = get().indexOf(t); i >= 0; i = get().indexOf(t, i + t.length)) items.push(hit(p, start + i, t.length));
          return { items, load() {} };
        },
        paragraphs: { getFirst: () => strict("Word.Paragraph", paragraphObj(p)) },
        font: {
          set bold(v: boolean) {
            change("bold", { text: get() });
            if (v) p.marks.push({ start, end: start + len, bold: true });
          },
          set italic(v: boolean) {
            change("italic", { text: get() });
            if (v) p.marks.push({ start, end: start + len, italic: true });
          },
        },
        set hyperlink(url: string) {
          change("hyperlink", { text: get(), url });
          p.marks.push({ start, end: start + len, link: url });
        },
      });
    }

    function search(ps: P[], t: string) {
      const items: unknown[] = [];
      for (const p of ps) for (let i = p.text.indexOf(t); i >= 0; i = p.text.indexOf(t, i + t.length)) items.push(hit(p, i, t.length));
      return { items, load() {} };
    }

    const whole = strict("Word.Range(body)", {
      load() {},
      get text() {
        return paras.map((p) => p.text).join("\n");
      },
      get paragraphs() {
        const items = paras.map((p) => strict("Word.Paragraph", paragraphObj(p)));
        return { items, load() {}, getFirst: () => items[0], getLast: () => items[items.length - 1] };
      },
      search: (t: string) => search(paras, t),
    });
    const empty = strict("Word.Range(selection)", { text: "", load() {} });
    const body = strict("Word.Body", {
      getRange: () => whole,
      insertParagraph(text: string, loc: string) {
        const ref = loc === "Start" ? paras[0] : paras[paras.length - 1];
        return (paragraphObj(ref).insertParagraph as (t: string, l: string) => unknown)(text, loc === "Start" ? "Before" : "After");
      },
      getOoxml: () => ({ value: JSON.stringify(paras) }),
      insertOoxml(xml: string, loc: string) {
        record({ kind: "insertOoxml", loc });
        paras = JSON.parse(xml);
      },
      getTrackedChanges: () => ({
        items: tracked.map((t) =>
          strict("Word.TrackedChange", {
            date: t.date,
            reject() {
              record({ kind: "rejectTrackedChange" });
              restoreIdx = Math.min(restoreIdx ?? Infinity, tracked.indexOf(t));
            },
          }),
        ),
        load() {},
      }),
      getComments: () => ({
        items: comments.map((c) =>
          strict("Word.Comment", {
            id: c.id,
            delete() {
              record({ kind: "deleteComment", id: c.id });
              comments.splice(comments.indexOf(c), 1);
            },
          }),
        ),
        load() {},
      }),
    });
    const document = strict("Word.Document", {
      getSelection: () => empty,
      body,
      load() {},
      get changeTrackingMode() {
        return state.tracking;
      },
      set changeTrackingMode(v: string) {
        state.tracking = v;
      },
    });
    const ctx = strict("Word.RequestContext", {
      document,
      trackedObjects: { add() {}, remove() {} },
      async sync() {
        if (restoreIdx !== undefined) {
          paras = clone(tracked[restoreIdx].before);
          tracked.splice(restoreIdx);
          restoreIdx = undefined;
        }
      },
    });
    w.Word = strict("Word", {
      ChangeTrackingMode: { trackAll: "TrackAll", off: "Off" },
      run: (a: unknown, b?: unknown) => (typeof a === "function" ? a(ctx) : (b as (c: unknown, r: unknown) => unknown)(ctx, a)),
    });
  }

  // --- Word -------------------------------------------------------------------
  if (o.host === "Word" && !o.paragraphs) {
    const picture = (base64: string, loc: string) => record({ kind: "picture", bytes: base64.length, loc });
    // Afsnit: typografi registreres, når agenten sætter den (#6).
    const paragraph = (text = "") => {
      const p: Record<string, unknown> = {
        insertInlinePictureFromBase64: picture,
        insertHtml: (html: string, loc: string) => record({ kind: "insertHtml", html, loc }),
        insertParagraph: (t: string, loc: string) => (record({ kind: "insertParagraph", text: t, loc }), paragraph(t)),
      };
      Object.defineProperty(p, "styleBuiltIn", { set: (v: string) => record({ kind: "style", text, style: v }), enumerable: true });
      return strict("Word.Paragraph", p);
    };
    // Et søgeresultat: ét sted i teksten, hvor adapterne kan erstatte, slette, kommentere eller indsætte.
    const hit = (text: string): unknown =>
      strict("Word.Range(search)", {
        text,
        load() {},
        insertText: (rep: string, loc: string) => record({ kind: loc === "Replace" ? "replace" : "insertText", from: text, to: rep, text: rep, loc }),
        delete: () => record({ kind: "delete", text }),
        insertComment: (c: string) => record({ kind: "comment", on: text, text: c }),
        insertParagraph: (t: string, loc: string) => (record({ kind: "insertParagraph", text: t, loc, anchor: text }), paragraph(t)),
        paragraphs: {
          getLast: () =>
            strict("Word.Paragraph(anchor)", {
              insertParagraph: (t: string, loc: string) => (record({ kind: "insertParagraph", text: t, loc, anchor: text }), paragraph(t)),
            }),
        },
      });
    const makeRange = (getText: () => string): Record<string, unknown> => {
      const r: Record<string, unknown> = strict("Word.Range", {
        get text() {
          return getText();
        },
        load() {},
        insertText: (text: string, loc: string) => record({ kind: "insertText", text, loc }),
        insertHtml: (html: string, loc: string) => record({ kind: "insertHtml", html, loc }),
        insertComment: (text: string) => record({ kind: "comment", text }),
        insertParagraph: (t: string, loc: string) => (record({ kind: "insertParagraph", text: t, loc }), paragraph(t)),
        insertInlinePictureFromBase64: picture,
        getRange: () => r,
        // Afsnit (én pr. linje) til agentens struktur; dokumentmodellen ovenfor bruges til agent-tests.
        get paragraphs() {
          const items = getText()
            .split("\n")
            .map((t) =>
              strict("Word.Paragraph(simple)", { text: t, style: "Normal", styleBuiltIn: "Normal", isListItem: false, tableNestingLevel: 0, load() {} }),
            );
          return { items, load() {} };
        },
        search(text: string) {
          // Én forekomst pr. gang teksten findes.
          const items: unknown[] = [];
          const all = getText();
          let i = all.indexOf(text);
          while (i >= 0) {
            items.push(hit(text));
            i = all.indexOf(text, i + text.length);
          }
          return { items, load() {} };
        },
      });
      return r;
    };
    const range = makeRange(() => selection);
    const bodyRange = makeRange(() => o.document ?? selection);
    const document = strict("Word.Document", {
      getSelection: () => range,
      body: strict("Word.Body", {
        getRange: () => (o.document !== undefined ? bodyRange : range),
        insertParagraph: (t: string, loc: string) => (record({ kind: "insertParagraph", text: t, loc }), paragraph(t)),
        getOoxml: () => ({ value: "<pkg:package>øjebliksbillede</pkg:package>" }),
        insertOoxml: (ooxml: string, loc: string) => record({ kind: "insertOoxml", ooxml, loc }),
      }),
      load() {},
      changeTrackingMode: "Off",
    });
    const ctx = strict("Word.RequestContext", { document, trackedObjects: { add() {}, remove() {} }, sync });
    w.Word = strict("Word", {
      ChangeTrackingMode: { trackAll: "TrackAll", off: "Off" },
      run: (a: unknown, b?: unknown) => (typeof a === "function" ? a(ctx) : (b as (c: unknown, r: unknown) => unknown)(ctx, a)),
    });
  }

  // --- Excel ------------------------------------------------------------------
  // Et ark med værdier, formater (talformat, fed, fyld) og tabeller, der ændres af det,
  // adapterne skriver. Tilstanden kan læses med window.__miai.sheet().
  if (o.host === "Excel") {
    // Projektmappen kan have flere ark (#21). Hvert ark har sin egen tilstand; buildSheet laver arket.
    interface SheetDef {
      sheet: string;
      values?: (string | number)[][];
      formats?: Record<string, { numberFormat?: string; bold?: boolean; fill?: string }>;
      tables?: { name: string; range: string }[];
    }
    interface Built {
      name: () => string;
      worksheet: unknown;
      state: () => unknown;
      used: () => unknown;
      tableObjs: () => unknown[];
    }
    const built: Built[] = [];
    let tableSeq = 0;
    const buildSheet = (def: SheetDef): Built => {
      let sheet = def.sheet;
      const grid: unknown[][] = (def.values ?? [[""]]).map((r) => [...r]);
      interface Fmt {
        numberFormat?: string;
        bold?: boolean;
        fill?: string;
      }
      const fmts = new Map<string, Fmt>();
      const key = (r: number, c: number) => `${r},${c}`;
      const fmtOf = (r: number, c: number): Fmt => fmts.get(key(r, c)) ?? {};
      const setFmt = (r: number, c: number, f: Fmt) => fmts.set(key(r, c), { ...fmtOf(r, c), ...f });
      const tables: { name: string; range: string; style?: string }[] = (def.tables ?? []).map((t) => ({ ...t }));
      const widths = new Map<number, number>();
      const charts: { name: string; type: string; source: string; title?: string; at?: string }[] = [];
      const condFormats: { id: string; address: string; type: string; rule?: unknown }[] = [];
      let frozen: number | null = null;
      const pivots: { name: string; source: string; at: string; rows: string[]; columns: string[]; values: { field: string; fn: string }[] }[] = [];
      const validations = new Map<string, Record<string, unknown>>();
      let cfSeq = 0;
      /** Sorterer rækkerne r0+1..r0+rows-1 (overskriften står) efter kolonne c0+key; formater følger rækken. */
      const sortRows = (r0: number, c0: number, rows: number, cols: number, key: number, ascending: boolean) => {
        const body = Array.from({ length: rows - 1 }, (_, i) => r0 + 1 + i).map((r) => ({
          v: Array.from({ length: cols }, (_, c) => grid[r]?.[c0 + c] ?? ""),
          f: Array.from({ length: cols }, (_, c) => fmts.get(key2(r, c0 + c))),
        }));
        const cmp = (a: unknown, b: unknown) => (typeof a === "number" && typeof b === "number" ? a - b : String(a).localeCompare(String(b), "da"));
        body.sort((a, b) => (ascending ? 1 : -1) * cmp(a.v[key], b.v[key]));
        body.forEach((row, i) =>
          row.v.forEach((v, c) => {
            put(r0 + 1 + i, c0 + c, v);
            if (row.f[c]) fmts.set(key2(r0 + 1 + i, c0 + c), row.f[c]!);
            else fmts.delete(key2(r0 + 1 + i, c0 + c));
          }),
        );
      };
      const key2 = (r: number, c: number) => `${r},${c}`;
      /**
       * Flytter celler (værdier og formater) i kolonnerne c0..c1 fra række `from` og ned med `delta`
       * (negativ = op). Tabeller helt under og inden for kolonnerne flytter med; en tabel, der kun
       * delvist ligger i kolonnerne, afvises som i Excel.
       */
      const shift = (from: number, delta: number, c0: number, c1: number) => {
        for (const t of tables) {
          const [tr, tc, trows, tcols] = parseA(t.range);
          if (tr + trows - 1 < from || tc > c1 || tc + tcols - 1 < c0) continue;
          if (tr < from || tc < c0 || tc + tcols - 1 > c1) throw new Error("InvalidOperation: This operation is attempting to shift cells in a table");
          t.range = `${col(tc)}${tr + 1 + delta}:${col(tc + tcols - 1)}${tr + trows + delta}`;
        }
        const height = Math.max(grid.length, ...[...fmts.keys()].map((k) => Number(k.split(",")[0]) + 1));
        const cells: { r: number; c: number; v: unknown; f?: Fmt }[] = [];
        for (let r = from; r < height; r++)
          for (let c = c0; c <= c1; c++) {
            cells.push({ r, c, v: grid[r]?.[c] ?? "", f: fmts.get(key(r, c)) });
            if (grid[r]) grid[r][c] = "";
            fmts.delete(key(r, c));
          }
        for (const x of cells) {
          const r = x.r + delta;
          if (r < 0) continue;
          if (x.v !== "") put(r, x.c, x.v);
          if (x.f) fmts.set(key(r, x.c), x.f);
        }
        while (grid.length && grid[grid.length - 1].every((v) => v === "" || v === undefined)) grid.pop();
      };
      const col = (i: number) => String.fromCharCode(65 + i);
      const cellOf = (x: string) => ({ c: x.charCodeAt(0) - 65, r: Number(x.slice(1)) - 1 });
      /** "B2" eller "B2:D4" → [r0, c0, rows, cols] (kun enkelt-bogstavs kolonner i simuleringen). */
      const parseA = (a: string): [number, number, number, number] => {
        const [from, to = from] = a.replace(/^.*!/, "").toUpperCase().split(":");
        const f = cellOf(from);
        const t = cellOf(to);
        return [f.r, f.c, t.r - f.r + 1, t.c - f.c + 1];
      };
      for (const [a, f] of Object.entries(def.formats ?? {})) {
        const [r0, c0, rows, cols] = parseA(a);
        for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) setFmt(r0 + r, c0 + c, f);
      }
      const addr = (r0: number, c0: number, rows: number, cols: number) =>
        `${sheet}!${col(c0)}${r0 + 1}${rows * cols > 1 ? `:${col(c0 + cols - 1)}${r0 + rows}` : ""}`;
      const width = () => Math.max(1, ...grid.map((r) => r.length));
      const put = (r: number, c: number, v: unknown) => {
        while (grid.length <= r) grid.push([]);
        grid[r][c] = v;
      };
      const state = () => ({
        values: grid.map((r) => Array.from({ length: width() }, (_, c) => r[c] ?? "")),
        formats: Object.fromEntries(
          [...fmts]
            .map(([k, f]) => [k, Object.fromEntries(Object.entries(f).filter(([, v]) => v !== undefined))] as [string, Fmt])
            .filter(([, f]) => Object.keys(f).length)
            .map(([k, f]) => [
              k
                .split(",")
                .map(Number)
                .reverse()
                .map((n, i) => (i === 0 ? col(n) : n + 1))
                .join(""),
              f,
            ]),
        ),
        tables: tables.map((t) => ({ name: t.name, range: t.range, ...(t.style ? { style: t.style } : {}) })),
        widths: Object.fromEntries([...widths].map(([c, wd]) => [col(c), wd])),
        charts: charts.map((c) => ({ ...c })),
        conditionalFormats: condFormats.map(({ id: _, ...c }) => c),
        frozenRows: frozen,
        pivots: pivots.map(({ name: _, ...p }) => ({ ...p })),
        validations: Object.fromEntries(validations),
      });

      function rng(r0: number, c0: number, rows: number, cols: number): unknown {
        // Udfyldt med "" uden for arkets data, som i Excel.
        const slice = () => Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => grid[r0 + r]?.[c0 + c] ?? ""));
        const each = (fn: (r: number, c: number) => void) => {
          for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) fn(r0 + r, c0 + c);
        };
        const write = (v: unknown[][]) => v.forEach((row, r) => row.forEach((x, c) => x !== null && put(r0 + r, c0 + c, x)));
        const target = {
          address: addr(r0, c0, rows, cols),
          rowCount: rows,
          columnCount: cols,
          rowIndex: r0,
          columnIndex: c0,
          isNullObject: false,
          load() {},
          getCell: (r: number, c: number) => rng(r0 + r, c0 + c, 1, 1),
          getResizedRange: (dr: number, dc: number) => rng(r0, c0, rows + dr, cols + dc),
          getOffsetRange: (dr: number, dc: number) => rng(r0 + dr, c0 + dc, rows, cols),
          getRow: (i: number) => rng(r0 + i, c0, 1, cols),
          getColumn: (i: number) => rng(r0, c0 + i, rows, 1),
          select() {
            record({ kind: "select", address: addr(r0, c0, rows, cols) });
          },
          dataValidation: strict("Excel.DataValidation", {
            load() {},
            get rule() {
              return validations.get(addr(r0, c0, rows, cols).replace(/^.*!/, "")) ?? {};
            },
            set rule(v: Record<string, unknown>) {
              record({ kind: "validation", address: addr(r0, c0, rows, cols), rule: v });
              validations.set(addr(r0, c0, rows, cols).replace(/^.*!/, ""), v);
            },
            clear() {
              validations.delete(addr(r0, c0, rows, cols).replace(/^.*!/, ""));
            },
          }),
          sort: strict("Excel.RangeSort", {
            apply(fields: { key: number; ascending?: boolean }[], _matchCase?: boolean, hasHeaders?: boolean) {
              record({ kind: "sort", address: addr(r0, c0, rows, cols), fields, hasHeaders });
              if (!hasHeaders) throw new Error("Simuleringen kræver en overskriftsrække");
              sortRows(r0, c0, rows, cols, fields[0].key, fields[0].ascending !== false);
            },
          }),
          conditionalFormats: strict("Excel.ConditionalFormatCollection", {
            add(type: string) {
              const cf = { id: `cf${++cfSeq}`, address: addr(r0, c0, rows, cols), type } as { id: string; address: string; type: string; rule?: unknown };
              condFormats.push(cf);
              record({ kind: "conditionalFormat", address: cf.address, type });
              const fmt = () => ({ fill: {}, font: {} });
              const sub = (name: string) => ({
                format: fmt(),
                set rule(v: unknown) {
                  cf.rule = v;
                },
                set criteria(v: unknown) {
                  cf.rule = v;
                },
                _name: name,
              });
              return strict("Excel.ConditionalFormat", {
                load() {},
                get id() {
                  return cf.id;
                },
                colorScale: sub("colorScale"),
                dataBar: sub("dataBar"),
                cellValue: sub("cellValue"),
                preset: sub("preset"),
                topBottom: sub("topBottom"),
              });
            },
            getItem: (id: string) =>
              strict("Excel.ConditionalFormat(item)", {
                delete() {
                  const i = condFormats.findIndex((c) => c.id === id);
                  if (i < 0) throw new Error(`ItemNotFound: ${id}`);
                  condFormats.splice(i, 1);
                },
              }),
          }),
          format: strict("Excel.RangeFormat", {
            load() {},
            get columnWidth() {
              return widths.get(c0) ?? 48;
            },
            set columnWidth(v: number) {
              for (let c = 0; c < cols; c++) widths.set(c0 + c, v);
            },
            autofitColumns() {
              record({ kind: "autofitColumns", address: addr(r0, c0, rows, cols) });
              for (let c = 0; c < cols; c++) widths.set(c0 + c, 10 + 7 * Math.max(1, ...grid.map((row) => String(row[c0 + c] ?? "").length)));
            },
          }),
          delete(direction: string) {
            record({ kind: "rangeDelete", address: addr(r0, c0, rows, cols), direction });
            each((r, c) => {
              if (grid[r]) grid[r][c] = "";
              fmts.delete(key(r, c));
            });
            shift(r0 + rows, -rows, c0, c0 + cols - 1);
          },
          insert(direction: string) {
            record({ kind: "rangeInsert", address: addr(r0, c0, rows, cols), direction });
            shift(r0, rows, c0, c0 + cols - 1);
            return rng(r0, c0, rows, cols);
          },
          worksheet: undefined as unknown,
          get values() {
            return slice();
          },
          set values(v: unknown[][]) {
            record({ kind: "values", address: addr(r0, c0, rows, cols), values: v });
            write(v);
          },
          get formulas() {
            return slice();
          },
          set formulas(v: unknown[][]) {
            record({ kind: "formulas", address: addr(r0, c0, rows, cols), formulas: v });
            write(v);
          },
          get numberFormat() {
            return Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => fmtOf(r0 + r, c0 + c).numberFormat ?? "General"));
          },
          set numberFormat(v: string[][]) {
            v.forEach((row, r) => row.forEach((f, c) => setFmt(r0 + r, c0 + c, { numberFormat: f === "General" ? undefined : f })));
          },
          copyFrom(src: { rowIndex: number; columnIndex: number }, type: string) {
            record({ kind: "copyFrom", address: addr(r0, c0, rows, cols), type });
            each((r, c) => fmts.set(key(r, c), { ...fmtOf(src.rowIndex + (r - r0), src.columnIndex + (c - c0)) }));
          },
          getCellProperties: () => ({
            value: Array.from({ length: rows }, (_, r) =>
              Array.from({ length: cols }, (_, c) => {
                const f = fmtOf(r0 + r, c0 + c);
                return { format: { font: { bold: !!f.bold }, fill: { color: f.fill ?? "#FFFFFF" } } };
              }),
            ),
          }),
          setCellProperties(p: { format: { font: { bold: boolean }; fill: { color: string } } }[][]) {
            p.forEach((row, r) =>
              row.forEach((x, c) =>
                setFmt(r0 + r, c0 + c, { bold: x.format.font.bold || undefined, fill: x.format.fill.color === "#FFFFFF" ? undefined : x.format.fill.color }),
              ),
            );
          },
          clear(applyTo: string) {
            record({ kind: "clear", address: addr(r0, c0, rows, cols), applyTo });
            if (applyTo === "Formats") each((r, c) => fmts.delete(key(r, c)));
            else each((r, c) => put(r, c, ""));
          },
          get valueTypes() {
            return slice().map((r) => r.map((x) => (typeof x === "number" ? "Double" : "String")));
          },
        };
        target.worksheet = worksheet;
        return strict("Excel.Range", target);
      }
      const used = () => rng(0, 0, grid.length, width());
      const parse = (a: string) => rng(...parseA(a));
      const tableObj = (t: { name: string; range: string; style?: string }) =>
        strict("Excel.Table", {
          name: t.name,
          load() {},
          get style() {
            return t.style ?? "TableStyleMedium2";
          },
          set style(v: string) {
            record({ kind: "tableStyle", table: t.name, style: v });
            t.style = v;
          },
          convertToRange() {
            record({ kind: "convertToRange", table: t.name });
            tables.splice(tables.indexOf(t), 1);
            return parse(t.range);
          },
          getRange: () => parse(t.range),
          sort: strict("Excel.TableSort", {
            apply(fields: { key: number; ascending?: boolean }[]) {
              const [r0, c0, rows, cols] = parseA(t.range);
              record({ kind: "tableSort", table: t.name, fields });
              sortRows(r0, c0, rows, cols, fields[0].key, fields[0].ascending !== false);
            },
          }),
          getDataBodyRange: () => {
            const [r0, c0, rows, cols] = parseA(t.range);
            return rng(r0 + 1, c0, rows - 1, cols);
          },
          getHeaderRowRange: () => {
            const [r0, c0, , cols] = parseA(t.range);
            return rng(r0, c0, 1, cols);
          },
          rows: strict("Excel.TableRowCollection", {
            load() {},
            get count() {
              return parseA(t.range)[2] - 1;
            },
            add(index: number | undefined, values: unknown[][]) {
              // Som Excel: rækkerne lægges ind i tabellen og får formatet fra tabellens sidste række.
              const [r0, c0, rows, cols] = parseA(t.range);
              record({ kind: "tableRowsAdd", table: t.name, index, values });
              if (index !== undefined && index < rows - 1) {
                // Midt i tabellen: rækkerne under rykker ned.
                const at = r0 + 1 + index;
                grid.splice(at, 0, ...values.map(() => [] as unknown[]));
                const moved = [...fmts].map(([kk, f]) => {
                  const [r, c] = kk.split(",").map(Number);
                  return [key(r >= at ? r + values.length : r, c), f] as const;
                });
                fmts.clear();
                for (const [kk, f] of moved) fmts.set(kk, f);
                values.forEach((row, i) => row.forEach((v, c) => put(at + i, c0 + c, v ?? "")));
                t.range = `${col(c0)}${r0 + 1}:${col(c0 + cols - 1)}${r0 + rows + values.length}`;
                return;
              }
              values.forEach((row, i) => {
                const r = r0 + rows + i;
                for (let c = 0; c < cols; c++) {
                  put(r, c0 + c, row[c] ?? "");
                  fmts.set(key(r, c0 + c), { ...fmtOf(r0 + rows - 1, c0 + c) });
                }
              });
              t.range = `${col(c0)}${r0 + 1}:${col(c0 + cols - 1)}${r0 + rows + values.length}`;
            },
            getItemAt: (k: number) =>
              strict("Excel.TableRow", {
                getRange: () => {
                  const [r0, c0, , cols] = parseA(t.range);
                  return rng(r0 + 1 + k, c0, 1, cols);
                },
                delete() {
                  const [r0, c0, rows, cols] = parseA(t.range);
                  record({ kind: "tableRowDelete", table: t.name, index: k });
                  const gone = r0 + 1 + k;
                  grid.splice(gone, 1);
                  // Formaterne følger cellerne: rækken forsvinder, og rækkerne under rykker op.
                  const moved = [...fmts].flatMap(([kk, f]) => {
                    const [r, c] = kk.split(",").map(Number);
                    return r === gone ? [] : [[key(r > gone ? r - 1 : r, c), f] as const];
                  });
                  fmts.clear();
                  for (const [kk, f] of moved) fmts.set(kk, f);
                  t.range = `${col(c0)}${r0 + 1}:${col(c0 + cols - 1)}${r0 + rows - 1}`;
                },
              }),
          }),
        });
      const worksheet = strict("Excel.Worksheet", {
        activate() {
          record({ kind: "activate", sheet });
        },
        get name() {
          return sheet;
        },
        set name(v: string) {
          if (built.some((b) => b.name().toLowerCase() === v.toLowerCase())) throw new Error(`InvalidArgument: arket ${v} findes allerede`);
          record({ kind: "sheetRename", from: sheet, to: v });
          sheet = v;
        },
        load() {},
        delete() {
          if (built.length === 1) throw new Error("InvalidOperation: projektmappen skal have mindst ét ark");
          record({ kind: "sheetDelete", name: sheet });
          built.splice(built.indexOf(self), 1);
        },
        charts: strict("Excel.ChartCollection", {
          load() {},
          get items() {
            return charts.map((c) => strict("Excel.Chart(item)", { name: c.name, delete: () => charts.splice(charts.indexOf(c), 1) }));
          },
          add(type: string, source: { address: string }, seriesBy: string) {
            const c = { name: `Diagram ${charts.length + 1}`, type, source: source.address.replace(/^.*!/, "") } as (typeof charts)[number];
            charts.push(c);
            record({ kind: "chartAdd", type, source: c.source, seriesBy });
            return strict("Excel.Chart", {
              load() {},
              get name() {
                return c.name;
              },
              title: strict("Excel.ChartTitle", {
                set text(v: string) {
                  c.title = v;
                },
              }),
              setPosition(start: { address: string }) {
                c.at = start.address.replace(/^.*!/, "");
              },
            });
          },
          getItem: (name: string) =>
            strict("Excel.Chart(item)", {
              delete() {
                const i = charts.findIndex((c) => c.name === name);
                if (i < 0) throw new Error(`ItemNotFound: ${name}`);
                charts.splice(i, 1);
              },
            }),
        }),
        pivotTables: strict("Excel.PivotTableCollection", {
          load() {},
          get items() {
            return pivots.map((p) => strict("Excel.PivotTable(item)", { name: p.name, delete: () => pivots.splice(pivots.indexOf(p), 1) }));
          },
          // Kilden er et område (med address) eller en tabel (med overskriftsrække).
          add(
            name: string,
            source: { address?: string; name?: string; getHeaderRowRange?: () => { values: unknown[][] }; values?: unknown[][] },
            dest: { address: string },
          ) {
            const isTable = "getHeaderRowRange" in source;
            const headers = (isTable ? source.getHeaderRowRange!().values[0] : (source.values?.[0] ?? [])).map(String);
            const p = {
              name,
              source: isTable ? String(source.name) : String(source.address),
              at: dest.address.replace(/^.*!/, ""),
              rows: [],
              columns: [],
              values: [],
            } as (typeof pivots)[number];
            pivots.push(p);
            record({ kind: "pivotAdd", source: p.source, at: p.at });
            const hierarchy = (field: string) => {
              // Som Excel: et felt, der ikke er en overskrift i kilden, afvises.
              if (!headers.includes(field)) throw new Error(`ItemNotFound: feltet ${field} findes ikke i kilden`);
              return { field };
            };
            return strict("Excel.PivotTable", {
              hierarchies: strict("Excel.PivotHierarchyCollection", { getItem: hierarchy }),
              rowHierarchies: strict("Excel.RowColumnPivotHierarchyCollection", { add: (h: { field: string }) => void p.rows.push(h.field) }),
              columnHierarchies: strict("Excel.RowColumnPivotHierarchyCollection", { add: (h: { field: string }) => void p.columns.push(h.field) }),
              dataHierarchies: strict("Excel.DataPivotHierarchyCollection", {
                add(h: { field: string }) {
                  const v = { field: h.field, fn: "Sum" };
                  p.values.push(v);
                  return {
                    set summarizeBy(fn: string) {
                      v.fn = fn;
                    },
                  };
                },
              }),
            });
          },
          getItem: (name: string) =>
            strict("Excel.PivotTable(item)", {
              delete() {
                const i = pivots.findIndex((x) => x.name === name);
                if (i < 0) throw new Error(`ItemNotFound: ${name}`);
                pivots.splice(i, 1);
              },
            }),
        }),
        freezePanes: strict("Excel.WorksheetFreezePanes", {
          freezeRows(n: number) {
            record({ kind: "freezeRows", n });
            frozen = n;
          },
          unfreeze() {
            record({ kind: "unfreeze" });
            frozen = null;
          },
          freezeAt(r: { rowCount: number }) {
            frozen = r.rowCount;
          },
          getLocationOrNullObject: () =>
            frozen === null
              ? strict("Excel.Range(null)", { isNullObject: true, load() {} })
              : (() => {
                  const r = rng(0, 0, frozen, 1) as Record<string, unknown>;
                  return r;
                })(),
        }),
        // Et helt tomt ark (fx et nyt) har intet brugt område.
        getUsedRangeOrNullObject: () => (grid.length === 0 ? strict("Excel.Range(null)", { isNullObject: true, load() {} }) : used()),
        getRange: (a?: string) => (a ? parse(a) : used()),
        tables: strict("Excel.TableCollection", {
          load() {},
          get items() {
            return tables.map(tableObj);
          },
          getItem: (name: string) => {
            const t = tables.find((x) => x.name === name);
            if (!t) throw new Error(`ItemNotFound: ${name}`);
            return tableObj(t);
          },
          add(r: { address: string }, hasHeaders: boolean) {
            const range = r.address.replace(/^.*!/, "");
            const [r0, c0, rows, cols] = parseA(range);
            if (
              tables.some((t) => {
                const [a, b, n, m] = parseA(t.range);
                return a <= r0 + rows - 1 && r0 <= a + n - 1 && b <= c0 + cols - 1 && c0 <= b + m - 1;
              })
            )
              throw new Error("InvalidOperation: A table can't overlap another table");
            const t = { name: `Tabel${++tableSeq}`, range };
            record({ kind: "tableAdd", address: range, hasHeaders });
            tables.push(t);
            return tableObj(t);
          },
        }),
      });
      const self: Built = { name: () => sheet, worksheet, state, used, tableObjs: () => tables.map(tableObj) };
      return self;
    };
    built.push(buildSheet({ sheet: o.excel?.sheet ?? "Ark1", values: o.excel?.values, formats: o.excel?.formats, tables: o.excel?.tables }));
    for (const d of o.excel?.sheets ?? []) built.push(buildSheet(d));
    // Det første ark er det aktive (markeringen); __miai.sheet() viser det, __miai.sheets() alle ark.
    (w.__miai as Record<string, unknown>).sheet = () => built[0].state();
    (w.__miai as Record<string, unknown>).sheets = () => Object.fromEntries(built.map((b) => [b.name(), b.state()]));
    const findSheet = (name: string) => {
      const b = built.find((x) => x.name().toLowerCase() === name.toLowerCase());
      if (!b) throw new Error(`ItemNotFound: arket ${name} findes ikke`);
      return b;
    };
    const workbook = strict("Excel.Workbook", {
      getSelectedRange: () => built[0].used(),
      worksheets: strict("Excel.WorksheetCollection", {
        load() {},
        get items() {
          return built.map((b) => b.worksheet);
        },
        getItem: (name: string) => findSheet(name).worksheet,
        getActiveWorksheet: () => built[0].worksheet,
        add(name: string) {
          if (built.some((b) => b.name().toLowerCase() === name.toLowerCase())) throw new Error(`InvalidArgument: arket ${name} findes allerede`);
          record({ kind: "sheetAdd", name });
          const b = buildSheet({ sheet: name, values: [] });
          built.push(b);
          return b.worksheet;
        },
      }),
      get tables() {
        return strict("Excel.TableCollection(workbook)", {
          load() {},
          get items() {
            return built.flatMap((b) => b.tableObjs());
          },
          getItem(name: string) {
            const t = built.flatMap((b) => b.tableObjs() as { name: string }[]).find((x) => x.name === name);
            if (!t) throw new Error(`ItemNotFound: tabellen ${name} findes ikke`);
            return t;
          },
        });
      },
      pivotTables: strict("Excel.PivotTableCollection", { load() {}, items: [] }),
      comments: { add: (address: string, text: string) => record({ kind: "comment", address, text }) },
    });
    w.Excel = strict("Excel", {
      RangeValueType: { error: "Error" },
      ClearApplyTo: { contents: "Contents", formats: "Formats", all: "All" },
      DeleteShiftDirection: { up: "Up" },
      ChartSeriesBy: { auto: "Auto" },
      ConditionalFormatType: { colorScale: "ColorScale", dataBar: "DataBar", cellValue: "CellValue", presetCriteria: "PresetCriteria", topBottom: "TopBottom" },
      ConditionalFormatColorCriterionType: { lowestValue: "LowestValue", highestValue: "HighestValue", percentile: "Percentile" },
      ConditionalFormatPresetCriterion: { duplicateValues: "DuplicateValues" },
      InsertShiftDirection: { down: "Down" },
      RangeCopyType: { formats: "Formats" },
      run: async (fn: (c: unknown) => unknown) => fn(strict("Excel.RequestContext", { workbook, sync })),
    });
  }

  // --- PowerPoint med præsentationsmodel (agent-tilstand, #6) ------------------------
  // Slides med layout, pladsholdere og formateret tekst. Som PowerPoint: ny tekst i en figur får
  // figurens formatering (første tegns), og en del af teksten kan ændres for sig (getSubstring).
  // Tilstanden kan læses med window.__miai.deck().
  if (o.host === "PowerPoint" && o.slides) {
    interface Mark {
      start: number;
      end: number;
      bold?: boolean;
      italic?: boolean;
    }
    interface Sh {
      name: string;
      type: string;
      ph?: string;
      text: string;
      /** Figurens egen formatering (fx farve og størrelse fra skabelonen) – skal bevares. */
      style: Record<string, unknown>;
      marks: Mark[];
    }
    interface Sl {
      id: string;
      layout: string;
      shapes: Sh[];
    }
    const layouts = (o.layouts ?? [{ name: "Title and Content", placeholders: ["Title", "Body"] }]).map((l, i) => ({ id: `L${i + 1}`, ...l }));
    let nextId = 100;
    let deck: Sl[] = o.slides.map((sl, i) => ({
      id: `S${i + 1}`,
      layout: sl.layout,
      shapes: sl.shapes.map((sh) => ({ type: "Placeholder", style: {}, marks: [], ...sh })),
    }));
    const clone = <T>(x: T): T => JSON.parse(JSON.stringify(x));
    (w.__miai as Record<string, unknown>).deck = () => clone(deck);
    (w.__miai as Record<string, unknown>).pptx = [] as string[];

    const textRange = (sh: Sh) =>
      strict("PowerPoint.TextRange", {
        load() {},
        get text() {
          return sh.text;
        },
        set text(v: string) {
          // Hele teksten erstattes; den nye tekst får formateringen fra første tegn.
          record({ kind: "setText", shape: sh.name, text: v });
          const first = sh.marks.find((m) => m.start === 0);
          sh.marks = first ? [{ ...first, start: 0, end: v.length }] : [];
          sh.text = v;
        },
        getSubstring: (start: number, len: number) =>
          strict("PowerPoint.TextRange(sub)", {
            set text(v: string) {
              record({ kind: "replace", shape: sh.name, from: sh.text.slice(start, start + len), to: v });
              const delta = v.length - len;
              sh.marks = sh.marks
                .map((m) =>
                  m.start >= start + len
                    ? { ...m, start: m.start + delta, end: m.end + delta }
                    : m.end > start && m.start <= start
                      ? { ...m, end: Math.max(start + v.length, m.end + delta) }
                      : m,
                )
                .filter((m) => m.end > m.start);
              sh.text = sh.text.slice(0, start) + v + sh.text.slice(start + len);
            },
            font: {
              set bold(b: boolean) {
                if (b) sh.marks.push({ start, end: start + len, bold: true });
              },
              set italic(b: boolean) {
                if (b) sh.marks.push({ start, end: start + len, italic: true });
              },
            },
          }),
      });
    const shapeObj = (sh: Sh) =>
      strict("PowerPoint.Shape", {
        name: sh.name,
        type: sh.type,
        textFrame: strict("PowerPoint.TextFrame", { hasText: sh.text.length > 0, textRange: textRange(sh) }),
        placeholderFormat: strict("PowerPoint.PlaceholderFormat", { type: sh.ph ?? "Body", load() {} }),
      });
    const slideObj = (sl: Sl) =>
      strict("PowerPoint.Slide", {
        id: sl.id,
        load() {},
        get index() {
          return deck.indexOf(sl);
        },
        layout: { name: sl.layout, id: layouts.find((l) => l.name === sl.layout)?.id ?? "L1" },
        shapes: { load() {}, items: sl.shapes.map(shapeObj) },
        delete() {
          record({ kind: "deleteSlide", id: sl.id });
          deck = deck.filter((x) => x !== sl);
        },
        moveTo(i: number) {
          record({ kind: "moveSlide", id: sl.id, to: i });
          deck = deck.filter((x) => x !== sl);
          deck.splice(i, 0, sl);
        },
        exportAsBase64: () => ({ value: JSON.stringify(sl) }),
      });
    const slidesCol = () =>
      strict("PowerPoint.SlideCollection", {
        load() {},
        get items() {
          return deck.map(slideObj);
        },
        add(opts?: { layoutId?: string }) {
          const l = layouts.find((x) => x.id === opts?.layoutId) ?? layouts[0];
          record({ kind: "addSlide", layout: l.name });
          deck.push({
            id: `S${nextId++}`,
            layout: l.name,
            shapes: l.placeholders.map((ph, i) => ({ name: `${ph} ${i + 1}`, type: "Placeholder", ph, text: "", style: { fromLayout: l.name }, marks: [] })),
          });
        },
      });
    const presentation = strict("PowerPoint.Presentation", {
      get slides() {
        return slidesCol();
      },
      slideMasters: strict("PowerPoint.SlideMasterCollection", {
        load() {},
        items: [{ id: "M1", layouts: { items: layouts.map((l) => ({ id: l.id, name: l.name })) } }],
      }),
      insertSlidesFromBase64(b64: string, opts: { targetSlideId?: string; formatting?: string }) {
        const at = opts.targetSlideId ? deck.findIndex((x) => x.id === opts.targetSlideId) + 1 : 0;
        let sl: Sl | undefined;
        try {
          sl = JSON.parse(b64) as Sl; // et slide eksporteret af simuleringen selv
        } catch {
          // En rigtig .pptx (designets master og slides): filnavnene står ukomprimeret i zip-filen,
          // så antallet af slides kan tælles. Filen gemmes, så testen kan pakke den ud.
          const names = new Set(atob(b64).match(/ppt\/slides\/slide\d+\.xml/g) ?? []);
          record({ kind: "insertPptx", after: opts.targetSlideId ?? null, formatting: opts.formatting, slides: names.size });
          ((w.__miai as Record<string, unknown>).pptx as string[]).push(b64);
          const added = [...names].map(() => ({ id: `S${nextId++}`, layout: "(fra designet)", shapes: [] as Sh[] }));
          deck.splice(at, 0, ...added);
          return;
        }
        record({ kind: "insertSlides", after: opts.targetSlideId ?? null });
        deck.splice(at, 0, { ...sl, id: `S${nextId++}` });
      },
    });
    w.PowerPoint = strict("PowerPoint", {
      ShapeType: { placeholder: "Placeholder" },
      InsertSlideFormatting: { keepSourceFormatting: "KeepSourceFormatting" },
      run: async (fn: (c: unknown) => unknown) => fn(strict("PowerPoint.RequestContext", { presentation, sync })),
    });
  }

  // --- PowerPoint -------------------------------------------------------------
  if (o.host === "PowerPoint" && !o.slides) {
    const textFrame = () => ({ textRange: { text: "", font: {} }, wordWrap: false });
    const shape = (kind: string, type: string, opts: Record<string, number>) => {
      const s = {
        kind,
        type,
        ...opts,
        textFrame: textFrame(),
        lineFormat: { visible: true },
        fill: { setImage: (b64: string) => record({ kind: "fillImage", bytes: b64.length }) },
        name: "",
      };
      record(s);
      return s;
    };
    const shapes = strict("PowerPoint.ShapeCollection", {
      addLine: (type: string, opts: Record<string, number>) => {
        if (opts.width < 0 || opts.height < 0) throw new Error("InvalidArgument: negative mål");
        return shape("line", type, opts);
      },
      addGeometricShape: (type: string, opts: Record<string, number>) => {
        if (opts.width < 0 || opts.height < 0) throw new Error("InvalidArgument: negative mål");
        return shape("shape", type, opts);
      },
      addGroup: (items: unknown[]) => record({ kind: "group", count: items.length }),
    });
    const ctx = strict("PowerPoint.RequestContext", {
      presentation: strict("PowerPoint.Presentation", { getSelectedSlides: () => ({ items: [{ shapes }], load() {} }) }),
      sync,
    });
    w.PowerPoint = strict("PowerPoint", {
      ConnectorType: { straight: "Straight" },
      GeometricShapeType: { roundRectangle: "RoundRectangle", chevron: "Chevron", rectangle: "Rectangle" },
      run: async (fn: (c: unknown) => unknown) => fn(ctx),
    });
  }
}

/** Scriptet, der serveres i stedet for office.js. */
export function fakeOfficeScript(o: FakeOfficeOptions): string {
  return `(${installFakeOffice.toString()})(${JSON.stringify(o)});`;
}
