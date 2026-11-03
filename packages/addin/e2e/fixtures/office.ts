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

export type HostName = "Word" | "Excel" | "PowerPoint";

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
  };
  displayLanguage?: string;
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

  // --- Fælles Office-objekt -------------------------------------------------
  let selection = o.selection ?? "";
  w.Office = strict("Office", {
    HostType: { Word: "Word", Excel: "Excel", PowerPoint: "PowerPoint" },
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    EventType: { OfficeThemeChanged: "officeThemeChanged" },
    context: strict("Office.context", {
      displayLanguage: o.displayLanguage ?? "da-DK",
      officeTheme,
      requirements: { isSetSupported: (name: string, v = "1.1") => versionAtLeast(o.apiSets?.[name], v) },
      document: strict("Office.context.document", {
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
    const sheet = o.excel?.sheet ?? "Ark1";
    const grid: unknown[][] = (o.excel?.values ?? [[""]]).map((r) => [...r]);
    interface Fmt {
      numberFormat?: string;
      bold?: boolean;
      fill?: string;
    }
    const fmts = new Map<string, Fmt>();
    const key = (r: number, c: number) => `${r},${c}`;
    const fmtOf = (r: number, c: number): Fmt => fmts.get(key(r, c)) ?? {};
    const setFmt = (r: number, c: number, f: Fmt) => fmts.set(key(r, c), { ...fmtOf(r, c), ...f });
    const tables = (o.excel?.tables ?? []).map((t) => ({ ...t }));
    const col = (i: number) => String.fromCharCode(65 + i);
    const cellOf = (x: string) => ({ c: x.charCodeAt(0) - 65, r: Number(x.slice(1)) - 1 });
    /** "B2" eller "B2:D4" → [r0, c0, rows, cols] (kun enkelt-bogstavs kolonner i simuleringen). */
    const parseA = (a: string): [number, number, number, number] => {
      const [from, to = from] = a.replace(/^.*!/, "").toUpperCase().split(":");
      const f = cellOf(from);
      const t = cellOf(to);
      return [f.r, f.c, t.r - f.r + 1, t.c - f.c + 1];
    };
    for (const [a, f] of Object.entries(o.excel?.formats ?? {})) {
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
    (w.__miai as Record<string, unknown>).sheet = () => ({
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
      tables: tables.map((t) => ({ name: t.name, range: t.range })),
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
          each((r, c) => put(r, c, ""));
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
    const tableObj = (t: { name: string; range: string }) =>
      strict("Excel.Table", {
        name: t.name,
        getRange: () => parse(t.range),
        getHeaderRowRange: () => {
          const [r0, c0, , cols] = parseA(t.range);
          return rng(r0, c0, 1, cols);
        },
        rows: strict("Excel.TableRowCollection", {
          load() {},
          get count() {
            return parseA(t.range)[2] - 1;
          },
          add(_index: unknown, values: unknown[][]) {
            // Som Excel: rækkerne lægges ind i tabellen og får formatet fra tabellens sidste række.
            const [r0, c0, rows, cols] = parseA(t.range);
            record({ kind: "tableRowsAdd", table: t.name, values });
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
      name: sheet,
      load() {},
      getUsedRangeOrNullObject: used,
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
      }),
    });
    const workbook = strict("Excel.Workbook", {
      getSelectedRange: used,
      worksheets: strict("Excel.WorksheetCollection", {
        load() {},
        items: [worksheet],
        getItem: () => worksheet,
      }),
      comments: { add: (address: string, text: string) => record({ kind: "comment", address, text }) },
    });
    w.Excel = strict("Excel", {
      RangeValueType: { error: "Error" },
      ClearApplyTo: { contents: "Contents" },
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
