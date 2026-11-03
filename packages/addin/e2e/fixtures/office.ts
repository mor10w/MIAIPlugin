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
  /** Excel: arkets værdier (række 1 = A1). Markeringen er hele området. */
  excel?: { sheet: string; values: (string | number)[][] };
  displayLanguage?: string;
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
  if (o.host === "Excel") {
    const sheet = o.excel?.sheet ?? "Ark1";
    const grid = o.excel?.values ?? [[""]];
    const col = (i: number) => String.fromCharCode(65 + i);
    const addr = (r0: number, c0: number, rows: number, cols: number) =>
      `${sheet}!${col(c0)}${r0 + 1}${rows * cols > 1 ? `:${col(c0 + cols - 1)}${r0 + rows}` : ""}`;
    function rng(r0: number, c0: number, rows: number, cols: number): unknown {
      // Udfyldt med "" uden for arkets data, som i Excel.
      const slice = () => Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => grid[r0 + r]?.[c0 + c] ?? ""));
      const target = {
        address: addr(r0, c0, rows, cols),
        rowCount: rows,
        columnCount: cols,
        isNullObject: false,
        load() {},
        getCell: (r: number, c: number) => rng(r0 + r, c0 + c, 1, 1),
        getResizedRange: (dr: number, dc: number) => rng(r0, c0, rows + dr, cols + dc),
        worksheet: undefined as unknown,
        get values() {
          return slice();
        },
        set values(v: unknown) {
          record({ kind: "values", address: addr(r0, c0, rows, cols), values: v });
        },
        get formulas() {
          return slice();
        },
        set formulas(v: unknown) {
          record({ kind: "formulas", address: addr(r0, c0, rows, cols), formulas: v });
        },
        clear: (applyTo: string) => record({ kind: "clear", address: addr(r0, c0, rows, cols), applyTo }),
        get valueTypes() {
          return slice().map((r) => r.map((x) => (typeof x === "number" ? "Double" : "String")));
        },
      };
      target.worksheet = worksheet;
      return strict("Excel.Range", target);
    }
    const used = () => rng(0, 0, grid.length, grid[0]?.length ?? 1);
    /** "B2" eller "B2:D4" → område (kun enkelt-bogstavs kolonner i simuleringen). */
    const parse = (a: string) => {
      const [from, to = from] = a.toUpperCase().split(":");
      const cell = (x: string) => ({ c: x.charCodeAt(0) - 65, r: Number(x.slice(1)) - 1 });
      const f = cell(from);
      const t = cell(to);
      return rng(f.r, f.c, t.r - f.r + 1, t.c - f.c + 1);
    };
    const worksheet = strict("Excel.Worksheet", { name: sheet, load() {}, getUsedRangeOrNullObject: used });
    const workbook = strict("Excel.Workbook", {
      getSelectedRange: used,
      worksheets: strict("Excel.WorksheetCollection", {
        load() {},
        items: [strict("Excel.Worksheet", { name: sheet, getUsedRangeOrNullObject: used })],
        getItem: () => strict("Excel.Worksheet", { getRange: (a?: string) => (a ? parse(a) : used()) }),
      }),
      comments: { add: (address: string, text: string) => record({ kind: "comment", address, text }) },
    });
    w.Excel = strict("Excel", {
      RangeValueType: { error: "Error" },
      ClearApplyTo: { contents: "Contents" },
      run: async (fn: (c: unknown) => unknown) => fn(strict("Excel.RequestContext", { workbook, sync })),
    });
  }

  // --- PowerPoint -------------------------------------------------------------
  if (o.host === "PowerPoint") {
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
