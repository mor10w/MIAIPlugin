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

  // --- Word -------------------------------------------------------------------
  if (o.host === "Word") {
    const picture = (base64: string, loc: string) => record({ kind: "picture", bytes: base64.length, loc });
    const paragraph = () =>
      strict("Word.Paragraph", {
        insertInlinePictureFromBase64: picture,
        insertHtml: (html: string, loc: string) => record({ kind: "insertHtml", html, loc }),
      });
    const range: Record<string, unknown> = strict("Word.Range", {
      text: selection,
      load() {},
      insertText: (text: string, loc: string) => record({ kind: "insertText", text, loc }),
      insertHtml: (html: string, loc: string) => record({ kind: "insertHtml", html, loc }),
      insertComment: (text: string) => record({ kind: "comment", text }),
      insertParagraph: () => paragraph(),
      insertInlinePictureFromBase64: picture,
      getRange: () => range,
      search(text: string) {
        // Én forekomst pr. gang teksten findes i markeringen.
        const items: unknown[] = [];
        let i = selection.indexOf(text);
        while (i >= 0) {
          items.push(strict("Word.Range(search)", { insertText: (rep: string, loc: string) => record({ kind: "replace", from: text, to: rep, loc }) }));
          i = selection.indexOf(text, i + text.length);
        }
        return { items, load() {} };
      },
    });
    const document = strict("Word.Document", {
      getSelection: () => range,
      body: strict("Word.Body", { getRange: () => range, insertParagraph: () => paragraph() }),
      load() {},
      changeTrackingMode: "Off",
    });
    const ctx = strict("Word.RequestContext", { document, trackedObjects: { add() {}, remove() {} }, sync });
    w.Word = strict("Word", {
      ChangeTrackingMode: { trackAll: "TrackAll", off: "Off" },
      run: (a: unknown, b?: unknown) => (typeof a === "function" ? a(ctx) : (b as (c: unknown, r: unknown) => unknown)(ctx, range)),
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
      const slice = () => grid.slice(r0, r0 + rows).map((r) => r.slice(c0, c0 + cols));
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
        get valueTypes() {
          return slice().map((r) => r.map((x) => (typeof x === "number" ? "Double" : "String")));
        },
      };
      target.worksheet = worksheet;
      return strict("Excel.Range", target);
    }
    const used = () => rng(0, 0, grid.length, grid[0]?.length ?? 1);
    const worksheet = strict("Excel.Worksheet", { name: sheet, load() {}, getUsedRangeOrNullObject: used });
    const workbook = strict("Excel.Workbook", {
      getSelectedRange: used,
      worksheets: strict("Excel.WorksheetCollection", {
        load() {},
        items: [strict("Excel.Worksheet", { name: sheet, getUsedRangeOrNullObject: used })],
        getItem: () => strict("Excel.Worksheet", { getRange: () => used() }),
      }),
      comments: { add: (address: string, text: string) => record({ kind: "comment", address, text }) },
    });
    w.Excel = strict("Excel", {
      RangeValueType: { error: "Error" },
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
