import { rangeBox, writesCells, type AgentAction } from "@miai/shared";
import type { Overwrite } from "./types";

/**
 * Rene hjælpefunktioner til Excel-adapteren (#57): kolonnebogstaver, TSV, filoversigt,
 * rækkegrupper, tælling af overskrevne celler og markering af ændrede celler (#75).
 * Uden Office.js, så de kan testes direkte.
 */

export const toTsv = (values: unknown[][]) => values.map((r) => r.map((v) => String(v ?? "").replace(/[\t\n\r]+/g, " ")).join("\t")).join("\n");

/** 0 → "A", 26 → "AA" */
export const colName = (index: number): string => {
  let n = index + 1;
  let name = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    name = String.fromCharCode(65 + r) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
};

/** "C7" (eller "C7:E9") → 2, 0-baseret kolonneindex */
export const colIndex = (address: string): number => {
  const m = /^([A-Z]+)/.exec(address);
  if (!m) return 0;
  let n = 0;
  for (const ch of m[1]) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
};

const fmtNum = (x: number): string => x.toFixed(2).replace(/\.?0+$/, "") || "0";

interface ColumnStat {
  letter: string;
  nonEmpty: number;
  numeric?: { sum: number; min: number; max: number };
}

/** Tæller ikke-tomme celler og summer/min/max pr. kolonne (kun ægte tal; datoer tæller som tekst). */
function columnStats(values: unknown[][], baseCol: number): ColumnStat[] {
  const cols = values[0]?.length ?? 0;
  const stats: ColumnStat[] = Array.from({ length: cols }, (_, c) => ({ letter: colName(baseCol + c), nonEmpty: 0 }));
  for (const row of values) {
    for (let c = 0; c < cols && c < row.length; c++) {
      const v = row[c];
      if (v === null || v === undefined || v === "") continue;
      stats[c].nonEmpty++;
      if (typeof v === "number" && Number.isFinite(v)) {
        const n = stats[c].numeric ?? (stats[c].numeric = { sum: 0, min: v, max: v });
        n.sum += v;
        if (v < n.min) n.min = v;
        if (v > n.max) n.max = v;
      }
    }
  }
  return stats;
}

/**
 * Kompakt oversigt over filen: arkliste, headers og – med `withStats` – kolonne-statistik
 * regnet over hele used range. Uden statistik vises de første rækker i stedet.
 */
export function buildOverview(
  sheetLines: string[],
  activeName: string,
  usedAddress: string, // fx "A1:E30" – uden arknavn
  usedRows: number,
  usedCols: number,
  values: unknown[][], // første rækker × maks. MAX_COLS kolonner
  withStats = true,
): string {
  const lines: string[] = [];
  lines.push(`Active sheet '${activeName}', used range ${usedAddress} (${usedRows} rows x ${usedCols} columns)`);
  if (sheetLines.length > 1) lines.push(`All sheets: ${sheetLines.join(" | ")}`);

  const stats = columnStats(values, colIndex(usedAddress));
  const firstRow = values[0] ?? [];
  if (firstRow.some((v) => typeof v === "string" && v.trim() !== "")) {
    lines.push(`Headers (row 1): ${stats.map((s, i) => `${s.letter}=${String(firstRow[i] ?? "").trim() || "-"}`).join(" ")}`);
  }
  if (!withStats) {
    lines.push(`First rows:\n${toTsv(values)}`);
    return lines.join("\n");
  }
  lines.push(
    `Column stats (${values.length} rows): ${stats.map((s) => `${s.letter}: ${s.nonEmpty} non-empty${s.numeric ? `, numeric sum=${fmtNum(s.numeric.sum)} min=${fmtNum(s.numeric.min)} max=${fmtNum(s.numeric.max)}` : ""}`).join(" ")}`,
  );
  if (values.length < usedRows) lines.push(`Stats cover the first ${values.length} of ${usedRows} rows.`);
  const statCols = values[0]?.length ?? 0;
  if (statCols < usedCols) lines.push(`Stats cover the first ${statCols} of ${usedCols} columns.`);

  return lines.join("\n");
}

/** Heltal fra og med `from` til og med `to`. */
export const range = (from: number, to: number) => Array.from({ length: to - from + 1 }, (_, k) => from + k);

/** Sorterede rækkenumre → sammenhængende stykker [fra, til]. */
export function groups(rows: number[]): [number, number][] {
  const out: [number, number][] = [];
  for (const r of rows) {
    const last = out[out.length - 1];
    if (last && r === last[1] + 1) last[1] = r;
    else out.push([r, r]);
  }
  return out;
}

type Box = ReturnType<typeof rangeBox>;

/** 0-baserede, inklusive indeks → "B2:D4" (eller "B2" for én celle), uden arknavn. */
export const boxAddress = (b: Box): string => {
  const from = `${colName(b.c0)}${b.r0 + 1}`;
  const to = `${colName(b.c1)}${b.r1 + 1}`;
  return from === to ? from : `${from}:${to}`;
};

/**
 * Advarsel før overskrivning (#75): det område, en handling skriver i (uden arknavn), og det, den
 * skriver. Samme mål som ved udførelsen: setValues skriver fra områdets øverste venstre celle og så
 * langt, værdierne rækker; setFormula rammer én celle; clear rydder hele området (ingen nye værdier).
 * Andre handlinger erstatter ikke celleindhold og giver undefined.
 */
export function writeTarget(a: AgentAction): { address: string; incoming?: unknown[][] } | undefined {
  if (!writesCells(a)) return undefined;
  const b = rangeBox(a.address);
  if (![b.r0, b.c0, b.r1, b.c1].every((n) => Number.isInteger(n) && n >= 0)) return undefined;
  if (a.action === "clear") return { address: boxAddress(b) };
  if (a.action === "setFormula") {
    const f = a.formula.trim();
    return { address: boxAddress({ ...b, r1: b.r0, c1: b.c0 }), incoming: [[f.startsWith("=") ? f : `=${f}`]] };
  }
  const cols = Math.max(...a.values.map((row) => row.length));
  if (cols < 1) return undefined;
  return {
    address: boxAddress({ r0: b.r0, c0: b.c0, r1: b.r0 + a.values.length - 1, c1: b.c0 + cols - 1 }),
    // Korte rækker fyldes ud med null, som ved udførelsen: null lader cellen være, som den er.
    incoming: a.values.map((row) => [...row, ...Array<null>(cols - row.length).fill(null)]),
  };
}

/**
 * Tæller de celler med indhold, der bliver overskrevet, og hvor mange af dem der er formler (#75).
 * `formulas` er områdets formler, som Office.js giver dem (formel som "=…", ellers værdien, tom
 * celle som ""). `incoming` er det, der skrives: null rører ikke cellen, og en celle, der får
 * præcis det, den har i forvejen, tæller heller ikke med. Uden `incoming` ryddes hele området.
 */
export function countOverwrite(formulas: unknown[][], incoming?: unknown[][]): Overwrite {
  const out: Overwrite = { cells: 0, formulas: 0 };
  formulas.forEach((row, r) =>
    row.forEach((old, c) => {
      if (old === "" || old === null || old === undefined) return;
      if (incoming) {
        const next = incoming[r]?.[c] ?? null;
        if (next === null || String(next) === String(old)) return;
      }
      out.cells++;
      if (typeof old === "string" && old.startsWith("=")) out.formulas++;
    }),
  );
  return out;
}

/** De sidste `n` rækker af et område (uden arknavn), fx nye rækker sidst i en tabel. */
export const lastRows = (address: string, n: number): string => {
  const b = rangeBox(address);
  return boxAddress({ ...b, r0: Math.max(b.r0, b.r1 - n + 1) });
};

/**
 * Markering af ændrede celler (#75): hvad der skal gøres for at give et markeret område sit fyld
 * tilbage. `fills` er cellernes fyldfarve fra før markeringen (undefined = intet fyld). Celler uden
 * fyld får fyldet ryddet (`colors` mangler), celler med fyld får farven sat igen. Er hele området
 * ens (alt med fyld eller alt uden), er det ét stykke; ellers ét stykke pr. sammenhængende række-del.
 */
export function fillPieces(address: string, fills: (string | undefined)[][]): { address: string; colors?: string[][] }[] {
  const b = rangeBox(address);
  const flat = fills.flat();
  if (flat.every((f) => f === undefined)) return [{ address: boxAddress(b) }];
  if (flat.every((f) => f !== undefined)) return [{ address: boxAddress(b), colors: fills as string[][] }];
  const out: { address: string; colors?: string[][] }[] = [];
  fills.forEach((row, r) => {
    for (let c = 0; c < row.length;) {
      const filled = row[c] !== undefined;
      let end = c;
      while (end + 1 < row.length && (row[end + 1] !== undefined) === filled) end++;
      const piece = boxAddress({ r0: b.r0 + r, r1: b.r0 + r, c0: b.c0 + c, c1: b.c0 + end });
      out.push(filled ? { address: piece, colors: [row.slice(c, end + 1) as string[]] } : { address: piece });
      c = end + 1;
    }
  });
  return out;
}
