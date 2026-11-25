/**
 * Regnearksfunktionen =MIAI.ASK() (#80): fra Excels argumenter til prompttekst, og fra modellens
 * svar til en celleværdi. Ren logik uden Office.js, så den kan testes for sig.
 */

/** En celle kan højst rumme 32.767 tegn. */
export const MAX_CELL_CHARS = 32_767;
/** Mere input end dette sendes ikke (DocContext.text har samme loft). */
export const MAX_INPUT_CHARS = 200_000;

const cellText = (v: unknown): string => (v === null || v === undefined ? "" : String(v).replace(/[\t\r\n]+/g, " "));

/**
 * Funktionens `input`: Excel giver et område som rækker af celler og en enkelt celle som 1×1.
 * Én celle bliver til ren tekst (linjeskift bevares), et område til TSV, præcis som cellerne står.
 */
export function inputToText(input: unknown): string {
  if (input === null || input === undefined) return "";
  if (!Array.isArray(input)) return String(input);
  const rows = (input as unknown[]).map((r) => (Array.isArray(r) ? r : [r]));
  if (rows.length === 1 && rows[0].length === 1) return rows[0][0] === null || rows[0][0] === undefined ? "" : String(rows[0][0]);
  const lines = rows.map((r) => r.map(cellText).join("\t"));
  // Tomme rækker i bunden af et for stort område (fx A2:A1000) er støj for modellen.
  while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
  return lines.join("\n");
}

/** Fjerner det, modeller pakker en enkelt værdi ind i: kodehegn og omsluttende citationstegn. */
export function cleanAnswer(text: string): string {
  let s = text.trim();
  const fence = /^```[^\n]*\n([\s\S]*?)\n?```$/.exec(s);
  if (fence) s = fence[1].trim();
  const quoted = /^(["'“”„«»`])([\s\S]*)(["'“”„«»`])$/.exec(s);
  if (quoted && !/["'“”„«»`]/.test(quoted[2])) s = quoted[2].trim();
  return s.length > MAX_CELL_CHARS ? s.slice(0, MAX_CELL_CHARS) : s;
}

const leadingZeroOk = (int: string) => int === "0" || !int.startsWith("0");
const grouped = (s: string, sep: string) => new RegExp(`^\\d{1,3}(?:\\${sep}\\d{3})+$`).test(s);

/**
 * Tallet i en talstreng uden fortegn, eller undefined, hvis den ikke er et entydigt tal.
 * Både dansk (1.234,5) og engelsk (1,234.5) skrivemåde læses. "1.234" og "1,234" kan være både
 * tusinder og decimaler og er derfor ikke entydige.
 */
function parseUnsigned(b: string): number | undefined {
  // 15 cifre er, hvad et tal i Excel kan rumme uden at miste præcision (kortnumre m.m. skal blive tekst).
  if (/^\d+$/.test(b)) return leadingZeroOk(b) && b.length <= 15 ? Number(b) : undefined;
  const dots = b.split(".").length - 1;
  const commas = b.split(",").length - 1;
  let int: string;
  let frac = "";
  if (dots && commas) {
    // Det sidste tegn er decimaltegnet; det andet skal dele tusinder korrekt.
    const dec = b.lastIndexOf(".") > b.lastIndexOf(",") ? "." : ",";
    const th = dec === "." ? "," : ".";
    if (b.split(dec).length !== 2) return undefined;
    const [i, f] = b.split(dec);
    if (!grouped(i, th) || !/^\d+$/.test(f)) return undefined;
    int = i.split(th).join("");
    frac = f;
  } else {
    const sep = dots ? "." : ",";
    if (dots + commas > 1) {
      // 1.234.567: flere ens tegn kan kun være tusindtalsseparatorer.
      if (!grouped(b, sep)) return undefined;
      int = b.split(sep).join("");
    } else {
      const [i, f] = b.split(sep);
      if (!/^\d+$/.test(i) || !/^\d+$/.test(f)) return undefined;
      // "1.234": tusinder på dansk, decimaler på engelsk. "0,125" og "1234,567" kan kun være decimaler.
      if (f.length === 3 && i !== "0" && i.length <= 3) return undefined;
      int = i;
      frac = f;
    }
  }
  if (!leadingZeroOk(int) || int.length > 15) return undefined;
  const n = Number(frac ? `${int}.${frac}` : int);
  return Number.isFinite(n) ? n : undefined;
}

/**
 * Celleværdien for et svar: et tal, kun når hele svaret er et entydigt tal uden foranstillet nul.
 * Alt andet er tekst, så postnumre ("0045"), varenumre og "2300 København" ikke ødelægges.
 */
export function toCellValue(answer: string): string | number {
  const s = answer.trim();
  if (!s || s.length > 24 || !/^-?[\d.,]+$/.test(s)) return s;
  const neg = s.startsWith("-");
  const n = parseUnsigned(neg ? s.slice(1) : s);
  if (n === undefined) return s;
  return neg ? -n : n;
}
