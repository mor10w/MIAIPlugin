/**
 * Excel-adressegrammatik, fx "Ark1!A1:B2" eller bare "A1:E30" (med valgfri $-escapes).
 * En apostrof i et citeret arknavn skrives som to, som Excel gør: "'Bob''s'!A1".
 */
export const EXCEL_ADDRESS = /^(?:(?:'(?:[^']|'')+'|[^!'\s]+)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/i;

/** "'Mit ark'!A1:B2" → { sheet: "Mit ark", range: "A1:B2" } */
export function splitAddress(address: string): { sheet: string; range: string } {
  const i = address.lastIndexOf("!");
  if (i < 0) return { sheet: "", range: address };
  const sheet = address
    .slice(0, i)
    .replace(/^'(.*)'$/, "$1")
    .replace(/''/g, "'");
  return { sheet, range: address.slice(i + 1) };
}
