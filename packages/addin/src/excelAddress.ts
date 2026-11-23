/** Excel-adressegrammatik, fx "Ark1!A1:B2" eller bare "A1:E30" (med valgfri $-escapes). */
export const EXCEL_ADDRESS = /^(?:(?:'[^']+'|[^!'\s]+)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/i;
