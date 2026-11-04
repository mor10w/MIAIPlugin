/** Formater, der gemmes til fortryd (ExcelApi 1.9). Kanter med, fordi tabeltypografier kan sætte dem. */
export const FORMAT_PROPS = {
  format: {
    fill: { color: true },
    font: { bold: true, italic: true, color: true, name: true, size: true, underline: true, strikethrough: true },
    borders: { style: true, color: true, weight: true },
    horizontalAlignment: true,
    verticalAlignment: true,
    wrapText: true,
    indentLevel: true,
  },
} as Excel.CellPropertiesLoadOptions;
