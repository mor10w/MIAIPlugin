import { existsSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, beforeAll, beforeEach, describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import type { FastifyInstance } from "fastify";
import { type ExportSheetRequest } from "@miai/shared";
import { ConfigStore } from "../src/config.js";
import { buildServer } from "../src/server.js";
import { buildXlsx, exportPath } from "../src/export.js";

const PORT = 44321;
const H = { host: `localhost:${PORT}`, "x-miai": "1" };
let app: FastifyInstance;
let home: string;

beforeAll(() => {
  // MIAI_EXPORT_HOME, så testene aldrig skriver på den rigtige brugerprofil.
  home = mkdtempSync(join(tmpdir(), "mmai-export-"));
  process.env.MIAI_EXPORT_HOME = home;
});

afterAll(() => {
  delete process.env.MIAI_EXPORT_HOME;
});

beforeEach(async () => {
  const dir = mkdtempSync(join(tmpdir(), "miai-"));
  const store = new ConfigStore(dir, join(dir, "no-policy.json"));
  app = await buildServer({ store, port: PORT });
});

/** Et lille ark med værdier, en formel (med resultat), talformater og kolonnebredder. */
const payload = (): ExportSheetRequest => ({
  fileName: "Svar 2026.xlsx",
  folder: "desktop",
  name: "Data",
  address: "B2:D3",
  values: [
    ["Navn", 100, true],
    ["Anna", 200.5, null],
  ],
  formulas: [
    ["", "", ""],
    ["", "=10*20", ""],
  ],
  numberFormats: [
    ["General", "#,##0 kr", "General"],
    ["General", "0.0%", "General"],
  ],
  columnWidths: [12, 8.43],
});

/** Et ark med 50.011 celler (over MAX_EXPORT_CELLS). */
const big = (): ExportSheetRequest => ({
  ...payload(),
  fileName: "ForStor.xlsx",
  address: "A1:K5000",
  values: Array.from({ length: 5_000 }, () => Array(11).fill("x")),
  formulas: Array.from({ length: 5_000 }, () => Array(11).fill("")),
  numberFormats: Array.from({ length: 5_000 }, () => Array(11).fill("General")),
});

describe("exportPath", () => {
  it("afviser stipelementer og navne uden .xlsx-påliggende", () => {
    expect(() => exportPath("../evil.xlsx", "desktop")).toThrow();
    expect(() => exportPath("fil.txt", "desktop")).toThrow();
    expect(exportPath("Svar 2026.xlsx", "desktop")).toBe(join(home, "Desktop", "Svar 2026.xlsx"));
  });
});

describe("buildXlsx", () => {
  it("bygger filen med værdier, formel (med resultat), talformat og kolonnebredder", async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await buildXlsx(payload()));
    const ws = wb.getWorksheet("Data")!;
    expect(ws.getCell("B2").value).toBe("Navn"); // adressen starter i B2
    expect(ws.getCell("C2").value).toBe(100);
    expect(ws.getCell("D2").value).toBe(true); // booleans bevares
    const f = ws.getCell("C3").value as { formula: string; result: number };
    expect(f.formula).toBe("10*20"); // uden "="
    expect(f.result).toBe(200.5); // formelens nuværende resultat
    expect(ws.getCell("C3").numFmt).toBe("0.0%");
    expect(ws.getColumn(2).width).toBe(12); // B er den første kolonne (adressen starter i B)
  });

  it("afviser ugyldigt arknavn, inkonsekvent data og for mange celler", async () => {
    await expect(buildXlsx({ ...payload(), name: "a[1]" })).rejects.toThrow("Ugyldigt arknavn");
    await expect(buildXlsx({ ...payload(), formulas: [[""]] })).rejects.toThrow("Inkonsekvent data");
    await expect(buildXlsx(big())).rejects.toThrow("For mange celler");
  });
});

describe("POST /api/export", () => {
  it("skriver filen og svarer med stien; samme navn igen afvises (409)", async () => {
    const r = await app.inject({ method: "POST", url: "/api/export", headers: H, payload: payload() });
    expect(r.statusCode).toBe(200);
    const body = r.json() as { ok: boolean; path: string };
    expect(body.path).toBe(join(home, "Desktop", "Svar 2026.xlsx"));
    expect(existsSync(body.path)).toBe(true);

    const again = await app.inject({ method: "POST", url: "/api/export", headers: H, payload: payload() });
    expect(again.statusCode).toBe(409);
    expect(again.json().code).toBe("EXISTS");
  });

  it("skriver i den ønskede mappe (documents)", async () => {
    const r = await app.inject({ method: "POST", url: "/api/export", headers: H, payload: { ...payload(), fileName: "Salg.xlsx", folder: "documents" } });
    expect(r.statusCode).toBe(200);
    expect((r.json() as { path: string }).path).toBe(join(home, "Documents", "Salg.xlsx"));
  });

  it("afviser stipelementer (400), ugyldigt skema (400) og for mange celler (400)", async () => {
    const traversal = await app.inject({ method: "POST", url: "/api/export", headers: H, payload: { ...payload(), fileName: "../evil.xlsx" } });
    expect(traversal.statusCode).toBe(400);

    const bad = await app.inject({ method: "POST", url: "/api/export", headers: H, payload: { fileName: "x.xlsx" } });
    expect(bad.statusCode).toBe(400);

    const tooBig = await app.inject({ method: "POST", url: "/api/export", headers: H, payload: big() });
    expect(tooBig.statusCode).toBe(400);
  });
});
