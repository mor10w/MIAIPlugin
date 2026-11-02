import type { HostAdapter } from "./types";
import { ExcelAdapter } from "./excel";
import { PowerPointAdapter } from "./powerpoint";
import { WordAdapter } from "./word";

export function createAdapter(host: Office.HostType): HostAdapter | undefined {
  switch (host) {
    case Office.HostType.Word:
      return new WordAdapter();
    case Office.HostType.Excel:
      return new ExcelAdapter();
    case Office.HostType.PowerPoint:
      return new PowerPointAdapter();
    default:
      return undefined;
  }
}

export type { HostAdapter } from "./types";
