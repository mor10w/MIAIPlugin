import { describe, expect, it } from "vitest";
import { pointSize, svgSize } from "./image.js";

describe("svgSize", () => {
  it("læser viewBox med negative og decimale tal (Mermaid)", () => {
    expect(svgSize('<svg width="100%" viewBox="-8 -8 512.5 300" style="max-width: 512px">')).toEqual({ w: 512.5, h: 300 });
  });
  it("læser komma-separeret viewBox", () => {
    expect(svgSize('<svg viewBox="0,0,200,100">')).toEqual({ w: 200, h: 100 });
  });
  it("falder tilbage til width/height", () => {
    expect(svgSize('<svg width="320" height="240">')).toEqual({ w: 320, h: 240 });
  });
});

describe("pointSize", () => {
  it("bruger den faktiske skala og bevarer forholdet", () => {
    // 400x200 px renderet i 2x → 800x400; naturlig bredde 400 px = 300 pt
    expect(pointSize({ base64: "", width: 800, height: 400, scale: 2 })).toEqual({ width: 300, height: 150 });
  });
  it("begrænser til A4's tekstbredde", () => {
    const p = pointSize({ base64: "", width: 2048, height: 1024, scale: 1 });
    expect(p.width).toBe(450);
    expect(p.height).toBe(225);
  });
});
