import { describe, expect, it } from "vitest";
import { parseDdgHtml } from "../src/search.js";

// Uddrag af DuckDuckGo's HTML-resultatside (html.duckduckgo.com/html/).
const FIXTURE = `
<html><body>
<div class="result results_links web-result">
  <div class="links_main links_deep result__body">
    <h2 class="result__title">
      <a rel="nofollow" class="result__a" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fwww.eksempel.dk%2Fnyhed&rut=abc">Eksempel &amp; Co slår til</a>
    </h2>
    <a class="result__snippet" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fwww.eksempel.dk%2Fnyhed&rut=abc">Her er <b>snippet</b> med &lt;tegn&gt; og mere tekst.</a>
  </div>
</div>
<div class="result results_links web-result">
  <div class="links_main links_deep result__body">
    <h2 class="result__title">
      <a rel="nofollow" class="result__a" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fanden.dk&rut=xyz">Anden side</a>
    </h2>
    <a class="result__snippet" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fanden.dk&rut=xyz">Toet snippet.</a>
  </div>
</div>
</body></html>`;

describe("parseDdgHtml", () => {
  it("finder titler, URLs og snippets (parvis i rækkefølge)", () => {
    const results = parseDdgHtml(FIXTURE);
    expect(results).toHaveLength(2);
    expect(results[0].title).toBe("Eksempel & Co slår til");
    expect(results[0].url).toBe("https://www.eksempel.dk/nyhed");
    expect(results[0].snippet).toContain("Her er snippet med <tegn> og mere tekst.");
    expect(results[1].title).toBe("Anden side");
    expect(results[1].url).toBe("https://anden.dk");
  });

  it("tåler manglende snippet", () => {
    const html = '<a rel="nofollow" class="result__a" href="//duckduckgo.com/l/?uddg=https%3A%2F%2Fx.dk">Kun titel</a>';
    const results = parseDdgHtml(html);
    expect(results).toHaveLength(1);
    expect(results[0].title).toBe("Kun titel");
    expect(results[0].snippet).toBe("");
  });

  it("returnerer tom liste uden resultater", () => {
    expect(parseDdgHtml("<html><body>ingen resultater</body></html>")).toHaveLength(0);
  });

  it("filtrerer resultater uden URL", () => {
    const html = '<a rel="nofollow" class="result__a" href="">Kun titel</a>';
    expect(parseDdgHtml(html)).toHaveLength(0);
  });
});
