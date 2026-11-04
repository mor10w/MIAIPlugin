import { applyEdits, type ApplyMode, type DocContext, type GrammarResult } from "@miai/shared";
import { mdToPlain, mdToSafeHtml } from "../markdown";
import { truncate, type ApplyOutcome, type ContextOptions, type HostAdapter, type TextFormat } from "./types";

/**
 * Outlook (#7, ADR-0015): mailen i læsetilstand (opsummér, svar, chat om mailen) og udkastet i
 * skrivetilstand (omskriv, ret, oversæt markeringen eller hele teksten). Mailbox-API'et er
 * callback-baseret; hjælperne nedenfor gør det til promises.
 */

type Item = Office.MessageRead & Office.MessageCompose;
const item = () => Office.context.mailbox.item as unknown as Item;

function call<T>(fn: (cb: (r: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) =>
    fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(new Error(r.error?.message ?? "Outlook afviste handlingen.")))),
  );
}

/** Læsetilstand har svar-formularer; skrivetilstand har ikke. */
const isRead = () => typeof item().displayReplyForm === "function";

const escapeHtml = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
/** Ren tekst som HTML med bevarede linjeskift. */
const plainToHtml = (s: string) => escapeHtml(s).replace(/\n/g, "<br>");

export class OutlookAdapter implements HostAdapter {
  readonly host = "outlook" as const;

  supports(mode: ApplyMode): boolean {
    if (isRead()) return mode === "reply" || mode === "replyAll";
    return mode === "replace" || mode === "insertAfter";
  }

  async getContext(opts: ContextOptions): Promise<DocContext> {
    const it = item();
    if (isRead()) {
      const body = await call<string>((cb) => it.body.getAsync(Office.CoercionType.Text, cb));
      const from = it.from ? `${it.from.displayName} <${it.from.emailAddress}>` : "";
      const text = [`Subject: ${it.subject ?? ""}`, ...(from ? [`From: ${from}`] : []), "", body].join("\n");
      return { host: "outlook", ...truncate(text, opts.maxChars), isWholeDocument: true };
    }
    // Skrivetilstand: markeringen, ellers (hvis opgaven tillader det) hele udkastet.
    const selected = await call<{ data: string }>((cb) => it.getSelectedDataAsync(Office.CoercionType.Text, cb)).catch(() => ({ data: "" }));
    if (selected.data.trim() || !opts.wholeDocumentFallback) return { host: "outlook", ...truncate(selected.data, opts.maxChars), isWholeDocument: false };
    const body = await call<string>((cb) => it.body.getAsync(Office.CoercionType.Text, cb));
    return { host: "outlook", ...truncate(body, opts.maxChars), isWholeDocument: true };
  }

  /** Udkastets format: HTML (de fleste mails) eller ren tekst. */
  private async bodyIsHtml(): Promise<boolean> {
    const type = await call<Office.CoercionType>((cb) => item().body.getTypeAsync(cb)).catch(() => Office.CoercionType.Html);
    return type === Office.CoercionType.Html;
  }

  async applyText(text: string, format: TextFormat, mode: ApplyMode, dc: DocContext): Promise<ApplyOutcome> {
    const html = format === "markdown" ? mdToSafeHtml(text) : plainToHtml(text);
    const it = item();
    if (mode === "reply" || mode === "replyAll") {
      // Svaret åbnes som et udkast – brugeren ser, retter og sender selv.
      if (mode === "replyAll") it.displayReplyAllForm({ htmlBody: html });
      else it.displayReplyForm({ htmlBody: html });
      return { message: "Svaret er åbnet som et udkast." };
    }
    const asHtml = await this.bodyIsHtml();
    const content = asHtml ? html : format === "markdown" ? mdToPlain(text) : text;
    const coercionType = asHtml ? Office.CoercionType.Html : Office.CoercionType.Text;
    if (mode === "replace" && dc.isWholeDocument) {
      await call<void>((cb) => it.body.setAsync(content, { coercionType }, cb));
      return {};
    }
    if (mode === "insertAfter" && dc.text.trim()) {
      // Efter markeringen: markeringen (med sin formatering) + den nye tekst.
      const sel = await call<{ data: string }>((cb) => it.getSelectedDataAsync(coercionType, cb));
      await call<void>((cb) => it.body.setSelectedDataAsync(sel.data + (asHtml ? "<br>" : "\n") + content, { coercionType }, cb));
      return {};
    }
    // Erstat markeringen – eller indsæt ved markøren, når intet er markeret.
    await call<void>((cb) => it.body.setSelectedDataAsync(content, { coercionType }, cb));
    return {};
  }

  async applyGrammar(result: GrammarResult, _mode: ApplyMode, dc: DocContext): Promise<ApplyOutcome> {
    const { text, skipped } = applyEdits(dc.text, result.edits);
    await this.applyText(text, "plain", "replace", dc);
    const n = result.edits.length - skipped;
    return { message: `${n} af ${result.edits.length} rettelser anvendt.`, warning: skipped > 0 };
  }
}
