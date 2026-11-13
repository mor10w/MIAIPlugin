/**
 * Markeringen i dokumentet (#26): fælles hjælpere til Word, Excel og PowerPoint, som alle sender
 * DocumentSelectionChanged. Outlook har intet dokument og får ingen besked.
 */

/** Kalder `cb`, hver gang brugeren markerer noget andet. Returnerer en funktion, der stopper lytningen. */
export function watchSelection(cb: () => void): () => void {
  const handler = () => cb();
  try {
    Office.context.document?.addHandlerAsync?.(Office.EventType.DocumentSelectionChanged, handler);
  } catch {
    // Værten understøtter ikke hændelsen: markeringen vises så kun, når chatten åbnes eller sender.
  }
  return () => {
    try {
      Office.context.document?.removeHandlerAsync?.(Office.EventType.DocumentSelectionChanged, { handler });
    } catch {
      // Intet at rydde op.
    }
  };
}

/** Den markerede tekst (Word og PowerPoint), eller "" hvis intet er markeret. */
export function selectedText(): Promise<string> {
  return new Promise((resolve) => {
    try {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (r) =>
        resolve(r.status === Office.AsyncResultStatus.Succeeded ? String(r.value ?? "") : ""),
      );
    } catch {
      resolve("");
    }
  });
}

/** Antal ord i en tekst (til "12 ord markeret"). */
export const wordCount = (text: string): number => text.trim().split(/\s+/).filter(Boolean).length;

/** Dokumentets adresse (sti eller URL). Tom for dokumenter, der ikke er gemt endnu. */
export function documentUrl(): string {
  try {
    return Office.context.document?.url ?? "";
  } catch {
    return "";
  }
}
