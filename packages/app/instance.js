// Når MIAI startes, mens en udgave allerede kører (Electrons enkelt-instans-lås), lukker den nye
// sig selv. Før skete det uden et ord – en menulinje-app uden vindue ser så ud til ikke at reagere,
// og blev appen erstattet af en ny udgave, mens den gamle kørte, blev den nye aldrig startet.
// Her er den del, der ikke kræver Electron: hvad den kørende udgave skal sige og tilbyde.

/** Det, den nye instans sender med til den kørende (app.requestSingleInstanceLock(data)). */
export function launchInfo(version, execPath) {
  return { version: String(version), execPath: String(execPath) };
}

/** Hvor MIAI ligger, sagt så brugeren kan finde den. */
export function whereToFind(platform) {
  return platform === "win32" ? "i proceslinjen ved uret" : "i menulinjen øverst til højre";
}

/**
 * Dialogen i den kørende udgave, når en anden instans blev startet.
 * Samme udgave (eller ukendt): sig, hvor MIAI er. Anden udgave: tilbyd at skifte til den.
 */
export function secondInstanceDialog(current, other, platform) {
  const where = whereToFind(platform);
  const otherVersion = other && typeof other.version === "string" ? other.version : undefined;
  const otherPath = other && typeof other.execPath === "string" ? other.execPath : undefined;
  if (!otherVersion || otherVersion === current || !otherPath) {
    return {
      kind: "running",
      message: "MIAI kører allerede",
      detail: `Du finder MIAI ${where}. Klik på ikonet for menuen.`,
      buttons: ["OK"],
    };
  }
  return {
    kind: "switch",
    message: `Skift til MIAI ${otherVersion}?`,
    detail: `MIAI ${current} kører allerede, så ${otherVersion} kunne ikke starte. Skift lukker ${current} og starter ${otherVersion}.`,
    buttons: [`Skift til ${otherVersion}`, `Behold ${current}`],
    execPath: otherPath,
  };
}
