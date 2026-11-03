/**
 * Nodes indbyggede SQLite (node:sqlite, bruges af vidensindekset – F15) udskriver en
 * "ExperimentalWarning" ved hver start. Den forvirrer brugere i bridgens log, så netop den
 * advarsel skjules; alle andre advarsler vises som normalt. Skal importeres først.
 */
const original = process.emitWarning.bind(process);
process.emitWarning = ((warning: string | Error, ...rest: unknown[]) => {
  const msg = typeof warning === "string" ? warning : warning.message;
  if (msg.includes("SQLite is an experimental feature")) return;
  return (original as (...a: unknown[]) => void)(warning, ...rest);
}) as typeof process.emitWarning;

export {};
