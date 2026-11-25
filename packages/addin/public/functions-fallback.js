// Regnearksfunktionerne =MIAI.ASK() og =MIAI.MODELS() (#80, ADR-0018) kører i samme side som task
// pane (shared runtime). I en Excel uden shared runtime indlæser Excel i stedet denne fil i en ren
// JavaScript-runtime uden localStorage og DOM, hvor MIAI's indstillinger og politik ikke kan nås.
// Dér svarer funktionerne med en forklaring i stedet for at hænge på #BUSY!.
/* global CustomFunctions */
(function () {
  function unsupported() {
    var message = "MIAI.ASK kræver en nyere Excel (shared runtime). / MIAI.ASK requires a newer Excel (shared runtime).";
    if (CustomFunctions.Error && CustomFunctions.ErrorCode) throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable, message);
    throw new Error(message);
  }
  CustomFunctions.associate("ASK", unsupported);
  CustomFunctions.associate("MODELS", unsupported);
})();
