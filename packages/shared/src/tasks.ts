/**
 * Opgaver (#60): typerne, de fælles prompt-byggesten, svar-skemaerne og selve kataloget ligger i
 * tasks/. Denne fil samler den offentlige del, så addin og bridge importerer som før.
 */
export * from "./tasks/definition.js";
export { doc, hostName, instructionsBlock, SYSTEM_BASE, sourcesBlock, system, withInstructions } from "./tasks/prompts.js";
export * from "./tasks/results.js";
export { getTask, TASKS, tasksForHost, textFormat } from "./tasks/catalog.js";
