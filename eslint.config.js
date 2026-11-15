// ESLint (flad konfiguration) for hele monorepoet (#14).
import js from "@eslint/js";
import reactHooks from "eslint-plugin-react-hooks";
import globals from "globals";
import tseslint from "typescript-eslint";

export default tseslint.config(
  { ignores: ["**/dist/**", "**/node_modules/**", "**/test-results/**", "**/demo-results/**", "**/playwright-report/**", "packages/app/build/**"] },
  js.configs.recommended,
  ...tseslint.configs.recommended,

  // Miljøer: browser for task pane, Node for bridge, app, scripts og tests.
  {
    files: ["packages/addin/src/**"],
    languageOptions: {
      globals: { ...globals.browser, Office: "readonly", Word: "readonly", Excel: "readonly", PowerPoint: "readonly", OfficeRuntime: "readonly" },
    },
  },
  {
    files: [
      "packages/bridge/**",
      "packages/shared/**",
      "packages/app/**",
      "packages/addin/e2e/**",
      "scripts/**",
      "*.config.{js,ts}",
      "packages/*/vite.config.ts",
    ],
    languageOptions: { globals: globals.node },
  },

  // Typebevidste regler for produktionskoden: et glemt await på Office.js eller fetch er en reel fejl.
  {
    files: ["packages/*/src/**/*.{ts,tsx}"],
    ignores: ["**/*.test.ts"],
    languageOptions: { parserOptions: { projectService: true, tsconfigRootDir: import.meta.dirname } },
    rules: {
      "@typescript-eslint/no-floating-promises": "error",
      "@typescript-eslint/no-misused-promises": ["error", { checksVoidReturn: { attributes: false } }],
    },
  },

  // React: hooks-regler.
  {
    files: ["packages/addin/src/**/*.tsx", "packages/addin/src/**/*.ts"],
    plugins: { "react-hooks": reactHooks },
    rules: {
      "react-hooks/rules-of-hooks": "error",
      "react-hooks/exhaustive-deps": "error",
    },
  },

  // HTML fra modellen må kun indsættes renset. Hver brug af dangerouslySetInnerHTML skal begrundes
  // med en eslint-disable-kommentar, der siger, hvorfor indholdet er sikkert.
  // Komponenterne må desuden ikke have hårdkodede farver eller pixelværdier (#17) – de hører
  // hjemme i src/theme.ts (Fluent-tokens eller temaets konstanter).
  {
    files: ["packages/addin/src/**/*.tsx"],
    ignores: ["packages/addin/src/theme.ts"],
    rules: {
      "no-restricted-syntax": [
        "error",
        {
          selector: "JSXAttribute[name.name='dangerouslySetInnerHTML']",
          message:
            "Indhold fra modellen skal renses (DOMPurify/escaping). Begrund med: // eslint-disable-next-line no-restricted-syntax -- <hvorfor det er sikkert>",
        },
        {
          selector: "Literal[value=/#[0-9a-fA-F]{3,8}\\b/]",
          message: "Ingen hårdkodede farver i komponenter – brug Fluent-tokens eller konstanter fra src/theme.ts.",
        },
        {
          selector: "Literal[value=/\\d+px\\b/], TemplateElement[value.raw=/\\d+px\\b/]",
          message: "Ingen hårdkodede pixelværdier i komponenter – brug tokens (spacing, radius) eller konstanter fra src/theme.ts (ui).",
        },
      ],
    },
  },

  // Tests må bruge any til JSON-svar og simulerede objekter.
  { files: ["**/*.test.ts", "**/test/**", "packages/addin/e2e/**"], rules: { "@typescript-eslint/no-explicit-any": "off" } },

  // Fælles: ubrugte variabler er fejl; _-præfiks markerer bevidst ubrugte.
  {
    rules: {
      "@typescript-eslint/no-unused-vars": ["error", { argsIgnorePattern: "^_", varsIgnorePattern: "^_", caughtErrors: "none" }],
    },
  },
);
