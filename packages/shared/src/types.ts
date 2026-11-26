import { z } from "zod";
import { Attachment, MAX_ATTACHMENTS } from "./attachments.js";

export const HostKind = z.enum(["word", "excel", "powerpoint", "outlook", "unknown"]);
export type HostKind = z.infer<typeof HostKind>;

/** Det udsnit af dokumentet, som en opgave arbejder på. Bygges af add-in'ets host-adapter. */
export const DocContext = z.object({
  host: HostKind,
  /** Markeret tekst (eller hele dokumentet, hvis intet er markeret og opgaven tillader det). */
  text: z.string().max(200_000).default(""),
  /** true hvis `text` er hele dokumentet frem for en markering. */
  isWholeDocument: z.boolean().default(false),
  /** Excel: adressen på markeringen, fx "Ark1!C5". */
  address: z.string().optional(),
  /** Excel: formlen i markeringens øverste venstre celle. */
  formula: z.string().optional(),
  /** Excel: markeringen som TSV (værdier), begrænset i størrelse. */
  table: z.string().max(100_000).optional(),
  /** Excel: oversigt over filen – arkliste, headers og kolonne-statistik regnet over hele used range (bygget af add-in'et). */
  surroundings: z.string().max(20_000).optional(),
  /**
   * Agent i flere trin (#20): det, agenten selv har bedt om at læse i dokumentet (fx celler på et
   * andet ark), hentet af add-in'et uden at ændre noget. Vises for modellen i næste runde.
   */
  observations: z.string().max(40_000).optional(),
  /** Sat af add-in'et, hvis teksten blev afkortet, før den blev sendt (#3). */
  truncated: z.boolean().optional(),
  /** Tekstens oprindelige længde i tegn, når den er afkortet. */
  originalChars: z.number().int().nonnegative().optional(),
  /**
   * Word, agent-tilstand: dokumentets afsnit med typografi, så modellen kan se (og bevare)
   * overskrifter, lister og tabeller. Nummereres fra 1 i prompten. Udeladt i ældre Word.
   */
  paragraphs: z
    .array(
      z.object({
        text: z.string().max(20_000),
        /** Indbygget typografi (fx "Heading1", "ListBullet", "Normal") eller dokumentets eget navn. */
        style: z.string().max(80),
        list: z.boolean().optional(),
        table: z.boolean().optional(),
        /** Afsnittet står i en tabelcelle (#24): tabellens nummer i dokumentet, række og kolonne (1 = første). */
        cell: z.object({ table: z.number().int().min(1), row: z.number().int().min(1), col: z.number().int().min(1) }).optional(),
      }),
    )
    .max(5_000)
    .optional(),
  /**
   * Excel, agent-tilstand: tabeller (ListObjects) på arket og kolonnernes talformat, så nye
   * rækker og celler kan få samme formatering som data.
   */
  excel: z
    .object({
      tables: z
        .array(z.object({ name: z.string().max(255), address: z.string().max(100), headers: z.array(z.string().max(255)).max(100) }))
        .max(50)
        .default([]),
      /** Talformat pr. kolonne i used range (fx "A: General", "C: #,##0.00 kr"). */
      columnFormats: z.array(z.string().max(120)).max(100).default([]),
      /** Det aktive ark (#21): adresser uden arknavn gælder dette ark. */
      activeSheet: z.string().max(255).optional(),
      /** Alle ark i projektmappen (#21) med brugt område og overskrifter, så agenten kan arbejde på tværs af ark. */
      sheets: z
        .array(z.object({ name: z.string().max(255), used: z.string().max(100).optional(), headers: z.array(z.string().max(255)).max(100).default([]) }))
        .max(50)
        .optional(),
    })
    .optional(),
  /**
   * PowerPoint, agent-tilstand: slides med layout og figurer (som "2.1"), og præsentationens
   * layouts, så nye slides kan oprettes fra skabelonen i stedet for med egne tekstbokse.
   */
  slides: z
    .array(
      z.object({
        layout: z.string().max(120),
        shapes: z
          .array(
            z.object({
              name: z.string().max(255),
              /** title/body = pladsholdere fra layoutet; text = egen tekstboks; other = figur uden tekst. */
              kind: z.enum(["title", "body", "text", "other"]),
              text: z.string().max(10_000),
            }),
          )
          .max(100),
      }),
    )
    .max(300)
    .optional(),
  layouts: z.array(z.string().max(120)).max(100).optional(),
});
export type DocContext = z.infer<typeof DocContext>;

export const ChatMessage = z.object({
  role: z.enum(["system", "user", "assistant"]),
  content: z.string(),
});
export type ChatMessage = z.infer<typeof ChatMessage>;

/**
 * Uddrag fra brugerens egne filer (F15, lokal RAG – ADR-0013). Sættes af bridgen ud fra
 * vidensindekset; det, en klient selv sender, overskrives i bridge-tilstand.
 */
export const Source = z.object({
  /** Nummeret, modellen citerer med, fx [1]. */
  n: z.number().int().min(1).max(50),
  title: z.string().max(300),
  path: z.string().max(2_000),
  text: z.string().max(6_000),
});
export type Source = z.infer<typeof Source>;

/** Kildehenvisning, som vises under svaret (uden hele teksten). */
export interface SourceRef {
  n: number;
  title: string;
  path: string;
  /** Chunk-id i bridgens indeks – til at hente uddraget og åbne filen. */
  chunkId: number;
  /** Lighed mellem spørgsmål og uddrag (0–1). */
  score: number;
  snippet: string;
}

/** Vidensindeksets tilstand (F15), som bridgen viser under Indstillinger. */
export interface KnowledgeStatus {
  state: "disabled" | "idle" | "indexing" | "error";
  /** Under indeksering: filer behandlet / i alt. */
  done: number;
  total: number;
  documents: number;
  chunks: number;
  /** Filer, der ikke kunne læses (fx beskyttede PDF'er). */
  failed: { path: string; error: string }[];
  /** Mapper, der ikke findes eller ikke er tilladt af politikken. */
  skippedFolders: { folder: string; reason: "missing" | "policy" }[];
  embeddingModel: string;
  lastIndexed?: string;
  /** Fejl, der stoppede indekseringen (fx endpointen understøtter ikke embeddings). */
  error?: string;
  /** Der var flere filer end MAX_FILES; resten er ikke indekseret. */
  truncated?: boolean;
}

export const TaskRequest = z.object({
  context: DocContext,
  /** Opgaveparametre, fx { targetLanguage: "engelsk" } eller { tone: "formel" }. */
  params: z.record(z.string(), z.string().max(4_000)).default({}),
  /** Kun for chat: tidligere beskeder i samtalen. */
  history: z.array(ChatMessage).max(100).default([]),
  /** Overstyr model for denne forespørgsel. */
  model: z.string().optional(),
  /** Chat: uddrag fra brugerens egne mapper (F15). Udfyldes af bridgen. */
  sources: z.array(Source).max(20).default([]),
  /** Billeder og tekst fra filer, brugeren har vedhæftet beskeden (#27). */
  attachments: z.array(Attachment).max(MAX_ATTACHMENTS).default([]),
});
export type TaskRequest = z.input<typeof TaskRequest>;
export type ParsedTaskRequest = z.output<typeof TaskRequest>;

/** Hændelser bridgen sender som Server-Sent Events. */
export type TaskEvent =
  | { type: "meta"; model: string; endpoint: string }
  | { type: "delta"; text: string }
  /** Fremdrift for strukturerede (JSON) svar, som ikke vises løbende. */
  | { type: "progress"; chars: number; step?: { done: number; total: number } }
  /** Fjern de sidste `chars` tegn fra den viste tekst (fx indledning fra en runde, der endte med en søgning). */
  | { type: "retract"; chars: number }
  /** Kilder fra brugerens mapper, som svaret bygger på (F15). Sendes før svaret. */
  | { type: "sources"; sources: SourceRef[] }
  /** Modellen søger på webbet (vises i UI, fx "Søger på web…"). */
  | { type: "search"; query: string }
  | { type: "result"; data: unknown }
  | { type: "error"; message: string; code?: string; status?: number }
  | {
      type: "done";
      usage?: { promptTokens?: number; completionTokens?: number };
      /** Sat, hvis motoren afbrød svaret, fx fordi modellen gik i en gentagelsessløjfe. */
      stopReason?: "repetition";
      /** input-truncated: modellen så mindre, end der blev sendt (kontekstgrænse). output-truncated: svaret ramte max_tokens. */
      warnings?: TaskWarning[];
      /** Ved small-context: modellens kontekst og værktøjet, så UI'et kan vise den rigtige rettelse (#73). */
      context?: { tokens: number; server?: "ollama" | "lmstudio" };
    };

// ---------------------------------------------------------------------------
// Indstillinger og politik
// ---------------------------------------------------------------------------

/**
 * thinking-ignored: serveren eller modellen afviste indstillingen for tænkning/indsats, så
 * forespørgslen blev sendt uden den (svaret er lavet med modellens standard).
 * small-context: agent eller chat kørte på en model med under 16k tokens kontekst (#73).
 */
export type TaskWarning = "input-truncated" | "output-truncated" | "thinking-ignored" | "small-context";

/**
 * Søgetjeneste til web-søgning (#12). duckduckgo = offentlig, uden opsætning (søgeord sendes til
 * DuckDuckGo). searxng = selvhostet metasøgemaskine (fx organisationens egen), så søgeord ikke
 * sendes direkte til en tredjepart. exa = Exas hostede MCP-server, uden konto og nøgle (søgeord
 * sendes til Exa); den giver uddrag af selve siderne og har en gratis grænse for antal søgninger.
 */
export const SearchConfig = z.object({
  provider: z.enum(["duckduckgo", "searxng", "exa"]).default("duckduckgo"),
  /** SearXNG: basisadresse, fx https://search.firma.local */
  url: z.string().url().optional(),
});
export type SearchConfig = z.infer<typeof SearchConfig>;

export const ThinkingMode = z.enum(["default", "on", "off"]);
export type ThinkingMode = z.infer<typeof ThinkingMode>;
export const ReasoningEffort = z.enum(["default", "low", "medium", "high"]);
export type ReasoningEffort = z.infer<typeof ReasoningEffort>;

export const StructuredOutputMode = z.enum(["json_schema", "json_object", "none"]);
export type StructuredOutputMode = z.infer<typeof StructuredOutputMode>;

export const Endpoint = z.object({
  id: z.string().regex(/^[a-z0-9-]{1,40}$/),
  name: z.string().min(1).max(80),
  /** Basis-URL inkl. /v1, fx http://127.0.0.1:11434/v1 */
  baseUrl: z.string().url(),
  apiKey: z.string().max(500).default(""),
  structuredOutput: StructuredOutputMode.default("json_schema"),
  stripThinking: z.boolean().default(true),
  /**
   * OpenAI `presence_penalty` (0–2) mod gentagelsessløjfer – Qwen anbefaler 1.0–1.5 for
   * kvantiserede modeller. Udeladt = sendes ikke (serverens standard). Bruges kun til
   * fritekst-opgaver; JSON og Excel-celler har legitime gentagelser.
   */
  presencePenalty: z.number().min(0).max(2).optional(),
  /**
   * Modellens kontekstlængde i tokens (#3). Udeladt = aflæses automatisk: fra /v1/models (vLLM
   * m.fl.) eller værktøjets eget API (Ollama, LM Studio; #73). Ellers ukendt, og der planlægges
   * uden loft.
   */
  contextTokens: z.number().int().min(512).max(2_000_000).optional(),
  /**
   * Ekstra felter, der flettes ind i hver chat/completions-forespørgsel (fx { "top_k": 20 } til
   * llama.cpp). Bemærk: Ollamas OpenAI-lag ignorerer num_ctx her; konteksten sættes i Ollama selv.
   */
  extraBody: z.record(z.string(), z.unknown()).default({}),
  /**
   * Hvor længe serveren må være tavs, før forespørgslen opgives: før svaret begynder (fx mens en
   * lang prompt behandles på CPU) og mellem stykkerne i et svar. Et svar, der strømmer, afbrydes
   * aldrig, uanset hvor langt det er. 5 min som standard; op til 30 min for langsom hardware.
   */
  timeoutMs: z.number().int().min(1_000).max(1_800_000).default(300_000),
});
export type Endpoint = z.infer<typeof Endpoint>;

export const Settings = z.object({
  endpoints: z.array(Endpoint).min(1),
  activeEndpoint: z.string(),
  /** Standardmodel. Tom = første model, endpointen rapporterer. */
  defaultModel: z.string().default(""),
  /**
   * Model pr. opgavetype (modelGroups.ts), fx { "edit": "qwen3:4b", "agent": "qwen3:32b" }.
   * Typer uden valg bruger standardmodellen.
   */
  groupModels: z.record(z.string(), z.string()).default({}),
  /** Undtagelser for enkelte opgaver, fx { "explain-formula": "qwen3:14b" }. Har forrang for typen. */
  taskModels: z.record(z.string(), z.string()).default({}),
  temperature: z.number().min(0).max(2).default(0.3),
  maxTokens: z.number().int().min(16).max(65_536).default(8_192),
  /**
   * Tænkning (reasoning) for modeller, der kan tænke, før de svarer (Qwen3, DeepSeek-R1, gpt-oss
   * m.fl.). default = send intet, så serveren og modellen bestemmer; off = hurtigere svar uden
   * tænkning; on = slå den til, hvor modellen kan begge dele. Se ADR-0003, tillæg.
   */
  thinking: ThinkingMode.default("default"),
  /** Hvor meget modellen tænker (OpenAI reasoning_effort). default = send intet. */
  reasoningEffort: ReasoningEffort.default("default"),
  uiLanguage: z.enum(["da", "en"]).default("da"),
  auditLog: z.boolean().default(false),
  /**
   * Hvor mange celler regnearksfunktionen =MIAI.ASK() må sende til modellen på samme tid (#80).
   * 2 passer til en lokal model; en server med flere pladser kan tage flere.
   */
  cellConcurrency: z.number().int().min(1).max(16).default(2),
  /**
   * Modellen kan søge på webbet (kræver bridge – direkte tilstand har ingen eksekverer).
   * Slået FRA som standard: søgeordene dannes af modellen ud fra dokumentet og sendes til en
   * ekstern søgetjeneste, så det skal være et aktivt valg (løftet om, at data bliver lokalt).
   */
  webSearch: z.boolean().default(false),
  /** Hvilken søgetjeneste web-søgningen bruger (#12). */
  search: SearchConfig.default({ provider: "duckduckgo" }),
  /** Lokal RAG over brugerens mapper (F15, ADR-0013). Kun bridge. */
  knowledge: z
    .object({
      enabled: z.boolean().default(false),
      /** Mapper, der indekseres (absolutte stier på maskinen). */
      folders: z.array(z.string().min(1).max(1_024)).max(20).default([]),
      /** Embedding-model på den aktive endpoint (/v1/embeddings), fx "bge-m3". */
      embeddingModel: z.string().max(200).default(""),
      /** Antal uddrag, der gives til modellen pr. spørgsmål. */
      topK: z.number().int().min(1).max(12).default(5),
    })
    .default({ enabled: false, folders: [], embeddingModel: "", topK: 5 }),
  /**
   * Kodetilstand (ADR-0017): agenten må skrive Office.js-kode (Excel), som køres i en sandkasse
   * efter brugerens godkendelse. Slået FRA som standard og kan forbydes af politikken.
   */
  agentScripts: z.boolean().default(false),
  /**
   * Chatten husker samtalen pr. dokument (#26) i task panes lager på maskinen. Kan slås fra, hvis
   * samtaler om dokumenter ikke må ligge uden for dokumentet.
   */
  chatHistory: z.boolean().default(true),
  /**
   * Agenten kan gemme et Excel-ark som .xlsx-fil (tillæg ADR-0012). Kræver bridge – direkte
   * tilstand har ingen eksekverer, der kan skrive filer. Slået FRA som standard og kan forbydes
   * af politikken, så organisationen styrer, om agenten må skrive filer på maskinen.
   */
  agentExport: z.boolean().default(false),
  /** Mappe med brugerdefinerede opgaver (F14). Kun bridge kan læse mapper; tom = <config>/tasks. */
  customTasksDir: z.string().max(1024).default(""),
});
export type Settings = z.infer<typeof Settings>;

export const Policy = z.object({
  /** Hostnavne/IP'er bridgen må kalde. Understøtter "*.domæne" wildcard. */
  allowedHosts: z.array(z.string()).default(["127.0.0.1", "localhost", "::1"]),
  /** Låst politik kan ikke ændres via UI (sat af organisationen). */
  locked: z.boolean().default(false),
  /** Må web-søgning slås til? En organisation kan forbyde den helt (søgeord forlader maskinen). */
  webSearch: z.boolean().default(true),
  /** Låser søgetjenesten, fx til organisationens egen SearXNG (#12). Har forrang for brugerens valg. */
  search: SearchConfig.optional(),
  /** Må lokal RAG (F15) bruges? */
  knowledge: z.boolean().default(true),
  /**
   * Mapper, der må indekseres (og alt under dem). Udeladt = alle mapper brugeren vælger.
   * En organisation kan fx begrænse til et fælles drev.
   */
  knowledgeFolders: z.array(z.string()).optional(),
  /** Må agentens kodetilstand (Office.js-scripts, ADR-0017) slås til? */
  agentScripts: z.boolean().default(true),
  /** Må agenten gemme et Excel-ark som .xlsx-fil (tillæg ADR-0012)? */
  agentExport: z.boolean().default(true),
});
export type Policy = z.infer<typeof Policy>;

/** Kodetilstanden er kun aktiv, når brugeren har slået den til, og politikken tillader den. */
export function agentScriptsAllowed(settings: Pick<Settings, "agentScripts">, policy: Pick<Policy, "agentScripts">): boolean {
  return !!settings.agentScripts && policy.agentScripts !== false;
}

/** Eksporten (tillæg ADR-0012) er kun aktiv, når brugeren har slået den til, og politikken tillader den. */
export function agentExportAllowed(settings: Pick<Settings, "agentExport">, policy: Pick<Policy, "agentExport">): boolean {
  return !!settings.agentExport && policy.agentExport !== false;
}

/** Den søgetjeneste, der faktisk bruges: politikkens, hvis den er låst, ellers brugerens. */
export function effectiveSearch(settings: Settings, policy: Policy): SearchConfig {
  return policy.search ?? settings.search;
}

/**
 * Organisationens konfiguration til et centralt udrullet, hostet add-in (#7, ADR-0015): filen
 * `miai-config.json` ligger ved siden af taskpane.html på add-in'ets egen adresse og hentes ved
 * start. Politikken har forrang og kan ikke ændres i add-in'et; indstillingerne er standard for
 * brugere, der ikke selv har gemt nogen (fx organisationens interne inference-server).
 */
export const OrgConfig = z.object({
  policy: Policy.optional(),
  settings: z.record(z.string(), z.unknown()).optional(),
});
export type OrgConfig = z.infer<typeof OrgConfig>;
export const ORG_CONFIG_FILE = "miai-config.json";

/** Det, add-in'et får fra GET /api/settings (API-nøgler maskeres). */
export interface SettingsView {
  settings: Settings;
  policy: Policy;
  /** browser = gemt i add-in'et (direkte tilstand); machine = låst af organisationen (kun bridge). */
  policySource: "default" | "user" | "machine" | "browser" | "organization";
}

export interface ModelInfo {
  id: string;
  endpoint: string;
}

/**
 * Browserens tilladelse til lokal netværksadgang (Local Network Access, #74), når den står i vejen
 * for add-in'et i direkte tilstand: denied = afvist af brugeren eller en browserpolitik; frame =
 * siden, add-in'et vises i (fx Office på nettet), giver ikke tilladelsen videre til add-in'et;
 * prompt = browseren har ikke fået svar endnu, så et kald kan vente på (eller være afvist af) spørgsmålet.
 */
export type LocalNetworkAccess = "denied" | "frame" | "prompt";

export interface HealthInfo {
  ok: boolean;
  version: string;
  /** blocked: browseren blokerer kaldet (kun direkte tilstand, #74). */
  endpoint: { id: string; baseUrl: string; reachable: boolean; error?: string; blocked?: Exclude<LocalNetworkAccess, "prompt"> };
}

/**
 * Resultat af at teste en endpoint (#4). reason: cors = noget svarer, men browseren må ikke læse
 * svaret (CORS mangler); absent = intet svarer; policy = ikke tilladt; http = fejlsvar fra serveren;
 * blocked = browseren blokerer kaldet (lokal netværksadgang, #74), så det kan ikke ses, om noget svarer.
 * lna er sat ved blocked (denied/frame) og ved absent, når browseren endnu ikke har fået svar (prompt).
 */
export type ProbeResult =
  | { ok: true; models: string[] }
  | { ok: false; reason: "cors" | "absent" | "policy" | "http" | "protocol" | "blocked"; message: string; lna?: LocalNetworkAccess };
