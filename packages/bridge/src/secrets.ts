import { createRequire } from "node:module";

/**
 * Hemmeligheder (API-nøgler) i OS'ets nøglering (#16): Windows Credential Manager,
 * macOS Keychain eller Linux Secret Service. config.json gemmer kun en reference.
 */
export interface SecretStore {
  /** "keyring" når nøglerne ligger i OS'ets nøglering; "file" når de må ligge i config.json. */
  readonly kind: "keyring" | "file";
  get(account: string): string | null;
  set(account: string, value: string): void;
  delete(account: string): void;
}

/** Tjenestenavn i nøgleringen. */
export const KEYRING_SERVICE = "MIAI";

interface KeyringEntry {
  getPassword(): string | null;
  setPassword(password: string): void;
  deletePassword(): boolean;
}
type EntryCtor = new (service: string, account: string) => KeyringEntry;

/** Nøglering via @napi-rs/keyring (forbyggede binærer, ingen native build hos brugeren). */
export class KeyringSecrets implements SecretStore {
  readonly kind = "keyring";
  constructor(
    private readonly Entry: EntryCtor,
    private readonly service = KEYRING_SERVICE,
  ) {}
  get(account: string) {
    return new this.Entry(this.service, account).getPassword();
  }
  set(account: string, value: string) {
    new this.Entry(this.service, account).setPassword(value);
  }
  delete(account: string) {
    new this.Entry(this.service, account).deletePassword();
  }
}

/**
 * Ingen nøglering: nøglerne bliver i config.json (skrevet med rettighed 0600). Bruges, når
 * nøgleringen mangler (fx Linux uden Secret Service) eller er slået fra med MIAI_KEYRING=off.
 */
export class FileSecrets implements SecretStore {
  readonly kind = "file";
  get() {
    return null;
  }
  set() {
    throw new Error("Ingen nøglering");
  }
  delete() {}
}

/** Nøglering i hukommelsen – til tests. */
export class MemorySecrets implements SecretStore {
  readonly kind = "keyring";
  readonly data = new Map<string, string>();
  get(account: string) {
    return this.data.get(account) ?? null;
  }
  set(account: string, value: string) {
    this.data.set(account, value);
  }
  delete(account: string) {
    this.data.delete(account);
  }
}

/**
 * Vælger nøgleringen, hvis den kan bruges, ellers fil med en advarsel. Afprøves med et
 * opslag, så fx en låst eller manglende Secret Service opdages ved start og ikke ved første gem.
 */
export function defaultSecretStore(warn: (msg: string) => void = (m) => console.warn(m)): SecretStore {
  if (process.env.MIAI_KEYRING === "off") return new FileSecrets();
  try {
    const { Entry } = createRequire(import.meta.url)("@napi-rs/keyring") as { Entry: EntryCtor };
    const store = new KeyringSecrets(Entry);
    store.get("probe");
    return store;
  } catch (err) {
    warn(`[miai] OS'ets nøglering kan ikke bruges (${(err as Error).message}). API-nøgler gemmes i config.json (kun læsbar for dig).`);
    return new FileSecrets();
  }
}
