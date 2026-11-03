import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // Tests må ikke røre udviklerens rigtige nøglering (#16); keyring-testene bruger MemorySecrets.
    env: { MIAI_KEYRING: "off" },
  },
});
