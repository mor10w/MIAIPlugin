import type { Policy } from "@miai/shared";

/** Tjekker om en URL's host er tilladt af politikken. Understøtter "*.domæne" wildcards. */
export function isUrlAllowed(url: string, policy: Policy): boolean {
  let host: string;
  try {
    host = new URL(url).hostname.toLowerCase().replace(/^\[|\]$/g, "");
  } catch {
    return false;
  }
  return policy.allowedHosts.some((pattern) => {
    const p = pattern.toLowerCase().trim().replace(/^\[|\]$/g, "");
    if (p === "*") return true;
    if (p.startsWith("*.")) return host.endsWith(p.slice(1)) && host.length > p.length - 1;
    return host === p;
  });
}
