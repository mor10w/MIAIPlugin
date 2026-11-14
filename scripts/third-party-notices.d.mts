// Typer til third-party-notices.mjs, så vite.config.ts kan importere den.
export interface PackageNotice {
  name: string;
  version: string;
  license: string;
  author?: string;
  repository?: string;
  text: string;
}
export function packageRootOf(file: string): string | null;
export function describePackage(dir: string): PackageNotice | null;
export function collectPackages(dirs: Iterable<string>, options?: { exclude?: (name: string) => boolean }): PackageNotice[];
export function runtimeClosure(startDir: string): string[];
export function renderNotices(packages: PackageNotice[], intro: string): string;
