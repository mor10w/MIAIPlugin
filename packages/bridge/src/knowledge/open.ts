import { execFile } from "node:child_process";
import { platform } from "node:os";

/**
 * Åbner en fil i dens standardprogram (F15: "Åbn kilde"). Ingen shell – stien sendes som
 * argument, så specialtegn i filnavne ikke kan tolkes som kommandoer.
 */
export function openFile(path: string): Promise<void> {
  const [cmd, args] = platform() === "darwin" ? ["open", [path]] : platform() === "win32" ? ["explorer.exe", [path]] : ["xdg-open", [path]];
  return new Promise((ok, fail) => {
    const child = execFile(cmd, args as string[], { windowsHide: true }, (err) => {
      // explorer.exe returnerer 1, selv når den åbner filen.
      if (err && !(platform() === "win32" && err.code === 1)) fail(err);
      else ok();
    });
    child.unref();
  });
}
