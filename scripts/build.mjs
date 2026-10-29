// Build wrapper: the sandbox denies esbuild's cleanup of its service temp
// dir under %TEMP% ("Access is denied" from esbuild's remove in the transpile
// step), so point TMP/TEMP at a workspace-local directory for the child.
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const tmp = mkdtempSync(join(root, ".tmp-build-"));
try {
  const result = spawnSync(process.execPath, [join("node_modules", "vite", "bin", "vite.js"), "build"], {
    stdio: "inherit",
    cwd: root,
    env: { ...process.env, TMP: tmp, TEMP: tmp },
  });
  process.exitCode = result.status ?? 1;
} finally {
  rmSync(tmp, { recursive: true, force: true });
}
