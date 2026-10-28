/**
 * The one place that talks to the Go sidecar.
 *
 * The host bridge is absent when the page is opened directly in a browser, so
 * every call reports that instead of throwing: the calendar can then still
 * render with the data it computes locally.
 */
const bridge = typeof window !== "undefined" ? window.dbxPlugin : undefined;

export const bridgeAvailable = Boolean(bridge);

/**
 * Host locale, or null outside a host.
 *
 * Null rather than a guess: a page opened straight from the browser has no
 * locale to report, and the caller is the one that knows which language to
 * fall back to.
 */
export function hostLocale() {
  return bridge?.locale ?? bridge?.context?.locale ?? null;
}

/**
 * The host theme snapshot carried on the bridge (`dbxPlugin.theme`), or null
 * outside a host. Its `tokens` map carries the resolved design variables the
 * host pushes with every theme change — `--font-sans` among them — so the
 * calendar can follow the host's font setting instead of hardcoding a stack.
 */
export function hostTheme() {
  return bridge?.theme ?? null;
}

/**
 * Subscribes to host env pushes (`dbx-plugin-env`), the channel the host uses
 * to announce locale and theme changes. Returns an unsubscribe function; in a
 * plain browser tab there is nothing to subscribe to and this is a no-op.
 */
export function onHostEnv(handler) {
  if (typeof window === "undefined") return () => {};
  window.addEventListener("dbx-plugin-env", handler);
  return () => window.removeEventListener("dbx-plugin-env", handler);
}

export async function whenReady() {
  if (!bridge) return null;
  try {
    return (await bridge.ready) ?? null;
  } catch {
    return null;
  }
}

/**
 * Invokes a sidecar method. Resolves to an object carrying `ok:false` on
 * failure, so callers branch on data rather than on exceptions.
 */
export async function invoke(method, params = {}) {
  if (!bridge?.invoke) {
    return { ok: false, error: "The DBX host bridge is not available" };
  }
  try {
    return { ok: true, value: await bridge.invoke(method, params) };
  } catch (error) {
    return { ok: false, error: error?.message ?? String(error) };
  }
}

/**
 * Reads plugin preferences from the sidecar's persisted file. Returns
 * `persisted:false` when the sidecar keeps no data directory, so the caller
 * can keep relying on its in-page cache instead of treating that as an error.
 */
export async function readSettingsStorage() {
  const reply = await invoke("dbx-calendar/settings/get");
  if (!reply.ok) return { ok: false, persisted: false };
  const settings = reply.value?.settings;
  return { ok: true, persisted: Boolean(reply.value?.persistence), settings: settings ?? {} };
}

/**
 * Writes plugin preferences through the sidecar. Accepts a partial object:
 * the sidecar merges it over the stored settings, so a change to one key
 * cannot reset keys an older or newer release added.
 */
export async function writeSettingsStorage(settings) {
  const reply = await invoke("dbx-calendar/settings/set", { settings });
  return reply.ok;
}
