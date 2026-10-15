// Query parameters for hash-routed pages.
//
// wouter's hash navigate() writes "?a=b" from navigate("/x?a=b") into location.search, not
// into the hash, and never clears it on later navigations. So a parameter has to be looked
// for in both places, and taken once: read it when the page mounts, then remove it so it
// cannot leak into the next page (e.g. a stale ?returnTo= turning a later visit to the
// drawing into a pin errand).

export function readQueryParams(): URLSearchParams {
  const out = new URLSearchParams(window.location.search);
  const hash = window.location.hash;
  const q = hash.indexOf("?");
  if (q !== -1) new URLSearchParams(hash.substring(q)).forEach((v, k) => out.set(k, v));
  return out;
}

export function clearQueryParams(keys: string[]): void {
  const url = new URL(window.location.href);
  let changed = false;
  for (const k of keys) {
    if (url.searchParams.has(k)) { url.searchParams.delete(k); changed = true; }
  }
  const hash = url.hash;
  const q = hash.indexOf("?");
  if (q !== -1) {
    const hp = new URLSearchParams(hash.substring(q));
    let hChanged = false;
    for (const k of keys) if (hp.has(k)) { hp.delete(k); hChanged = true; }
    if (hChanged) {
      const rest = hp.toString();
      url.hash = hash.substring(0, q) + (rest ? `?${rest}` : "");
      changed = true;
    }
  }
  if (changed) window.history.replaceState(window.history.state, "", url.href);
}
