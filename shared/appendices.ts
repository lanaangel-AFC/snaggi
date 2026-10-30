// Report appendices (new report layout). Stored on projects.appendices as JSON
// [{name}], lettered A, B, C… in list order. NULL means the project has never set
// its list, so it gets the standard Appendix A below.

export const DEFAULT_APPENDIX_NAME = "Technical Guidance & Method Statements";

export interface Appendix { name: string }

export function parseAppendices(raw: unknown): Appendix[] {
  if (raw == null) return [{ name: DEFAULT_APPENDIX_NAME }];
  let v: any = raw;
  if (typeof raw === "string") {
    if (!raw.trim()) return [{ name: DEFAULT_APPENDIX_NAME }];
    try { v = JSON.parse(raw); } catch { return [{ name: DEFAULT_APPENDIX_NAME }]; }
  }
  if (!Array.isArray(v)) return [{ name: DEFAULT_APPENDIX_NAME }];
  return v.map((a: any) => ({ name: String(a?.name ?? "").replace(/\s+/g, " ").trim() }));
}

// A, B, … Z, AA, AB …
export function appendixLetter(i: number): string {
  let n = i + 1, s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
