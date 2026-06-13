// Single CSV field escaper, shared by the server export endpoint and any client-side
// CSV builders so escaping behaviour can never diverge. Quotes a field only when it
// contains a comma, double-quote, or newline; doubles embedded quotes per RFC 4180.
export function escapeCsvField(value: unknown): string {
  const s = value == null ? "" : String(value);
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}
