// WP1: AI action summaries say "to", not "must", and a manual action that says the works
// are done produces "No action necessary." instead of an invented task.
import {
  replaceMustWithTo, impliesNoActionRequired, generateActionSummary, NO_ACTION_SUMMARY,
} from "../server/action-summary";

let pass = 0, fail = 0;
const check = (name: string, ok: boolean, detail = "") => {
  if (ok) pass++; else { fail++; console.log(`FAIL ${name} ${detail}`); }
};

// --- "must" -> "to" ---
const mustCases: [string, string][] = [
  ["Contractor must provide a method statement.", "Contractor to provide a method statement."],
  ["Contractor MUST seal the joint.", "Contractor to seal the joint."],
  ["Builder must be advised before works proceed.", "Builder to be advised before works proceed."],
  ["Contractor must not proceed past the hold point.", "Contractor must not proceed past the hold point."],
  ["Works must not start; consultant must inspect first.", "Works must not start; consultant to inspect first."],
  ["Do not remove the props.", "Do not remove the props."],
  ["Mustard-coloured sealant to be replaced.", "Mustard-coloured sealant to be replaced."],
];
for (const [i, want] of mustCases) {
  const got = replaceMustWithTo(i);
  check(`must->to: ${i}`, got === want, `got ${JSON.stringify(got)}`);
}

// --- works-complete detection: real wording from the live app ---
const done = [
  "Works complete. No action necessary.", "Works are completed.", "Repair complete. No action necessary.",
  "Agreed works complete. No action necessary.", "What’s complete. No action necessary.",
  "Works are completed. No action required.", "No action necessary.", "No further action necessary.",
  "All works are complete, no further action necessary.", "Works are complete; no further works needed.",
  "No action needed - coating is applied.", "Nil. Repair done as per sow.", "Complete.",
  "no action required. observation only.",
];
for (const a of done) check(`no-action: ${a}`, impliesNoActionRequired(a) === true);

const notDone = [
  "Complete repair", "Complete repairs in line with method.", "9/6/26: complete repair.",
  "Complete standard detail: fleece, Aslan, mineral chip.", "AFC to confirm complete.",
  "Walsh Group to complete.", "Repair is complete. AFC to inspect completed work.",
  "Waterproofing completed. Install new reveal cladding to complete.",
  "Works complete but sealant still missing at the head.", "Works not complete.",
  "Remove partially complete repair material.", "Clean brickwork.", "", "   ",
];
for (const a of notDone) check(`needs-action: ${JSON.stringify(a)}`, impliesNoActionRequired(a) === false);

// --- end to end: no model call needed for a "done" action, and the hash is unchanged ---
const prevKey = process.env.OPENAI_API_KEY;
delete process.env.OPENAI_API_KEY;
const r = await generateActionSummary({
  observation: "Brickwork cracked.", actionRequired: "Works completed. No action necessary.",
  category: "Rectify", workType: "Brickwork",
});
check("generate: done action -> No action necessary.", r.summary === NO_ACTION_SUMMARY, r.summary);
check("generate: done action has a hash", /^[0-9a-f]{64}$/.test(r.hash));
check("generate: done action is not flagged as a fallback", r.source === "ai", r.source);

const r2 = await generateActionSummary({
  observation: "Brickwork cracked.", actionRequired: "Contractor must break out two bricks.",
  category: "Rectify", workType: "Brickwork",
});
check("generate: no key still falls back to the manual action untouched",
  r2.source === "fallback" && r2.summary === "Contractor must break out two bricks.", r2.summary);
if (prevKey) process.env.OPENAI_API_KEY = prevKey;

console.log(`${pass}/${pass + fail} passed`);
process.exit(fail ? 1 : 0);
