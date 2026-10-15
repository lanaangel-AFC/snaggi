// Browser check for the pin round trip (WP3). Runs against a local server seeded with one
// record carried from inspection 01 into 02 and one image elevation; the ids are read from
// /tmp/wp3seed.json ({P, R1, D1, R2, D2, UID, E}). Start the server with
// DATA_DIR=<scratch> PORT=5055 NODE_ENV=production node dist/index.cjs
const { chromium } = require('playwright');
const S = require('/tmp/wp3seed.json');
const B = 'http://localhost:5055';
const api = async (p, o) => (await fetch(B + p, o)).json();
let pass = 0, fail = 0;
const check = (n, ok, d = '') => { ok ? pass++ : fail++; console.log(`${ok ? 'PASS' : 'FAIL'} ${n}${d ? ' — ' + d : ''}`); };
(async () => {
  const browser = await chromium.launch();
  for (const vp of [{ name: 'desktop', width: 1280, height: 860 }, { name: 'phone', width: 390, height: 844 }]) {
    console.log(`\n=== ${vp.name}`);
    const ctx = await browser.newContext({ viewport: { width: vp.width, height: vp.height } });
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', e => errors.push(e.message));
    page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
    const markersBefore = (await api(`/api/elevations/${S.E}/markers`)).length;

    // A: unsaved edit, Mark on Elevation, place, move, accept
    const formUrl = `${B}/#/projects/${S.P}/reports/${S.R2}/defects/${S.D2}`;
    await page.goto(formUrl); await page.waitForSelector('#comment');
    const newText = `Brickwork cracked. Edited ${vp.name}.`;
    await page.fill('#comment', newText);
    await page.click('[data-testid=button-mark-on-elevation]');
    await page.waitForSelector('[data-testid=pin-errand-bar]', { timeout: 10000 });
    const u = new URL(page.url().replace('#/', ''));
    const hp = new URLSearchParams(page.url().split('?')[1]);
    check('drawing opens in the form\'s inspection', page.url().includes(`/reports/${S.R2}/elevations/${S.E}`), page.url());
    check('address cleared of hand-over parameters on arrival', !page.url().includes('returnTo'), page.url());
    const saved = await api(`/api/defects/${S.D2}`);
    check('unsaved edit was saved before leaving the form', saved.comment === newText, saved.comment);
    check('Accept disabled before a pin is placed', await page.isDisabled('[data-testid=button-accept-pin]'));
    await page.screenshot({ path: `/tmp/wp3shots/${vp.name}-1-errand.png` });
    const img = await page.waitForSelector('img[alt="Elevation drawing"]');
    const box = await img.boundingBox();
    await page.mouse.click(box.x + box.width * 0.25, box.y + box.height * 0.3);
    await page.waitForSelector('[data-testid=draft-pin]');
    await page.mouse.click(box.x + box.width * 0.6, box.y + box.height * 0.7);
    const dialogs = await page.$$('[role=dialog]');
    check('no link dialog appears', dialogs.length === 0, `${dialogs.length} dialogs`);
    check('only one draft pin after a second tap', (await page.$$('[data-testid=draft-pin]')).length === 1);
    await page.screenshot({ path: `/tmp/wp3shots/${vp.name}-2-draft.png` });
    await page.click('[data-testid=button-accept-pin]');
    await page.waitForURL(u => u.toString().includes(`/defects/${S.D2}`) && !u.toString().includes('elevations'), { timeout: 10000 }).catch(() => {});
    check('back on the same form after Accept', page.url().endsWith(`/reports/${S.R2}/defects/${S.D2}`), page.url());
    await page.waitForSelector('#comment');
    check('form shows the saved text', (await page.inputValue('#comment')) === newText);
    const mk = await api(`/api/elevations/${S.E}/markers`);
    const m = mk[mk.length - 1];
    check('exactly one pin added', mk.length === markersBefore + 1, `${markersBefore} -> ${mk.length}`);
    check('pin linked to the row id', m.defectId === S.D2, String(m.defectId));
    check('pin carries the UID', m.defectUid === S.UID, m.defectUid);
    check('pin is where the SECOND tap was', Math.abs(m.xPercent - 60) < 1.5 && Math.abs(m.yPercent - 70) < 1.5, `${m.xPercent.toFixed(1)},${m.yPercent.toFixed(1)}`);
    await page.screenshot({ path: `/tmp/wp3shots/${vp.name}-3-back.png` });

    // B: Cancel returns without a pin
    await page.click('[data-testid=button-mark-on-elevation]');
    await page.waitForSelector('[data-testid=pin-errand-bar]');
    const box2 = await (await page.waitForSelector('img[alt="Elevation drawing"]')).boundingBox();
    await page.mouse.click(box2.x + box2.width * 0.4, box2.y + box2.height * 0.4);
    await page.click('[data-testid=button-cancel-pin]');
    await page.waitForSelector('#comment');
    check('Cancel returns to the form', page.url().endsWith(`/reports/${S.R2}/defects/${S.D2}`), page.url());
    check('Cancel adds no pin', (await api(`/api/elevations/${S.E}/markers`)).length === markersBefore + 1);

    // C: no edits -> no save (to-do must not tick)
    const histBefore = (await api(`/api/defects/${S.D2}/observation-history`)).length;
    await page.click('[data-testid=button-mark-on-elevation]');
    await page.waitForSelector('[data-testid=pin-errand-bar]');
    check('unchanged form is not re-saved', (await api(`/api/defects/${S.D2}/observation-history`)).length === histBefore);
    await page.click('[data-testid=button-cancel-pin]');
    await page.waitForSelector('#comment');

    check('no console or page errors', errors.length === 0, errors.slice(0, 3).join(' | '));
    await ctx.close();
  }

  // D: form opened on the inspection-01 row from inspection 02 -> pin goes to the 02 row
  console.log('\n=== older row opened from the current inspection');
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 860 } });
  const page = await ctx.newPage();
  await page.goto(`${B}/#/projects/${S.P}/reports/${S.R2}/defects/${S.D1}`); await page.waitForSelector('#comment');
  await page.fill('#comment', 'Amended from inspection 02 via the older row.');
  await page.click('[data-testid=button-mark-on-elevation]');
  await page.waitForSelector('[data-testid=pin-errand-bar]');
  const r1 = await api(`/api/defects/${S.D1}`);
  check('issued inspection-01 row left untouched', r1.comment === 'Brickwork cracked.', r1.comment);
  const bD = await (await page.waitForSelector('img[alt="Elevation drawing"]')).boundingBox();
  await page.mouse.click(bD.x + bD.width * 0.3, bD.y + bD.height * 0.5);
  await page.click('[data-testid=button-accept-pin]');
  await page.waitForSelector('#comment');
  const mD = (await api(`/api/elevations/${S.E}/markers`)).pop();
  check('pin links to the current-inspection row, not the issued one', mD.defectId === S.D2, String(mD.defectId));
  check('returns to the current-inspection row', page.url().endsWith(`/defects/${S.D2}`), page.url());

  // E: after an errand, opening the drawing normally is NOT an errand (no leaked params)
  console.log('\n=== no leftover parameters');
  {
    const c = await browser.newContext({ viewport: { width: 1280, height: 860 } });
    const p = await c.newPage();
    await p.goto(`${B}/#/projects/${S.P}/reports/${S.R2}/defects/${S.D2}`); await p.waitForSelector('#comment');
    await p.click('[data-testid=button-mark-on-elevation]'); await p.waitForSelector('[data-testid=pin-errand-bar]');
    await p.click('[data-testid=button-cancel-pin]'); await p.waitForSelector('#comment');
    check('address carries no leftover parameters', !p.url().includes('returnTo') && !p.url().includes('defectId'), p.url());
    await p.evaluate((u) => { window.location.hash = u; }, `/projects/${S.P}/reports/${S.R2}/elevations/${S.E}`);
    await p.waitForSelector('img[alt="Elevation drawing"]'); await p.waitForTimeout(800);
    check('plain visit to the drawing shows no pin-errand bar', (await p.$$('[data-testid=pin-errand-bar]')).length === 0);
    await p.click('text=Add');
    const bx = await (await p.waitForSelector('img[alt="Elevation drawing"]')).boundingBox();
    await p.mouse.click(bx.x + bx.width * 0.8, bx.y + bx.height * 0.2);
    await p.waitForTimeout(500);
    const btns = await p.$$eval('[role=dialog] button', bs => bs.map(b => b.textContent.trim()));
    console.log('   chooser buttons:', JSON.stringify(btns));
    await c.close();
  }
  await browser.close();
  console.log(`\n${pass}/${pass + fail} passed`);
})().catch(e => { console.error('ERROR', e.message); process.exit(1); });
