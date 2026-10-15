// Browser check (WP5): no-action notes. Needs a running app seeded with scripts/browser/seed-no-action-notes.py
// (writes /tmp/wp5seed.json). Usage: python3 scripts/browser/seed-no-action-notes.py [BASE]; B=[BASE] node scripts/browser/verify-no-action-notes.cjs
const { chromium } = require('playwright');
const S = require('/tmp/wp5seed.json'); const B = process.env.B || 'http://localhost:5058';
const fs = require('fs');
let pass = 0, fail = 0;
const check = (n, ok, d = '') => { ok ? pass++ : fail++; console.log(`${ok ? 'PASS' : 'FAIL'} ${n}${d ? ' — ' + d : ''}`); };
const api = async (p) => (await fetch(B + p)).json();
(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1280, height: 900 }, acceptDownloads: true });
  const p = await ctx.newPage();
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  p.on('dialog', d => d.accept());
  const go = async (h, sel) => { await p.evaluate(x => { window.location.hash = x; }, h); if (sel) await p.waitForSelector(sel, { timeout: 20000 }); await p.waitForTimeout(800); };
  const visible = async (sel) => !!(await p.$(sel)) && await p.isVisible(sel);
  await p.goto(`${B}/#/`); await p.waitForTimeout(1200);

  console.log('--- new observation as a note');
  await go(`/projects/${S.P}/reports/${S.R}`, 'h1');
  await go(`/projects/${S.P}/reports/${S.R}/defects/new-observation`, 'button:has-text("GWGeneral Works")');
  await p.click('button:has-text("GWGeneral Works")'); await p.waitForSelector('#comment');
  check('toggle shown on a new observation', await visible('[data-testid=action-needed-toggle]'));
  check('defaults to Yes', (await p.getAttribute('[data-testid=button-action-needed-yes]', 'class')).includes('bg-primary'));
  check('By Whom shown while Yes', await visible('[data-testid=input-assigned-to]'));
  await p.click('[data-testid=button-action-needed-no]'); await p.waitForTimeout(300);
  check('No: By Whom hidden', !(await visible('[data-testid=input-assigned-to]')));
  check('No: By When hidden', !(await visible('[data-testid=input-due-date]')));
  check('No: verification hidden', !(await visible('[data-testid=select-verification-method]')));
  check('No: category hidden', !(await visible('[data-testid=select-category]')));
  check('No: explainer shown', await visible('[data-testid=note-explainer]'));
  check('No: reference "Assigned on save"', (await p.textContent('[data-testid=text-defect-uid]')).includes('Assigned on save'));
  check('No: action text not required', (await p.getAttribute('#actionRequired', 'required')) === null);
  check('photos section still shown', (await p.textContent('body')).toLowerCase().includes('photo'));
  await p.fill('#comment', 'Downpipe bracket fixed, no issue');
  await p.click('[data-testid=button-save-defect]');
  await p.waitForTimeout(2500);
  const list = await api(`/api/reports/${S.R}/defects`);
  const created = list.filter(d => d.comment === 'Downpipe bracket fixed, no issue').sort((a, b) => b.id - a.id)[0];
  check('saved with a Note-12-nn reference', /^Note-12-\d\d$/.test(created?.uid || ''), created?.uid);

  console.log('--- open the note');
  await go(`/projects/${S.P}/reports/${S.R}/defects/${created.id}`, '#comment');
  check('heading "Note reference" + its reference', (await p.textContent('body')).includes('Note reference') && (await p.textContent('[data-testid=text-defect-uid]')).includes(created.uid));
  check('no Mark Complete on a note', !(await visible('[data-testid=button-mark-complete]')));
  check('no Action Summary box on a note', !(await p.$('[data-testid=input-action-summary]')));
  check('no Open badge on a note', !(await p.textContent('h1').then(() => p.$$eval('.flex.items-center.gap-3 > div, h1 ~ *', es => es.some(e => e.textContent.trim() === 'Open')))));
  check('toggle shows No', (await p.getAttribute('[data-testid=button-action-needed-no]', 'class')).includes('bg-primary'));
  await p.screenshot({ path: '/tmp/wp5shots/1-note-form.png', fullPage: false });

  console.log('--- switch to Yes');
  await p.click('[data-testid=button-action-needed-yes]'); await p.waitForTimeout(300);
  check('Yes: By Whom back', await visible('[data-testid=input-assigned-to]'));
  await p.fill('#actionRequired', 'Contractor to monitor bracket');
  await p.click('[data-testid=button-save-defect]'); await p.waitForTimeout(2500);
  const after = await api(`/api/defects/${created.id}`);
  check('Yes saved: full UID assigned', !after.uid.startsWith('Note-') && /GW-\d\d$/.test(after.uid) && after.actionNeeded === true, after.uid);
  await go(`/projects/${S.P}/reports/${S.R}/defects/${created.id}`, '#comment');
  check('form shows the new UID', (await p.textContent('[data-testid=text-defect-uid]')).includes(after.uid));
  check('Yes: Action Summary box back', !!(await p.$('[data-testid=input-action-summary]')));
  check('Mark Complete back', await visible('[data-testid=button-mark-complete]'));

  console.log('--- inspection page');
  await go(`/projects/${S.P}/reports/${S.R}`, `a[href*="/defects/${S.N}"]`);
  const noteCard = p.locator(`a[href*="/defects/${S.N}"]`);
  check('note card shows NOTE · NO ACTION', (await noteCard.textContent()).includes('NOTE · NO ACTION'));
  check('note card has no stale-summary badge', !(await noteCard.textContent()).includes('Stale summary'));
  check('note card shows its reference', (await noteCard.textContent()).includes(S.NUID));

  console.log('--- drawing');
  await go(`/projects/${S.P}/reports/${S.R}/elevations/${S.E}`, 'img[alt="Elevation drawing"]');
  await p.waitForTimeout(1200);
  check('legend shows Notes 1', ((await p.textContent('[data-testid=legend-notes]').catch(() => '')) || '').includes('Notes 1'));
  const fills = await p.$$eval('svg path, svg circle', es => es.map(e => e.getAttribute('fill')).filter(Boolean));
  check('a grey note pin is drawn', fills.some(f => f.toUpperCase() === '#94A3B8'), JSON.stringify([...new Set(fills)].slice(0, 8)));
  await p.screenshot({ path: '/tmp/wp5shots/2-drawing.png' });

  console.log('--- exports');
  await go(`/projects/${S.P}/reports/${S.R}`, 'button:has-text("Export Report")');
  const grab = async (label, sub) => {
    await p.click('button:has-text("Export Report")');
    if (sub) { await p.hover('text=Export split by audience'); await p.waitForTimeout(400); }
    const dlP = p.waitForEvent('download', { timeout: 90000 });
    if (sub) await p.locator('[role=menu]').last().locator(`text=${label}`).nth(sub === 'client' ? 1 : 0).click(); else await p.click(`text=${label}`);
    await p.waitForTimeout(1200);
    if (await p.$('button:has-text("Export anyway")')) await p.click('button:has-text("Export anyway")');
    const dl = await dlP;
    const f = `/tmp/wp5dl/${(sub || 'report')}-${await dl.suggestedFilename()}`; await dl.saveAs(f); return f;
  };
  const files = [];
  files.push(await grab('Report — Word Document (.docx)'));
  files.push(await grab('Report — PDF (.pdf)'));
  files.push(await grab('Word Document (.docx)', 'client'));
  files.push(await grab('PDF (.pdf)', 'client'));
  fs.writeFileSync('/tmp/wp5dl/files.json', JSON.stringify(files));
  check('four exports downloaded', files.length === 4 && files.every(f => fs.statSync(f).size > 1000), files.join(' '));

  console.log('--- old project');
  await go(`/projects/${S.OP}/reports/${S.OR}`, 'h1');
  await go(`/projects/${S.OP}/reports/${S.OR}/defects/new-observation`, 'button:has-text("GWGeneral Works")');
  await p.click('button:has-text("GWGeneral Works")'); await p.waitForSelector('#comment');
  check('old project: no Action required toggle', !(await p.$('[data-testid=action-needed-toggle]')));
  check('old project: By Whom still shown', await visible('[data-testid=input-assigned-to]'));
  check('no console or page errors', errors.length === 0, errors.slice(0, 3).join(' | '));

  console.log('--- phone');
  const pp = await (await b.newContext({ viewport: { width: 390, height: 844 } })).newPage();
  await pp.goto(`${B}/#/projects/${S.P}/reports/${S.R}/defects/${S.N}`); await pp.waitForSelector('[data-testid=action-needed-toggle]'); await pp.waitForTimeout(800);
  await (await pp.$('[data-testid=action-needed-card]')).scrollIntoViewIfNeeded();
  const t = await pp.evaluate(() => [...document.querySelectorAll('[data-testid^=button-action-needed-]')].map(e => Math.round(e.getBoundingClientRect().right)));
  check('phone: toggle fully visible', t.length === 2 && t.every(r => r <= 390), JSON.stringify(t));
  check('phone: no sideways scroll', await pp.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1));
  await pp.screenshot({ path: '/tmp/wp5shots/3-phone.png' });
  await b.close();
  console.log(`\n${pass}/${pass + fail} passed`);
})().catch(e => { console.error('ERROR', e.message); process.exit(1); });
