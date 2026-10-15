// Usage: python3 scripts/browser/seed-project-instructions.py [BASE]; B=[BASE] node scripts/browser/verify-project-instructions.cjs
// Needs an image at /tmp/wp6photo.jpg for the photo step.
// Browser check (WP6): Project Instructions. Needs /tmp/wp6seed.json from the seed script.
const { chromium } = require('playwright');
const fs = require('fs');
const S = require('/tmp/wp6seed.json'); const B = process.env.B || 'http://localhost:5059';
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
  const pickParty = async (label) => { await p.click('[data-testid=select-responsibility]'); await p.click(`[role=option]:has-text("${label}")`); await p.waitForTimeout(200); };
  await p.goto(`${B}/#/`); await p.waitForTimeout(1200);

  console.log('--- inspection page');
  await go(`/projects/${S.P}/reports/${S.R}`, 'h1');
  check('Add Project Instruction button shown', await visible('[data-testid=button-add-instruction]'));
  await p.click('[data-testid=button-add-instruction]');
  await p.waitForSelector('[data-testid=worktype-chips]');
  await p.waitForTimeout(600);

  console.log('--- new instruction');
  check('ID shows "Assigned on save"', (await p.textContent('[data-testid=text-instruction-ref]')).includes('Assigned on save'));
  check('By When fixed as Ongoing', (await p.textContent('[data-testid=text-by-when]')).trim() === 'Ongoing');
  check('no verification on the form', !(await p.textContent('body')).includes('Verification'));
  check('date issued defaults to the inspection date', (await p.inputValue('[data-testid=input-date-issued]')) === '2026-10-16');
  await p.click('[data-testid=button-save-instruction]'); await p.waitForTimeout(800);
  check('empty form not saved', (await api(`/api/projects/${S.P}/instructions`)).length === 0);
  if (!(await p.$('[data-testid=chip-worktype-WP]'))) await p.click('[data-testid=button-more-worktypes]');
  await p.click('[data-testid=chip-worktype-SE]'); await p.click('[data-testid=chip-worktype-WP]');
  check('two work types selected', (await p.getAttribute('[data-testid=chip-worktype-SE]', 'aria-pressed')) === 'true' && (await p.getAttribute('[data-testid=chip-worktype-WP]', 'aria-pressed')) === 'true');
  await p.fill('[data-testid=input-observation]', 'Sealant applied without primer at several joints');
  await p.fill('[data-testid=input-action-required]', 'Prime all joints before applying sealant');
  await pickParty('Add a new party');
  check('new party box shown', await visible('[data-testid=input-new-party]'));
  await p.fill('[data-testid=input-new-party]', 'Sealant subcontractor');
  await p.click('[data-testid=button-save-instruction]');
  await p.waitForFunction(() => /instructions\/\d+$/.test(location.hash), null, { timeout: 15000 });
  await p.waitForTimeout(1200);
  const list1 = await api(`/api/projects/${S.P}/instructions`);
  const pi1 = list1[0];
  check('saved as PI-01', pi1?.ref === 'PI-01', pi1?.ref);
  check('both work types saved', JSON.stringify(pi1?.workTypes) === '["SE","WP"]', JSON.stringify(pi1?.workTypes));
  check('new party saved', pi1?.responsibility === 'Sealant subcontractor');
  check('form shows PI-01 after save', (await p.textContent('[data-testid=text-instruction-ref]')).includes('PI-01'));

  console.log('--- edit');
  await p.fill('[data-testid=input-action-required]', 'Prime and tool all joints before applying sealant');
  await p.click('[data-testid=button-save-instruction]'); await p.waitForTimeout(1500);
  check('edit saved', (await api(`/api/instructions/${pi1.id}`)).actionRequired === 'Prime and tool all joints before applying sealant');
  check('edit history shown', ((await p.textContent('[data-testid=instruction-history]').catch(() => '')) || '').includes('Action required'));
  check('ID unchanged after edit', (await p.textContent('[data-testid=text-instruction-ref]')).includes('PI-01'));
  const [fc] = await Promise.all([p.waitForEvent('filechooser'), p.click('[data-testid=button-add-instruction-photo]')]);
  await fc.setFiles('/tmp/wp6photo.jpg'); await p.waitForTimeout(2500);
  check('photo added', (await api(`/api/instructions/${pi1.id}`)).photos.length === 1);
  await p.screenshot({ path: '/tmp/wp6shots/1-form.png', fullPage: true });

  console.log('--- second instruction, then withdraw the first');
  await go(`/projects/${S.P}/reports/${S.R}/instructions/new`, '[data-testid=worktype-chips]');
  check('second form starts empty', (await p.getAttribute('[data-testid=chip-worktype-SE]', 'aria-pressed')) === 'false' && (await p.inputValue('[data-testid=input-action-required]')) === '' && (await p.textContent('[data-testid=text-instruction-ref]')).includes('Assigned on save'));
  await p.click('[data-testid=chip-worktype-SE]');
  await p.fill('[data-testid=input-action-required]', 'Protect completed sealant from paint overspray');
  await pickParty('Sealant subcontractor');
  check('the new party is offered again', (await p.textContent('[data-testid=select-responsibility]')).includes('Sealant subcontractor'));
  await pickParty('Contractor');
  await p.fill('[data-testid=input-date-issued]', '2026-10-20');
  await p.click('[data-testid=button-save-instruction]');
  await p.waitForFunction(() => /instructions\/\d+$/.test(location.hash), null, { timeout: 15000 });
  await p.waitForTimeout(800);
  check('second is PI-02', (await p.textContent('[data-testid=text-instruction-ref]')).includes('PI-02'));
  await go(`/projects/${S.P}/reports/${S.R}/instructions/${pi1.id}`, '[data-testid=button-withdraw]');
  await p.click('[data-testid=button-withdraw]'); await p.waitForTimeout(1500);
  check('withdrawn badge shown', await visible('[data-testid=badge-withdrawn]'));
  check('Reinstate offered', await visible('[data-testid=button-reinstate]'));

  console.log('--- inspection list');
  await go(`/projects/${S.P}/reports/${S.R}`, '[data-testid=instructions-list]');
  const cards = await p.$$eval('[data-testid^=card-instruction-]', es => es.map(e => e.textContent));
  check('listed newest first (PI-02, PI-01)', cards.length === 2 && cards[0].includes('PI-02') && cards[1].includes('PI-01'), cards.map(c => c.slice(0, 6)).join(','));
  check('PI-01 marked WITHDRAWN', cards[1].includes('WITHDRAWN'));
  check('PI-02 shows Ongoing and work type', cards[0].includes('Ongoing') && cards[0].includes('SE Sealant'), cards[0]);
  check('PI-02 saved with SE only', JSON.stringify((await api(`/api/projects/${S.P}/instructions`)).find(i => i.ref === 'PI-02')?.workTypes) === '["SE"]');
  await p.screenshot({ path: '/tmp/wp6shots/2-inspection.png' });

  console.log('--- observation form hides the PI category');
  await go(`/projects/${S.P}/reports/${S.R}/defects/new-observation`, 'button:has-text("SESealant")');
  await p.click('button:has-text("SESealant")'); await p.waitForSelector('#comment');
  await p.click('[data-testid=select-category]'); await p.waitForTimeout(400);
  const opts = await p.$$eval('[role=option]', es => es.map(e => e.textContent));
  check('Project Instruction not offered as a category', opts.length > 0 && !opts.some(o => /Project Instruction/i.test(o)), JSON.stringify(opts));
  await p.keyboard.press('Escape');

  console.log('--- exports');
  await go(`/projects/${S.P}/reports/${S.R}`, 'button:has-text("Export Report")');
  const grab = async (label, sub, tag) => {
    await p.click('button:has-text("Export Report")');
    if (sub) { await p.hover('text=Export split by audience'); await p.waitForTimeout(400); }
    const dlP = p.waitForEvent('download', { timeout: 90000 });
    if (sub) await p.locator('[role=menu]').last().locator(`text=${label}`).nth(sub === 'client' ? 1 : 0).click(); else await p.click(`text=${label}`);
    await p.waitForTimeout(1200);
    if (await p.$('button:has-text("Export anyway")')) await p.click('button:has-text("Export anyway")');
    const dl = await dlP; const f = `/tmp/wp6dl/${tag}-${await dl.suggestedFilename()}`; await dl.saveAs(f); return f;
  };
  const files = [];
  files.push(await grab('Report — Word Document (.docx)', null, 'new'));
  files.push(await grab('Report — PDF (.pdf)', null, 'new'));
  files.push(await grab('Word Document (.docx)', 'client', 'new-client'));
  files.push(await grab('PDF (.pdf)', 'client', 'new-client'));
  await go(`/projects/${S.OP}/reports/${S.OR}`, 'button:has-text("Export Report")');
  check('old project: no Add Project Instruction button', !(await p.$('[data-testid=button-add-instruction]')));
  files.push(await grab('Report — Word Document (.docx)', null, 'old'));
  files.push(await grab('Report — PDF (.pdf)', null, 'old'));
  check('six exports downloaded', files.length === 6 && files.every(f => fs.statSync(f).size > 1000));
  check('no console or page errors', errors.length === 0, errors.slice(0, 3).join(' | '));

  console.log('--- phone');
  const pp = await (await b.newContext({ viewport: { width: 390, height: 844 } })).newPage();
  await pp.goto(`${B}/#/projects/${S.P}/reports/${S.R}/instructions/${pi1.id}`); await pp.waitForSelector('[data-testid=worktype-chips]'); await pp.waitForTimeout(1000);
  check('phone form: no sideways scroll', await pp.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1));
  await pp.screenshot({ path: '/tmp/wp6shots/3-phone-form.png', fullPage: true });
  await pp.goto(`${B}/#/projects/${S.P}/reports/${S.R}`); await pp.waitForSelector('[data-testid=button-add-instruction]'); await pp.waitForTimeout(1000);
  const r = await pp.evaluate(() => Math.round(document.querySelector('[data-testid=button-add-instruction]').getBoundingClientRect().right));
  check('phone inspection page: buttons fit', r <= 390 && await pp.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1), String(r));
  await pp.screenshot({ path: '/tmp/wp6shots/4-phone-inspection.png' });
  await b.close();
  console.log(`\n${pass}/${pass + fail} passed`);
})().catch(e => { console.error('ERROR', e.message); process.exit(1); });
