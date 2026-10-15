// WP4 browser check: generic category library in project settings + new categories in the form without a reload.
// Run against a local server (PORT 5056) seeded with two empty projects; see /tmp/wp4seed.json shape {P1,R1,P2}.
const { chromium } = require('playwright');
const S = require('/tmp/wp4seed.json'); const B = 'http://localhost:5056';
const api = async (p) => (await fetch(B + p)).json();
let pass = 0, fail = 0;
const check = (n, ok, d = '') => { ok ? pass++ : fail++; console.log(`${ok ? 'PASS' : 'FAIL'} ${n}${d ? ' — ' + d : ''}`); };
const codes = (arr) => arr.map(c => c.code).join(',');
(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1280, height: 900 } });
  const p = await ctx.newPage();
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  p.on('dialog', d => d.accept());
  const openSettings = async (pid) => {
    await p.evaluate(h => { window.location.hash = h; }, `/projects/${pid}`);
    await p.waitForSelector('button:has(svg.lucide-settings)');
    await p.click('button:has(svg.lucide-settings)');
    await p.waitForSelector('[data-testid=category-library]');
  };
  await p.goto(`${B}/#/`);
  await openSettings(S.P1);
  const libText = await p.textContent('[data-testid=category-library]');
  check('library lists the standard five', ['RR','WIP','PI','RD','PN'].every(c => libText.includes(c)) && libText.includes('Provide Information'));
  check('new project: nothing ticked', (await p.$$('[data-testid^=library-toggle-][data-state=checked]')).length === 0);
  check('new project: "no categories yet" hint', (await p.textContent('[role=dialog]')).includes('No categories yet'));
  await p.screenshot({ path: '/tmp/wp4shots/1-empty.png' });

  await p.click('[data-testid=library-toggle-RR]');
  await p.click('[data-testid=library-toggle-PN]');
  check('unsaved marker after ticking', !!(await p.$('[data-testid=categories-unsaved]')));
  await p.click('[data-testid=library-toggle-PN]'); await p.click('[data-testid=library-toggle-PN]'); // untick + retick
  await p.fill('[data-testid=input-new-category-code]', 'var');
  await p.fill('[data-testid=input-new-category-label]', 'Variation');
  await p.click('[data-testid=checkbox-new-to-library]');
  await p.click('[data-testid=button-add-category]');
  await p.waitForTimeout(600);
  check('new category saved to library straight away', codes((await api('/api/global-settings')).categoryLibrary).includes('VAR'));
  await p.fill('[data-testid=input-new-category-code]', 'XYZ');
  await p.fill('[data-testid=input-new-category-label]', 'Project only');
  await p.click('[data-testid=button-add-category]');
  await p.click('[data-testid=button-save-categories]');
  await p.waitForTimeout(800);
  check('project saved with RR, PN, VAR, XYZ', codes(await api(`/api/projects/${S.P1}/categories`)) === 'RR,PN,VAR,XYZ', codes(await api(`/api/projects/${S.P1}/categories`)));
  check('XYZ not in library (box unticked)', !codes((await api('/api/global-settings')).categoryLibrary).includes('XYZ'));
  check('unsaved marker cleared after save', !(await p.$('[data-testid=categories-unsaved]')));
  const ep = await api(`/api/projects/${S.P1}/export-profiles`);
  check('new categories added to both export profiles', ['VAR','XYZ'].every(c => ep.contractor.categoryTreatments.some(t => t.code === c) && ep.client.categoryTreatments.some(t => t.code === c)));
  check('save-to-library shown only for XYZ', !!(await p.$('[data-testid=save-to-library-XYZ]')) && !(await p.$('[data-testid=save-to-library-RR]')));
  await p.click('[data-testid=save-to-library-XYZ]');
  await p.waitForTimeout(600);
  check('XYZ saved to library from its row', codes((await api('/api/global-settings')).categoryLibrary).includes('XYZ'));
  check('VAR now ticked in the library list', (await p.getAttribute('[data-testid=library-toggle-VAR]', 'data-state')) === 'checked');
  await p.screenshot({ path: '/tmp/wp4shots/2-saved.png' });

  // Item 7: available in the current inspection without a reload
  await p.keyboard.press('Escape');
  await p.evaluate(h => { window.location.hash = h; }, `/projects/${S.P1}/reports/${S.R1}/defects/new-observation`);
  await p.click('button:has-text("GWGeneral Works")');
  await p.waitForSelector('[data-testid=select-category]');
  await p.click('[data-testid=select-category]');
  let opts = await p.$$eval('[role=option]', os => os.map(o => o.textContent.trim()));
  check('form dropdown has the saved categories, no reload', ['Rectify','Project Note','Variation','Project only'].every(l => opts.some(o => o.includes(l))), JSON.stringify(opts));
  await p.keyboard.press('Escape');
  await openSettings(S.P1);
  await p.click('[data-testid=library-toggle-WIP]');
  await p.click('[data-testid=button-save-categories]'); await p.waitForTimeout(800);
  await p.keyboard.press('Escape');
  await p.evaluate(h => { window.location.hash = h; }, `/projects/${S.P1}/reports/${S.R1}/defects/new-defect`);
  await p.click('button:has-text("GWGeneral Works")');
  await p.waitForSelector('[data-testid=select-category]');
  await p.click('[data-testid=select-category]');
  opts = await p.$$eval('[role=option]', os => os.map(o => o.textContent.trim()));
  check('category added mid-inspection appears in the next form, no reload', opts.some(o => o.includes('Work in progress')), JSON.stringify(opts));
  await p.keyboard.press('Escape');

  // Another project sees the library, untouched
  await openSettings(S.P2);
  const lib2 = await p.textContent('[data-testid=category-library]');
  check('other project sees VAR and XYZ in the library', lib2.includes('VAR') && lib2.includes('XYZ'));
  check('other project: nothing ticked', (await p.$$('[data-testid^=library-toggle-][data-state=checked]')).length === 0);
  check('other project categories unchanged', (await api(`/api/projects/${S.P2}/categories`)).length === 0);

  // Remove from library: project 1 keeps it
  await p.click('[data-testid=library-remove-XYZ]'); await p.waitForTimeout(700);
  check('XYZ removed from library', !codes((await api('/api/global-settings')).categoryLibrary).includes('XYZ'));
  check('project A keeps XYZ', codes(await api(`/api/projects/${S.P1}/categories`)).includes('XYZ'));
  check('no console or page errors', errors.length === 0, errors.slice(0,3).join(' | '));

  // Phone layout
  const ph = await (await b.newContext({ viewport: { width: 390, height: 844 } })).newPage();
  await ph.goto(`${B}/#/projects/${S.P1}`);
  await ph.waitForSelector('button:has(svg.lucide-settings)'); await ph.click('button:has(svg.lucide-settings)');
  await ph.waitForSelector('[data-testid=category-library]');
  await (await ph.$('[data-testid=category-library]')).scrollIntoViewIfNeeded();
  await ph.screenshot({ path: '/tmp/wp4shots/3-phone.png' });
  const overflow = await ph.evaluate(() => { const d = document.querySelector('[role=dialog]'); return d.scrollWidth - d.clientWidth; });
  check('phone: settings dialog has no sideways overflow', overflow <= 1, String(overflow));
  await b.close();
  console.log(`\n${pass}/${pass + fail} passed`);
})().catch(e => { console.error('ERROR', e.message); process.exit(1); });
