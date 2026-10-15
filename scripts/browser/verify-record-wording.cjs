// WP2 browser check: record wording in new projects (featureVersion 2) vs original wording in older ones.
// Run against a local server (PORT 5057) seeded per /tmp/wp2seed.json {N:{P,R,D,E}, O:{P,R,D}}.
const { chromium } = require('playwright');
const S = require('/tmp/wp2seed.json'); const B = 'http://localhost:5057';
let pass = 0, fail = 0;
const check = (n, ok, d = '') => { ok ? pass++ : fail++; console.log(`${ok ? 'PASS' : 'FAIL'} ${n}${d ? ' — ' + d : ''}`); };
(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1280, height: 900 } });
  const p = await ctx.newPage();
  const errors = []; p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
  let lastDialog = ''; p.on('dialog', d => { lastDialog = d.message(); d.dismiss(); });
  const go = async (h, sel) => { await p.evaluate(x => { window.location.hash = x; }, h); if (sel) await p.waitForSelector(sel, { timeout: 20000 }); await p.waitForTimeout(700); };
  const body = () => p.textContent('body');
  const saveText = () => p.textContent('[data-testid=button-save-defect]');
  const ph = (sel) => p.getAttribute(sel, 'placeholder');

  await p.goto(`${B}/#/`); await p.waitForTimeout(1500);
  check('project list: new tagline', (await body()).includes('Inspection records and elevation markup'));

  console.log('--- new project');
  const N = S.N;
  await go(`/projects/${N.P}/reports/${N.R}/defects/${N.D}`, '#comment');
  check('observation: "Observation ID"', (await body()).includes('Observation ID'));
  check('observation: "Update Observation"', (await saveText()).includes('Update Observation'), await saveText());
  check('observation: observation placeholder', (await ph('#comment')) === 'Describe what was observed...', await ph('#comment'));
  check('observation: action placeholder', (await ph('#actionRequired')) === 'What action is needed, if any?', await ph('#actionRequired'));
  await p.click('[data-testid=button-record-type-defect]'); await p.waitForTimeout(900);
  check('switched to defect: "Defect ID"', (await body()).includes('Defect ID'));
  check('switched to defect: "Update Defect"', (await saveText()).includes('Update Defect'), await saveText());
  check('switched to defect: defect placeholder', (await ph('#comment')) === 'Describe the defect observed...');
  await p.click('[data-testid=button-record-type-observation]'); await p.waitForTimeout(900);
  check('switched back: "Update Observation"', (await saveText()).includes('Update Observation'));

  await go(`/projects/${N.P}/reports/${N.R}`, 'h1');
  await go(`/projects/${N.P}/reports/${N.R}/defects/new-observation`, 'button:has-text("GWGeneral Works")');
  await p.click('button:has-text("GWGeneral Works")'); await p.waitForSelector('#comment');
  check('new observation: "Save Observation"', (await saveText()).includes('Save Observation'), await saveText());
  await go(`/projects/${N.P}/reports/${N.R}`, 'h1');
  await go(`/projects/${N.P}/reports/${N.R}/defects/new-defect`, 'button:has-text("GWGeneral Works")');
  await p.click('button:has-text("GWGeneral Works")'); await p.waitForSelector('#comment');
  check('new defect: "Save Defect"', (await saveText()).includes('Save Defect'), await saveText());

  await go(`/projects/${N.P}/reports/${N.R}`, `a[href*="/defects/${N.D}"]`);
  const card = p.locator(`a[href*="/defects/${N.D}"]`).locator('xpath=..');
  await card.hover(); await card.locator('button').last().click({ force: true }); await p.waitForTimeout(400);
  check('inspection page: delete asks about "this observation"', lastDialog === 'Delete this observation and all its photos?', lastDialog);
  check('record still there after dismissing', (await (await fetch(`${B}/api/defects/${N.D}`)).json()).id === N.D);

  await go(`/projects/${N.P}/reports/${N.R}/elevations/${N.E}`, 'img[alt="Elevation drawing"]');
  const img = await p.$('img[alt="Elevation drawing"]'); const bx = await img.boundingBox();
  await p.click('button:has(svg.lucide-crosshair)');
  await p.mouse.click(bx.x + bx.width * 0.5, bx.y + bx.height * 0.5);
  await p.waitForSelector('button:has-text("Link an existing record")', { timeout: 10000 });
  await p.click('button:has-text("Link an existing record")');
  await p.waitForSelector('[role=dialog]:has-text("UID"), [role=dialog]:has-text("Link to")', { timeout: 10000 }).catch(() => {});
  await p.waitForTimeout(800);
  const dlg = (await p.$$eval('[role=dialog]', ds => ds.map(d => d.textContent + ' ' + [...d.querySelectorAll('[placeholder], span')].map(e => e.getAttribute('placeholder') || e.textContent).join(' ')).join(' | ')));
  check('drawing: "Link to Record"', dlg.includes('Link to Record'), dlg.slice(0, 120));
  check('drawing: "Select a record" placeholder', dlg.includes('Select a record or enter custom'));
  await p.keyboard.press('Escape'); await p.waitForTimeout(500);
  const markers = await (await fetch(`${B}/api/elevations/${N.E}/markers`)).json();
  for (const m of markers) await fetch(`${B}/api/markers/${m.id}`, { method: 'DELETE' });

  await go(`/projects/${N.P}`, 'button:has(svg.lucide-settings)');
  await p.click('button:has(svg.lucide-settings)'); await p.waitForSelector('[role=dialog]');
  check('settings: "record ID"', (await p.textContent('[role=dialog]')).includes('Choose which fields appear in the record ID'));
  await p.keyboard.press('Escape');

  console.log('--- old project');
  const O = S.O;
  await go(`/projects/${O.P}/reports/${O.R}/defects/${O.D}`, '#comment');
  check('old observation: still "Defect ID"', (await body()).includes('Defect ID') && !(await body()).includes('Observation ID'));
  check('old observation: still "Update Defect"', (await saveText()).includes('Update Defect'), await saveText());
  check('old observation: still defect placeholder', (await ph('#comment')) === 'Describe the defect observed...');
  await go(`/projects/${O.P}/reports/${O.R}`, `a[href*="/defects/${O.D}"]`);
  const card2 = p.locator(`a[href*="/defects/${O.D}"]`).locator('xpath=..');
  await card2.hover(); await card2.locator('button').last().click({ force: true }); await p.waitForTimeout(400);
  check('old inspection page: delete still says "this defect"', lastDialog === 'Delete this defect and all its photos?', lastDialog);
  await go(`/projects/${O.P}`, 'button:has(svg.lucide-settings)');
  await p.click('button:has(svg.lucide-settings)'); await p.waitForSelector('[role=dialog]');
  check('old settings: still "defect ID"', (await p.textContent('[role=dialog]')).includes('Choose which fields appear in the defect ID'));
  check('no console or page errors', errors.length === 0, errors.slice(0, 3).join(' | '));

  console.log('--- phone');
  const pp = await (await b.newContext({ viewport: { width: 390, height: 844 } })).newPage();
  for (const [k, X] of [['new', N], ['old', O]]) {
    await pp.goto(`${B}/#/projects/${X.P}/reports/${X.R}/defects/${X.D}`); await pp.waitForSelector('[data-testid=record-type-toggle]'); await pp.waitForTimeout(800);
    const r = await pp.evaluate(() => [...document.querySelectorAll('[data-testid^=button-record-type-]')].map(e => { const b = e.getBoundingClientRect(); return { t: e.textContent, right: Math.round(b.right), clipped: e.scrollWidth > e.clientWidth + 1 }; }));
    check(`phone (${k}): toggle fully visible`, r.every(x => x.right <= 390 && !x.clipped), JSON.stringify(r));
    check(`phone (${k}): page has no sideways scroll`, await pp.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1));
    await pp.screenshot({ path: `/tmp/wp2shots/phone-${k}.png` });
  }
  await b.close();
  console.log(`\n${pass}/${pass + fail} passed`);
})().catch(e => { console.error('ERROR', e.message); process.exit(1); });
