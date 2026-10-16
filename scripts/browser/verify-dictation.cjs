// Usage: start the app with OPENAI_BASE_URL/WORD_LIBRARY_API_BASE pointing at scripts/browser/dictation-mock.mjs (port 5999),
// seed one project/report/defect (ids in /tmp/wp7ids), then: B=<base> node scripts/browser/verify-dictation.cjs. Needs /tmp/wp7tone.wav (ffmpeg sine).
const { chromium } = require('playwright');
const B = process.env.B || 'http://127.0.0.1:5055'; const M = 'http://127.0.0.1:5999';
const [P, R, D] = require('fs').readFileSync('/tmp/wp7ids', 'utf8').trim().split(' ');
const res = []; const ok = (n, c, d = '') => { res.push([n, !!c]); console.log(`${c ? 'PASS' : 'FAIL'} ${n}${c ? '' : ' ' + d}`); };
(async () => {
  const br = await chromium.launch({ args: ['--use-fake-ui-for-media-stream', '--use-fake-device-for-media-stream', '--use-file-for-fake-audio-capture=/tmp/wp7tone.wav'] });
  const ctx = await br.newContext({ viewport: { width: 1280, height: 900 }, permissions: ['microphone'] });
  const pg = await ctx.newPage(); const errs = []; pg.on('pageerror', (e) => errs.push(e.message));
  const ids = process.env.ONLY === 'nokey';
  await pg.goto(`${B}/#/projects/${P}/reports/${R}/defects/${D}`); await pg.waitForSelector('[data-testid=button-dictate]', { timeout: 20000 });
  await pg.waitForTimeout(1500);
  const mic = (i) => pg.locator('[data-testid=button-dictate]').nth(i);
  const commentVal = () => pg.evaluate(() => [...document.querySelectorAll('textarea')].map((t) => t.value).find((v) => v.includes('Existing note')) || '');
  if (ids) {
    await mic(0).click(); await pg.waitForTimeout(500);
    ok('no OpenAI key: the microphone uses browser dictation', (await mic(0).getAttribute('data-dictation-mode')) === 'browser');
    ok('no OpenAI key: page still works', errs.length === 0, errs.join('|'));
    await br.close(); console.log(`\n${res.filter((r) => r[1]).length}/${res.length} passed`); return;
  }
  await fetch(`${M}/__mode/ok`);
  await mic(0).click(); await pg.waitForTimeout(400);
  ok('tap 1: recording with the word library', (await mic(0).getAttribute('data-dictation-mode')) === 'library' && (await mic(0).getAttribute('aria-label')) === 'Stop dictation');
  await pg.waitForTimeout(2500); await mic(0).click();
  let sawSpin = false; for (let i = 0; i < 40; i++) { if ((await mic(0).getAttribute('data-transcribing')) === 'true') sawSpin = true; if ((await commentVal()).includes('Sealant')) break; await pg.waitForTimeout(100); }
  ok('tap 2: shows transcribing', sawSpin);
  const v = await commentVal();
  ok('text added after the existing comment', v === 'Existing note. Sealant has failed at the façade joint.', v);
  const log = await (await fetch(`${M}/__log`)).json();
  ok('audio uploaded with keywords to the transcription service', log.log.some((l) => l.startsWith('POST /v1/audio/transcriptions')) && !v.includes('BAD'));
  ok('"add a word" button appears after dictation', await pg.locator('[data-testid=button-add-word]').isVisible());
  await pg.locator('[data-testid=button-add-word]').click();
  await pg.fill('[data-testid=input-word-term]', 'parapet capping'); await pg.fill('[data-testid=input-word-heard]', 'parrot capping');
  ok('dialog links to the full list on GitHub', await pg.locator('a[href="https://github.com/lanaangel-AFC/afc-word-library/edit/main/words.txt"]').isVisible());
  await pg.screenshot({ path: '/tmp/wp7_dialog.png' });
  await pg.click('[data-testid=button-save-word]'); await pg.waitForTimeout(800);
  const log2 = await (await fetch(`${M}/__log`)).json();
  ok('word saved to the library file as a mix-up', log2.file.endsWith('parrot capping => parapet capping\n'), log2.file.slice(-60));
  const toastTxt = await pg.evaluate(() => [...document.querySelectorAll('[role=status], li[data-state], ol li')].map((e) => e.textContent).join(' | '));
  ok('confirmation shown', toastTxt.includes('Added "parrot capping => parapet capping"'), toastTxt);
  ok('dialog closed', !(await pg.locator('[data-testid=input-word-term]').isVisible()));
  // after a few seconds autosave should have stored the text
  await pg.waitForTimeout(3000);
  const saved = await (await fetch(`${B}/api/projects/${P}/defects`)).json();
  ok('dictated text autosaved', saved.find((d) => d.id == D)?.comment === 'Existing note. Sealant has failed at the façade joint.', saved.find((d) => d.id == D)?.comment);
  // failure -> fallback
  await fetch(`${M}/__mode/fail`);
  await mic(1).click(); await pg.waitForTimeout(1500); await mic(1).click();
  await pg.waitForTimeout(1500);
  ok('failed transcription: message explains the fallback', await pg.getByText("couldn't be transcribed").first().isVisible().catch(() => false));
  await mic(1).click(); await pg.waitForTimeout(500);
  ok('next tap uses browser dictation', (await mic(1).getAttribute('data-dictation-mode')) === 'browser');
  await mic(1).click().catch(() => {});
  ok('no page errors', errs.length === 0, errs.join('|'));
  // phone
  await fetch(`${M}/__mode/ok`);
  const ph = await br.newContext({ viewport: { width: 390, height: 844 }, permissions: ['microphone'], isMobile: true, hasTouch: true });
  const p2 = await ph.newPage(); p2.on('pageerror', (e) => errs.push(e.message));
  await p2.goto(`${B}/#/projects/${P}/reports/${R}/defects/${D}`); await p2.waitForSelector('[data-testid=button-dictate]'); await p2.waitForTimeout(1500);
  const m2 = p2.locator('[data-testid=button-dictate]').first();
  await m2.tap(); await p2.waitForTimeout(2000); await m2.tap(); await p2.waitForSelector('[data-testid=button-add-word]', { timeout: 8000 }).catch(() => {});
  ok('phone: dictation works', await p2.locator('[data-testid=button-add-word]').isVisible());
  await m2.scrollIntoViewIfNeeded(); await p2.screenshot({ path: '/tmp/wp7_phone.png' });
  ok('phone: no sideways scrolling', await p2.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1));
  await br.close();
  console.log(`\n${res.filter((r) => r[1]).length}/${res.length} passed`);
})().catch((e) => { console.error(e); process.exit(1); });
