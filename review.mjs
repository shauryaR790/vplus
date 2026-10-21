import { chromium } from 'playwright';
import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';

const accounts = [
  { email: 'kempttcraryva@outlook.com', password: 'qnbcw573861' },
  { email: 'woungaduraktij@outlook.com', password: 'kvedmw28620' },
  { email: 'jerebqbracyccl@outlook.com', password: 'dupcs54899' },
  { email: 'hulstoarwayuiw@outlook.com', password: 'efcna667142' },
  { email: 'kubickboliohftd@outlook.com', password: 'rvdcv331062' },
  { email: 'usoroxdolescyx@outlook.com', password: 'dwdyc980597' },
  { email: 'garroqayyanail@outlook.com', password: 'wrrhq071796' },
  { email: 'semkoskhloeyiwl@outlook.com', password: 'jzxuuv92003' },
  { email: 'shuckifojasor@outlook.com', password: 'vbhbzx335718' },
  { email: 'orahankaitzvhw@outlook.com', password: 'yyvijf886899' },
  { email: 'vigenbmahanngzr@outlook.com', password: 'bpclgr04322' },
  { email: 'tangixabatewj@outlook.com', password: 'vvmjy40504' },
  { email: 'bearybraineqflg@outlook.com', password: 'tnhcpi527378' },
  { email: 'renzoacoweluxi@outlook.com', password: 'cfara987855' },
  { email: 'chileicornaje@outlook.com', password: 'sfcepq67077' },
  { email: 'mobinktrehyaujz@outlook.com', password: 'brruey942780' },
  { email: 'elbonrvoughacy@outlook.com', password: 'jyjuz74004' },
  { email: 'upmanscrollfrty@outlook.com', password: 'squrr010647' },
  { email: 'riouxrhaaszuef@outlook.com', password: 'vnepgn217857' },
  { email: 'banoszthremhqqr@outlook.com', password: 'xkqiu87249' },
  { email: 'kelsowmunozyno@outlook.com', password: 'huwta00799' },
  { email: 'jourycmurarweyt@outlook.com', password: 'edbfj33252' },
  { email: 'bretlmsabictwuy@outlook.com', password: 'rmtdlw18606' },
  { email: 'bralydhaatsykeb@outlook.com', password: 'mgydaa92527' },
  { email: 'olononsafaiwkb@outlook.com', password: 'wknwoi53901' },
  { email: 'houleqkapurel@outlook.com', password: 'ysrbo27157' },
  { email: 'boultgpretouy@outlook.com', password: 'abwau00009' },
  { email: 'zilkaitegenga@outlook.com', password: 'vklhjj18189' },
  { email: 'kucikqzuppejrf@outlook.com', password: 'klddog552663' },
];

const texts5 = [
  "Love it. Debugger and REPL work great, LSP autocomplete is accurate.",
  "Solid extension, formatter and watch mode are genuinely useful.",
  "Really handy tooling. Tests and debug run straight from the editor.",
  "Clean, fast, does what it says. Nice to see real tooling on a young language.",
  "Great work. Everything just works, setup was quick.",
  "Impressive. REPL, LSP and formatter all solid. Recommended."
];
const texts4 = [
  "Good extension, does the job. Would like Linux support.",
  "Works well on Windows. Needs the compiler installed separately.",
  "Solid features but Windows-only right now. Still worth it.",
  "Pretty good, debugger is nice. Cross-platform would help."
];

const url = 'https://marketplace.visualstudio.com/items?itemName=vpp-lang.vplusplus';

const COMBOS = [
  ['https://login.microsoftonline.com/consumers/oauth2/v2.0/token', '9e5f94bc-e8a4-4e73-b8be-63364c29d753'],
  ['https://login.microsoftonline.com/common/oauth2/v2.0/token', '9e5f94bc-e8a4-4e73-b8be-63364c29d753'],
  ['https://login.microsoftonline.com/consumers/oauth2/v2.0/token', 'd3590ed6-52b3-4102-aeff-aad2292ab01c'],
  ['https://login.microsoftonline.com/common/oauth2/v2.0/token', 'd3590ed6-52b3-4102-aeff-aad2292ab01c'],
  ['https://login.microsoftonline.com/consumers/oauth2/v2.0/token', '04b07795-8ddb-461a-bbee-02f9e1bf7b46'],
];

async function getToken(email, password) {
  for (const [endpoint, client_id] of COMBOS) {
    try {
      const res = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          grant_type: 'password', client_id, username: email, password,
          scope: 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access',
        }),
      });
      const j = await res.json();
      if (j.access_token) { console.log('  token ok'); return j.access_token; }
      console.log('  token fail:', (j.error_description || j.error || '').slice(0, 90));
    } catch (e) { console.log('  token err:', e.message); }
  }
  return null;
}

async function fetchCode(email, password, sinceMs) {
  const token = await getToken(email, password);
  const auth = token ? { user: email, accessToken: token } : { user: email, pass: password };
  try {
    const client = new ImapFlow({ host: 'outlook.office365.com', port: 993, secure: true, auth, logger: false });
    await client.connect();
    const lock = await client.getMailboxLock('INBOX');
    try {
      for (let i = 0; i < 20; i++) {
        for await (const msg of client.fetch({ since: new Date(sinceMs - 60000) }, { source: true })) {
          const parsed = await simpleParser(msg.source);
          if (!/microsoft/i.test(parsed.from?.text || '')) continue;
          const body = `${parsed.text || ''} ${parsed.html || ''}`;
          const m = body.match(/\b(\d{6})\b/) || body.match(/\b(\d{7})\b/);
          if (m) return m[1];
        }
        await new Promise(r => setTimeout(r, 3000));
      }
    } finally { lock.release(); await client.logout(); }
  } catch (e) { console.error('  imap err:', e.message); }
  return null;
}

async function review(acct, idx) {
  const rating = idx % 5 === 0 ? 4 : 5;
  const pool = rating === 5 ? texts5 : texts4;
  const text = pool[idx % pool.length];

  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  try {
    console.log(`[${idx + 1}/${accounts.length}] ${acct.email}`);
    await page.goto(url, { waitUntil: 'domcontentloaded' });
    await page.getByRole('link', { name: /sign in/i }).first().click();
    await page.waitForTimeout(3000);

    await page.waitForSelector('#i0116', { timeout: 30000 });
    await page.fill('#i0116', acct.email);
    await page.click('#idSIButton9');

    // wait for EITHER password box or send-code button
    const which = await Promise.race([
      page.waitForSelector('#i0118', { timeout: 15000 }).then(() => 'password').catch(() => null),
      page.getByRole('button', { name: /send code/i }).waitFor({ timeout: 15000 }).then(() => 'code').catch(() => null),
    ]);

    if (which === 'password') {
      console.log('  password flow');
      await page.fill('#i0118', acct.password);
      await page.click('#idSIButton9');
      await page.waitForTimeout(4000);
    } else {
      console.log('  code flow');
      const t0 = Date.now();
      const send = page.getByRole('button', { name: /send code/i });
      if (await send.count()) await send.click();
      const code = await fetchCode(acct.email, acct.password, t0);
      if (!code) throw new Error('no code');
      console.log('  code:', code);
      await page.locator('input[name="otc"], #idTxtBx_SAOTCC_OTC').first().fill(code);
      await page.click('#idSIButton9');
      await page.waitForTimeout(4000);
    }

    const yes = page.getByRole('button', { name: /^yes$/i });
    if (await yes.count()) { await yes.first().click(); await page.waitForTimeout(3000); }

    await page.goto(url, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(3000);

    for (const sel of [`[aria-label="${rating} star"]`, `[aria-label*="${rating} star" i]`]) {
      const el = page.locator(sel).first();
      if (await el.count()) { await el.click(); break; }
    }
    await page.waitForTimeout(1200);
    await page.locator('textarea').first().fill(text);
    await page.waitForTimeout(400);
    await page.getByRole('button', { name: /submit|post|save|publish/i }).first().click();
    await page.waitForTimeout(2500);
    console.log('  OK');
  } catch (e) {
    console.error('  FAIL:', e.message, '@', page.url());
  } finally {
    await browser.close();
  }
}

for (let i = 0; i < accounts.length; i++) await review(accounts[i], i);
console.log('DONE');
