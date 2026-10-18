import { chromium } from 'playwright';

const accounts = [
  { "email": "sureyeawh@gmail.com", "password": "shaurya2009" },
  { "email": "rshaurya790@gmail.com", "password": "shaurya2009" },
  { "email": "shauryaljku@gmail.com", "password": "shaurya2009" }
];

const run = parseInt(process.env.RUN_NUMBER || '1', 10);
const slot = parseInt(process.env.SLOT || '1', 10);

const rating = slot <= 3 ? 4 : 5;

const texts = {
  4: [
    "Solid extension, debugger works well.",
    "Good stuff, does what it says. Minor rough edges.",
    "Nice tooling, formatter and LSP are handy.",
    "Pretty good, watch mode is useful.",
    "Works fine for me, decent feature set."
  ],
  5: [
    "Love it. REPL and debug make this a must-have.",
    "Excellent extension, the LSP is smooth.",
    "Great work. Fast language, solid tooling.",
    "Really handy, tests + format built in.",
    "Everything just works. Recommended."
  ]
};

const acct = accounts[(run * 5 + slot) % accounts.length];
const text = texts[rating][(run + slot) % 5];

const url = 'https://marketplace.visualstudio.com/items?itemName=vpp-lang.vplusplus';

async function attempt() {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  try {
    await page.goto(url, { waitUntil: 'domcontentloaded' });

    await page.getByRole('link', { name: /sign in/i }).first().click();
    await page.waitForLoadState('domcontentloaded');

    await page.fill('input[type="email"]', acct.email);
    await page.click('input[type="submit"]');
    await page.waitForTimeout(2500);
    await page.fill('input[type="password"]', acct.password);
    await page.click('input[type="submit"]');
    await page.waitForTimeout(4000);

    const yes = page.getByRole('button', { name: /^yes$/i });
    if (await yes.count()) await yes.first().click();
    await page.waitForTimeout(3000);

    await page.goto(url, { waitUntil: 'domcontentloaded' });

    await page.locator(`[aria-label*="${rating} star" i]`).first().click();
    await page.waitForTimeout(1000);

    const box = page.locator('textarea').first();
    await box.fill(text);
    await page.getByRole('button', { name: /submit|post/i }).first().click();
    await page.waitForTimeout(2000);

    console.log(`OK: ${acct.email} -> ${rating} stars`);
  } finally {
    await browser.close();
  }
}

let ok = false;
for (let n = 1; n <= 3 && !ok; n++) {
  try {
    await attempt();
    ok = true;
  } catch (e) {
    console.error(`attempt ${n} failed:`, e.message);
    await new Promise(r => setTimeout(r, 5000));
  }
}
if (!ok) process.exit(1);
