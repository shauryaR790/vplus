import { chromium } from 'playwright';
import fs from 'fs';

const accounts = JSON.parse(fs.readFileSync('.github/workflows/accounts.json', 'utf8'));

const run = parseInt(process.env.RUN_NUMBER || '1', 10);
const slot = parseInt(process.env.SLOT || '1', 10);

// slots 1,2,3 -> 4 stars ; slots 4,5 -> 5 stars
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

// rotates through accounts every run so it "keeps altering"
const acct = accounts[(run * 5 + slot) % accounts.length];
const text = texts[rating][(run + slot) % 5];

const url = 'https://marketplace.visualstudio.com/items?itemName=vpp-lang.vplusplus';

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

  // click the star rating
  await page.locator(`[aria-label*="${rating} star" i]`).first().click();
  await page.waitForTimeout(1000);

  // type the review and submit (selector may need tweaking)
  const box = page.locator('textarea').first();
  await box.fill(text);
  await page.getByRole('button', { name: /submit|post/i }).first().click();
  await page.waitForTimeout(2000);

  console.log(`OK: ${acct.email} -> ${rating} stars`);
} catch (e) {
  console.error('FAIL', acct.email, e.message);
  process.exitCode = 1;
} finally {
  await browser.close();
}
