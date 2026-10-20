import { chromium } from 'playwright';
import readline from 'readline/promises';
import fs from 'fs';

const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

// ====== FILL THIS IN - password can stay '' if account is passwordless ======
const accounts = [
  { email: 'mmmmmmaybachmusicdonunwgoup@gmail.com', password: 'shaurya2009', rating: 5, text: 'Debugger and REPL are handy' },
  // ... add up to 10
];
// ===========================================================================

const url = 'https://marketplace.visualstudio.com/items?itemName=vpp-lang.vplusplus';

async function review(acct) {
  const profile = `./profiles/${acct.email}`;
  fs.mkdirSync(profile, { recursive: true });
  const ctx = await chromium.launchPersistentContext(profile, { headless: false, slowMo: 50 });
  const page = ctx.pages()[0] || await ctx.newPage();
  try {
    console.log(`\n=== ${acct.email} ===`);
    await page.goto(url, { waitUntil: 'domcontentloaded' });
    await page.waitForTimeout(3000);

    const needLogin = await page.getByRole('link', { name: /sign in/i }).count();
    if (needLogin) {
      console.log('  logging in...');
      await page.getByRole('link', { name: /sign in/i }).first().click();
      await page.waitForTimeout(3000);
      await page.waitForSelector('#i0116', { timeout: 30000 });
      await page.fill('#i0116', acct.email);
      await page.click('#idSIButton9');
      await page.waitForTimeout(4000);

      if (await page.locator('#i0118').count()) {
        await page.fill('#i0118', acct.password);
        await page.click('#idSIButton9');
        await page.waitForTimeout(4000);
      } else {
        const send = page.getByRole('button', { name: /send code/i });
        if (await send.count()) await send.click();
        await page.waitForTimeout(2000);
        const code = (await rl.question(`  code for ${acct.email}: `)).trim();
        await page.locator('input[name="otc"], #idTxtBx_SAOTCC_OTC').first().fill(code);
        await page.click('#idSIButton9');
        await page.waitForTimeout(4000);
      }
      await rl.question('  solve any challenge in the window, then press Enter... ');
      const yes = page.getByRole('button', { name: /^yes$/i });
      if (await yes.count()) { await yes.first().click(); await page.waitForTimeout(3000); }
      await page.goto(url, { waitUntil: 'domcontentloaded' });
      await page.waitForTimeout(3000);
    } else {
      console.log('  already signed in (profile reused)');
    }

    for (const sel of [`[aria-label="${acct.rating} star"]`, `[aria-label*="${acct.rating} star" i]`]) {
      const el = page.locator(sel).first();
      if (await el.count()) { await el.click(); break; }
    }
    await page.waitForTimeout(1500);
    await page.locator('textarea').first().fill(acct.text);
    await page.waitForTimeout(500);
    await page.getByRole('button', { name: /submit|post|save|publish/i }).first().click();
    await page.waitForTimeout(3000);
    console.log(`OK ${acct.email}`);
  } catch (e) {
    await page.screenshot({ path: `fail-${acct.email}.png` }).catch(() => {});
    console.error(`FAIL ${acct.email}: ${e.message} @ ${page.url()}`);
  } finally {
    await ctx.close();
  }
}

for (const a of accounts) await review(a);
rl.close();
console.log('\nDONE');
