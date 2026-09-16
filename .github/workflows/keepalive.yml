/**
 * ITP Sales Hub keep-alive / auto-waker for Streamlit Community Cloud.
 *
 * Why a real (headless) browser is required, not a simple ping:
 *   A plain GET to https://<app>.streamlit.app/ returns Community Cloud's
 *   dashboard SPA (a static shell) with HTTP 200 whether the app is awake
 *   or asleep. It never touches the actual app container, so a cron job
 *   that just curls the URL registers no traffic and can't even tell if
 *   the app is asleep.
 *
 *   The real app is rendered inside an iframe at
 *   https://<app>.streamlit.app/~/+/, and the sleep screen ("Zzzz... /
 *   Yes, get this app back up!") is rendered by the outer shell with
 *   data-testid="wakeup-button-viewer".
 *
 * So this script visits the app the way a real person would: it loads the
 * page, clicks the wake-up button if the sleep screen is showing, waits
 * for the app to actually render inside that iframe, and stays connected
 * a few seconds so the visit counts as real traffic.
 */
const { chromium } = require("playwright");

// Set this to your app's real Streamlit URL, e.g.
// "https://itp-sales-hub.streamlit.app/" -- or pass it in via the
// APP_URL repo variable/secret referenced in keepalive.yml.
const APP_URL = process.env.APP_URL || "https://REPLACE-WITH-YOUR-APP.streamlit.app/";

const UA =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36";

const NAV_TIMEOUT = 90_000;
const RESOLVE_TIMEOUT = 90_000; // wait for app-or-sleep-screen to appear
const WAKE_TIMEOUT = 300_000; // cold boot (fresh container) can take minutes
const DWELL_MS = 8_000; // stay connected so the session registers as traffic

/** Cheap pre-check: this endpoint reaches the actual container, unlike the app root. */
async function healthCheck(url) {
  const origin = new URL(url).origin;
  const jar = new Map();
  let next = origin + "/";

  // Follow the auth-cookie handshake manually so cookies persist across hops.
  for (let hop = 0; hop < 12; hop++) {
    const cookie = [...jar].map(([k, v]) => `${k}=${v}`).join("; ");
    const res = await fetch(next, {
      redirect: "manual",
      headers: { "User-Agent": UA, ...(cookie && { Cookie: cookie }) },
    });
    for (const sc of res.headers.getSetCookie?.() ?? []) {
      const [pair] = sc.split(";");
      const i = pair.indexOf("=");
      const k = pair.slice(0, i).trim();
      const v = pair.slice(i + 1).trim();
      v === "" ? jar.delete(k) : jar.set(k, v);
    }
    const loc = res.headers.get("location");
    if (!loc) break;
    next = new URL(loc, next).toString();
  }

  const cookie = [...jar].map(([k, v]) => `${k}=${v}`).join("; ");
  try {
    const res = await fetch(`${origin}/~/+/_stcore/health`, {
      headers: { "User-Agent": UA, Cookie: cookie },
      signal: AbortSignal.timeout(30_000),
    });
    return res.ok && (await res.text()).trim() === "ok" ? "ok" : `http ${res.status}`;
  } catch (e) {
    return `unreachable (${e.message})`;
  }
}

/** The Streamlit app itself lives in the /~/+/ iframe. */
function appFrame(page) {
  return page.frames().find((f) => f.url().includes("/~/+/"));
}

async function appRendered(page) {
  const f = appFrame(page);
  if (!f) return false;
  return (await f.locator('[data-testid="stApp"], .stApp').count()) > 0;
}

async function visit(browser, url) {
  const t0 = Date.now();
  const say = (m) => console.log(`   [${((Date.now() - t0) / 1000).toFixed(1)}s] ${m}`);

  const ctx = await browser.newContext({ userAgent: UA, viewport: { width: 1440, height: 900 } });
  const page = await ctx.newPage();

  try {
    say(`health: ${await healthCheck(url)}`);

    await page.goto(url, { waitUntil: "domcontentloaded", timeout: NAV_TIMEOUT });

    const wakeBtn = page.locator(
      '[data-testid="wakeup-button-viewer"], [data-testid="wakeup-button-owner"]'
    );

    // Race: either the app iframe renders, or the sleep screen appears.
    const deadline = Date.now() + RESOLVE_TIMEOUT;
    let state = "unresolved";
    while (Date.now() < deadline) {
      if (await wakeBtn.count()) {
        state = "asleep";
        break;
      }
      if (await appRendered(page)) {
        state = "awake";
        break;
      }
      await page.waitForTimeout(1000);
    }

    if (state === "asleep") {
      say("ASLEEP -- clicking wake button");
      await wakeBtn.first().click();

      const wakeDeadline = Date.now() + WAKE_TIMEOUT;
      while (Date.now() < wakeDeadline) {
        if (await appRendered(page)) {
          say("WOKEN -- app rendered");
          await page.waitForTimeout(DWELL_MS);
          return "WOKEN";
        }
        await page.waitForTimeout(2000);
      }
      say("wake timed out");
      return "WAKE_TIMEOUT";
    }

    if (state === "awake") {
      say("AWAKE -- app already rendered");
      await page.waitForTimeout(DWELL_MS); // dwell so the session counts as traffic
      return "AWAKE";
    }

    say(`UNRESOLVED -- title="${await page.title()}"`);
    return "UNRESOLVED";
  } catch (e) {
    say(`ERROR: ${e.message.split("\n")[0]}`);
    return "ERROR";
  } finally {
    await ctx.close().catch(() => {});
  }
}

(async () => {
  if (APP_URL.includes("REPLACE-WITH-YOUR-APP")) {
    console.error(
      "Set APP_URL (top of keepalive.js, or the APP_URL repo variable) to your real " +
        "Streamlit app URL before this will do anything."
    );
    process.exit(1);
  }

  console.log(`=== ${APP_URL}`);
  const browser = await chromium.launch();
  const status = await visit(browser, APP_URL);
  await browser.close();

  console.log(`\n=== RESULT: ${status} ===`);

  if (process.env.GITHUB_STEP_SUMMARY) {
    const fs = require("fs");
    const icon = { AWAKE: "✅", WOKEN: "🔄", WAKE_TIMEOUT: "⚠️", UNRESOLVED: "⚠️", ERROR: "❌" };
    fs.appendFileSync(
      process.env.GITHUB_STEP_SUMMARY,
      `## ITP Sales Hub keep-alive\n\n${icon[status] ?? ""} **${status}** -- ${APP_URL}\n`
    );
  }

  process.exit(["ERROR", "WAKE_TIMEOUT", "UNRESOLVED"].includes(status) ? 1 : 0);
})();
