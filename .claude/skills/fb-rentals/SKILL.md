---
name: fb-rentals
description: Scrape Facebook rental group posts and generate a ranked Excel tracker. Usage: /fb-rentals [hours] or /fb-rentals setup
---

# FB Rentals Skill

You are executing the `/fb-rentals` skill. Follow these steps precisely using your available tools (PowerShell, Read, Write, Glob).

## Parse Arguments

- If the user typed `/fb-rentals setup` → jump to **SETUP WIZARD** below
- If the user typed `/fb-rentals <number>` (e.g. `/fb-rentals 48`) → use that number as `HOURS`
- Otherwise → use `HOURS = 24` (default)

The project root is the directory containing this `.claude/` folder.

---

## SETUP WIZARD

Run this when `config.json` does not exist or when explicitly invoked with `setup`.

Ask the user these questions **one at a time**, waiting for their reply before asking the next:

1. "**Facebook Group URLs** — Paste the URL(s) of the Facebook rental groups you want to monitor (one per line):"
2. "**Budget** — What is your monthly budget range in ₹? (e.g. `24000-26000` for ₹24k–₹26k). I will apply a ±10% buffer automatically."
3. "**Office Location** — Which area/locality is your office in? (e.g. `Gachibowli`, `HITEC City`). Used to rank listings by commute distance."
4. "**Preferred Areas** — Any specific areas/localities you prefer? (comma-separated, or press Enter to skip):"
5. "**Areas to Avoid** — Any areas to avoid? (comma-separated, or press Enter to skip):"
6. "**Room Type** — What are you looking for? Reply with a number: `1` Single room  `2` Master bedroom  `3` Any private room  `4` No preference"
7. "**Diet** — `1` Non-vegetarian  `2` Vegetarian  `3` No preference"
8. "**Gated Community** — `1` Must have  `2` Preferred (show all, flag if not gated)  `3` No preference"
9. "**Look-back window** — How many hours back should each run check? (default: `24`)"

After collecting all answers, write `config.json` with this structure (fill in the user's answers):

```json
{
  "group_urls": ["<url1>"],
  "time_window_hours": 24,
  "budget_ideal_min": 24000,
  "budget_ideal_max": 26000,
  "budget_hard_min": 21600,
  "budget_hard_max": 28600,
  "office_location": "Gachibowli",
  "preferences": {
    "gated_community": "preferred",
    "diet": "non-vegetarian",
    "gender": "male",
    "preferred_areas": [],
    "avoid_areas": [],
    "room_type": "any_private",
    "preferred_flat_types": ["2BHK", "3BHK", "4BHK"]
  },
  "output_file": "listings.xlsx"
}
```

Compute: `budget_hard_min = round(budget_ideal_min * 0.9)` and `budget_hard_max = round(budget_ideal_max * 1.1)`.

Room type mappings: 1→"single", 2→"master", 3→"any_private", 4→"any"
Diet mappings: 1→"non-vegetarian", 2→"vegetarian", 3→"any"
Gated mappings: 1→"required", 2→"preferred", 3→"any"

After writing config.json, tell the user: "Setup complete! Run `/fb-rentals` to start scraping."
Then **stop** — do not continue to the scraping steps.

---

## MAIN FLOW

### STEP 1 — Pre-flight checks

**1a.** Use the Read tool to read `config.json`. If it does not exist, tell the user to run `/fb-rentals setup` first and stop.

**1b.** Use Glob to check for `excel_generator.py`. If it is missing, tell the user the file is missing and stop.

**1c.** Activate the virtual environment and install/verify dependencies via PowerShell:
```
.venv\Scripts\Activate.ps1; pip install openpyxl python-dateutil --quiet
```

---

### STEP 2 — Export existing annotations (if listings.xlsx exists)

Use Glob to check if `listings.xlsx` exists. If it does, run in PowerShell:
```
python excel_generator.py --export-feedback
```
This writes `existing_data.json`. Read that file and note all rows where `feedback` is non-empty — these are your **HIL calibration examples** for Steps 4b and 5.

---

### STEP 3 — Scrape using Chrome MCP

*All browser interaction is done via Chrome DevTools MCP.

**3a. Open group tabs:**
For each URL in `group_urls`, call `mcp__chrome-devtools__new_page` with that URL.

Tell the user: "Facebook group(s) are now open in Chrome. Please log in if needed, solve any CAPTCHA, and confirm you can see the group posts. Then reply 'ready' to continue."

Wait for the user to reply before proceeding.

**3b. Extract posts via `evaluate_script`:**
For each group tab (use `mcp__chrome-devtools__select_page` to switch to it), call `mcp__chrome-devtools__evaluate_script` with the following JS. Replace `HOURS` with the actual number:

```javascript
async () => {
  const HOURS = HOURS_PLACEHOLDER;
  const cutoff = Date.now() - HOURS * 60 * 60 * 1000;

  // Scroll to top first, then slowly scroll down to load posts
  window.scrollTo(0, 0);
  await new Promise(r => setTimeout(r, 2000));

  for (let i = 0; i < 30; i++) {
    window.scrollBy(0, 700);
    await new Promise(r => setTimeout(r, 1000));
    // Click "See more" buttons as we scroll to expand truncated posts
    document.querySelectorAll('div[role="button"], span[role="button"]').forEach(el => {
      if (el.innerText.trim().toLowerCase() === 'see more') el.click();
    });
  }

  // Wait for expansions to settle
  await new Promise(r => setTimeout(r, 1500));

  const results = [];
  const seenUrls = new Set();
  const seenTexts = new Set();

  // Facebook now uses div[role="feed"] > div as feed item containers
  const feedItems = document.querySelectorAll('div[role="feed"] > div');

  feedItems.forEach((item, idx) => {
    // Get longest dir="auto" text block (the post body)
    let text = '';
    item.querySelectorAll('div[dir="auto"]').forEach(el => {
      const t = el.innerText.trim();
      if (t.length > text.length) text = t;
    });
    if (text.length < 40) return;

    // Deduplicate by first 80 chars of text
    const textKey = text.slice(0, 80);
    if (seenTexts.has(textKey)) return;
    seenTexts.add(textKey);

    // Post URL — strip query params to get clean post ID URL
    let url = '';
    item.querySelectorAll('a[href]').forEach(el => {
      const href = (el.href || '').split('?')[0];
      if (href.includes('/posts/') && !url) url = href;
    });
    if (!url) url = window.location.href.split('?')[0] + '#item-' + idx;
    if (seenUrls.has(url)) return;
    seenUrls.add(url);

    // Author — first non-group, non-post Facebook profile link
    let author = '';
    item.querySelectorAll('a[href]').forEach(el => {
      if (author) return;
      const href = el.href || '';
      const t = el.innerText.trim();
      if (t.length > 1 && t.length < 80
          && !href.includes('/groups/')
          && !href.includes('/posts/')
          && (href.includes('facebook.com') || href.includes('fb.com'))) {
        author = t;
      }
    });

    // Timestamp — parse relative time labels like "2h", "30m", "1d"
    let timestamp_iso = '';
    item.querySelectorAll('span, a').forEach(el => {
      if (timestamp_iso) return;
      const t = (el.innerText || '').trim();
      const now = Date.now();
      let ms = 0;
      if (t.match(/^(\d+)m$/)) ms = parseInt(t) * 60000;
      else if (t.match(/^(\d+)h$/)) ms = parseInt(t) * 3600000;
      else if (t.match(/^(\d+)d$/)) ms = parseInt(t) * 86400000;
      if (ms > 0) timestamp_iso = new Date(now - ms).toISOString();
    });

    // Skip if clearly outside time window (only if we have a timestamp)
    if (timestamp_iso) {
      const ts = new Date(timestamp_iso).getTime();
      if (ts < cutoff) return;
    }

    results.push({ url, author, timestamp_iso, text: text.slice(0, 4000) });
  });

  return JSON.stringify(results);
}
```

**3c. Collect and save:**
Collect the JSON results from all group tabs, merge into one array (deduplicate by `url`), and write to `raw_posts.json` using the Write tool.

If the merged array has 0 elements, tell the user no posts were found and suggest verifying that the group is visible when logged in.

---

### STEP 4 — Analyze posts (your intelligent layer)

Read the full contents of `raw_posts.json` and `config.json`. Load `existing_data.json` if it exists.

For **each post** in raw_posts.json, apply your language understanding to extract these structured fields. Process every post before writing any output.

**Fields to extract per post:**

| Field | Notes |
|-------|-------|
| `is_room_listing` | true if this is a room/flat-for-rent post; false for anything else |
| `society_name` | Building or society name, empty string if unknown |
| `area_locality` | Neighbourhood/area, empty string if unknown |
| `flat_type` | "2BHK", "3BHK", "4BHK", "Studio", "1BHK", "Unknown" |
| `room_type` | "Master", "Semi-Master", "Single", "Shared", "Full Flat", "Unknown" |
| `rent` | Monthly rent as integer (₹), 0 if not found |
| `maintenance` | Monthly maintenance as integer (₹), 0 if not mentioned |
| `total_monthly` | rent + maintenance |
| `deposit_months` | Number of months as integer, 0 if not mentioned |
| `brokerage` | "None", "Applicable", "₹X" (specific amount), or "Unknown" |
| `available_from` | Date string, "Immediate", or "Unknown" |
| `furnishing` | "Fully Furnished", "Semi-Furnished", "Unfurnished", "Unknown" |
| `gated_community` | "Yes", "No", "Unknown" — infer from "gated society/community/apartment", "independent house/villa" implies No |
| `amenities_summary` | 1–2 sentence plain-English summary of notable amenities |
| `contact_number` | Phone number string, or empty string |
| `google_maps_link` | Full URL if present in text, or empty string |
| `veg_only` | "Yes" if post explicitly says veg/vegetarian only; "No" if NV welcome; "Unknown" otherwise |
| `bachelor_friendly` | "Yes" if explicitly bachelor-friendly; "No" if no bachelors; "Unknown" otherwise |

**Hard filter — drop post entirely (do not include in output) if:**
- `is_room_listing = false`
- `total_monthly > 0` AND `total_monthly` is outside `budget_hard_min`–`budget_hard_max` from config

---

### STEP 4b — Apply HIL feedback calibration

Before scoring, review the HIL feedback entries from Step 2. Use them as few-shot examples to calibrate your analysis:
- `"not broker"` → treat similar phone/template patterns more leniently in this run
- `"is broker"` → treat similar patterns as stronger signals
- `"relevant - ignore veg"` or similar → note manual approval, don't penalise
- Any other free text → read and use as context for your judgment

---

### STEP 5 — Broker detection (cross-post analysis)

After extracting all individual posts, run these cross-post checks:

**5a. Phone repeats:** Build a map of `contact_number → [list of post urls]`. If the same non-empty number appears in 2+ posts → flag all those posts: `broker_flag = "Suspected"`, `broker_reason = "Same phone in X posts"`.

**5b. Same author, multiple listings:** Count posts per `author`. If author appears in 3+ posts → `broker_flag = "Likely"`, `broker_reason = "Author posted X listings"`.

**5c. Brokerage explicitly mentioned:** If the raw_text contains "brokerage applicable", "brokerage:", "broker fee", or "brokerage charges" (case-insensitive) → upgrade flag to at least `"Doubt"`. Append to broker_reason.

**5d. Template similarity:** If two posts share the same exact ALLCAPS section headers (e.g., "ROOMS AVAILABLE", "FINANCIALS & TERMS", "PLEASE CONTACT") → flag both as `"Suspected"` and note "Template-style post".

**5e. HIL overrides:** If a post URL appears in existing_data.json with feedback `"not broker"` → set `broker_flag = "None"`. If feedback is `"is broker"` → set `broker_flag = "Likely"`.

**Flag precedence (highest wins):** Likely > Suspected > Doubt > None. Never downgrade.

---

### STEP 6 — Rank and score each post (0–100)

Read `office_location` from config.json. For each post compute `match_score`:

| Component | Max pts | Logic |
|-----------|---------|-------|
| Budget fit | 30 | 30 if total_monthly is within ideal range. Scales to 0 at the hard limits. Formula: `max(0, 30 * (1 - abs(total_monthly - midpoint) / half_span))` where midpoint = (ideal_min+ideal_max)/2, half_span = (ideal_max-ideal_min)/2 + (hard_max-ideal_max). If total_monthly is 0/unknown, give 15 (neutral). |
| Commute | 25 | Compare area_locality to office_location using your city geography knowledge. Same area or adjacent=25, close <5km=18, moderate 5–15km=12, far >15km=5, unknown=10 |
| Gated community | 15 | Yes=15, Unknown=8, No=0 |
| Broker confidence | 15 | None=15, Doubt=9, Suspected=4, Likely=0 |
| Room type match | 10 | Compare room_type to preference from config. Exact=10, close alternative=5, mismatch=0 |
| Soft warnings | 5 | Start 5. Subtract 2.5 if veg_only=Yes. Subtract 2.5 if bachelor_friendly=No. Min 0. |

Set `commute_estimate` as: "Same area", "Close (<5km)", "Moderate (5–15km)", "Far (>15km)", or "Unknown".

Round `match_score` to nearest integer.

---

### STEP 7 — Write analyzed_posts.json

Write the full analyzed output as a JSON array. Each element must include ALL of these keys:

```
url, author, timestamp_iso, is_room_listing, society_name, area_locality, flat_type,
room_type, rent, maintenance, total_monthly, deposit_months, brokerage, available_from,
furnishing, gated_community, amenities_summary, contact_number, google_maps_link,
veg_only, bachelor_friendly, match_score, commute_estimate, broker_flag, broker_reason
```

Only include posts that passed the hard filter (is_room_listing=true and within budget).

---

### STEP 8 — Generate Excel

Run in PowerShell:
```
python excel_generator.py --generate
```

If it errors, show the full error output.

---

### STEP 9 — Summary

Print a summary block like this:

```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  FB Rentals — Run complete
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Posts scraped (raw):   42
  After filters:         28
  Broker-flagged:         6   (Likely: 2 | Suspected: 3 | Doubt: 1)
  Top match score:       87
  Output: listings.xlsx
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Top 3 listings:
  1. [87] Prestige Lakeside, Kondapur — ₹25,000/mo — Master room
  2. [82] Aparna Serene, Gachibowli — ₹24,500/mo — Semi-Master — Doubt broker
  3. [76] Rainbow Residency, Madhapur — ₹23,800/mo — Single
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

Then tell the user:
"Open `listings.xlsx` to review. Use the **Status** column (dropdown) to mark each listing. Use the **Feedback** column to correct my broker classifications — I will read these on your next run to improve accuracy."

---

## Edge case notes

- If Facebook shows a CAPTCHA in Chrome, the user can solve it in the browser window and then reply 'ready' to continue.
- If `raw_posts.json` has 0 posts, the group may require membership approval or Facebook's DOM selectors may have changed. Ask the user to verify they can see the group posts when logged in, and check if `div[role="feed"]` and its direct `div` children are present using the browser's DevTools. The selector `div[role="feed"] > div` is what the scraper relies on.
- On subsequent runs, Chrome may already be logged in — the user can reply 'ready' immediately without logging in again.
- On the very first run, `existing_data.json` will not exist — skip Step 2 gracefully.
- If a post's `total_monthly` is 0 (price not mentioned), include it in output but apply a neutral budget score of 15. The user can decide.
