# FB Rental Tracker — Claude Code Skill

A Claude Code slash command (`/fb-rentals`) that scrapes Facebook rental group posts, ranks them by how well they match your preferences, detects spam brokers, and exports everything to a tracked Excel file.

**Powered by Claude's language understanding** — no rigid regex, no API key needed beyond your Claude Code session.

---

## What it does

- Scrapes one or more Facebook group(s) for room/flat listings posted in the last N hours
- Extracts structured data from each post (rent, maintenance, deposit, location, amenities, contact, Maps link)
- **Ranks listings 0–100** based on budget fit, commute distance to your office, broker risk, gated community, and room type match
- **Detects broker/spam posts** via phone number repeats, identical templates, multiple listings per author, and explicit brokerage mentions
- Exports a colour-coded Excel with Status and Notes dropdowns for tracking
- **Learns from your feedback** — write corrections in the Feedback column; Claude reads them on the next run

---

## Prerequisites

- [Claude Code](https://claude.ai/code) installed and authenticated
- [Chrome DevTools MCP](https://github.com/nickytonline/chrome-devtools-mcp) configured in your `.mcp.json`
- Python 3.10+
- A Facebook account with access to the rental group(s)

---

## Setup

### 1. Clone the repo

```bash
git clone https://github.com/hruthikvs/fb-rental-tracker.git
cd fb-rental-tracker
```

### 2. Create a virtual environment and install dependencies

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
```

### 3. Open Claude Code in this folder

```bash
claude
```

### 4. Run the setup wizard

```
/fb-rentals setup
```

Claude will ask you:
- Facebook group URL(s) to monitor
- Your monthly budget (e.g. `24000-26000` for ₹24k–₹26k; a ±10% buffer is applied automatically)
- Your office location (for commute ranking)
- Preferred/avoided areas
- Room type, diet, gated community preference
- Look-back window (default: 24 hours)

This generates `config.json`. **Do not commit config.json** — it contains your personal preferences.

---

## Usage

### Run a scrape (last 24 hours)
```
/fb-rentals
```

### Run for a custom window
```
/fb-rentals 48
```

### Re-run setup / change preferences
```
/fb-rentals setup
```

---

## What happens when you run it

1. Claude opens the Facebook group in Chrome via the Chrome DevTools MCP
2. You log in to Facebook in the browser (your credentials never touch this tool), then reply "ready"
3. Claude scrolls the feed, expands truncated posts, and extracts all listings
4. Claude reads every post and extracts: society, area, rent, deposit, brokerage, amenities, contact, Maps link
5. Cross-post broker detection flags repeated phone numbers, template copy-pastes, and prolific posters
6. Every listing gets a **Match Score** (0–100) based on your config
7. `listings.xlsx` is written (or updated — your Status/Notes/Feedback are always preserved)

---

## Excel columns

| Column | Description |
|--------|-------------|
| Match Score | 0–100 rank (sorted descending) |
| Post Link | Hyperlink to the original Facebook post |
| Society / Building | Building or society name |
| Area / Locality | Neighbourhood |
| Commute Estimate | Same area / Close / Moderate / Far |
| Flat Type | 2BHK / 3BHK / 4BHK / etc. |
| Room Type | Master / Semi-Master / Single / etc. |
| Rent / Maintenance / Total | Monthly costs |
| Deposit | In months |
| Brokerage | None / Applicable / ₹X |
| Gated Community | Yes / No / Unknown |
| Broker Flag | None / Doubt / Suspected / Likely (colour coded) |
| **Status** | Dropdown: Visited / Interested / Rejected / Shortlisted |
| **Notes** | Free text for your notes |
| **Feedback** | Write corrections here — Claude reads on next run |

Broker flag row colours: 🔴 Likely · 🟠 Suspected · 🟡 Doubt

---

## Sample output

### The raw Facebook post

![Facebook post — Prestige High Fields, part 1](sample_run/Screenshot%202026-05-07%20014948.png)
![Facebook post — Prestige High Fields, part 2](sample_run/Screenshot%202026-05-07%20015300.png)

> Posted by **-** in *Flat and Flatmates (Hyderabad)*

### What the tool extracts and outputs in Excel

| Column | Value |
|--------|-------|
| **Match Score** | 81 |
| **Post Link** | [View Post](https://www.facebook.com/groups/320292845738195/posts/1630054568095343/) |
| **Society / Building** | Prestige High Fields |
| **Area / Locality** | Financial District |
| **Commute Estimate** | Same area |
| **Flat Type** | 3BHK |
| **Room Type** | Single |
| **Rent** | ₹18,500 |
| **Maintenance** | ₹0 (included) |
| **Total Monthly** | ₹18,500 |
| **Deposit** | 3 months |
| **Brokerage** | None |
| **Available From** | May 24th |
| **Furnishing** | Fully Furnished |
| **Gated Community** | Yes |
| **Broker Flag** | None |
| **Amenities** | Premium clubhouse, gym, swimming pool, sauna, tennis/badminton/squash courts, private cinema, café, 24/7 gated security. 31st floor with balcony view. |

Claude scores this **81/100** — high marks for being in the same area as the office (Financial District ↔ Raidurg), premium gated society, and no broker. Slight dip because ₹18,500 is near the lower edge of the ideal budget range.

---

## Feedback / Human-in-the-Loop

The **Feedback** column (light blue) makes the tool smarter over time:

| You write | Effect on next run |
|-----------|-------------------|
| `not broker` | Overrides broker flag to None for that post; similar patterns treated more leniently |
| `is broker` | Confirms broker flag; similar patterns weighted more heavily |
| `relevant - ignore veg` | Marks post as manually approved despite soft warning |

---

## Files

```
├── .claude/skills/fb-rentals/SKILL.md   Claude Code skill definition
├── excel_generator.py                    Excel writer with formatting
├── requirements.txt
├── config.example.json                   Template — copy to config.json and fill in
└── listings.xlsx                         Generated at runtime (gitignored)
```

---

## Notes & limitations

- Facebook's feed DOM uses `div[role="feed"] > div` as post containers. If scraping breaks, this selector is the first thing to check in the browser DevTools.
- This tool uses your logged-in browser session via Chrome DevTools MCP. Your credentials are never stored or transmitted by this code.
- Works best for Indian rental groups (₹ amounts, BHK terminology, Maps links) but the config is adaptable.

---

## License

MIT
