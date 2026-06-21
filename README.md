# VLR.gg Batch Stats Scraper

A command-line tool that scrapes player statistics from [VLR.gg](https://www.vlr.gg) match pages. Point it at any results or stats listing URL and it will find every match linked from that page, pull per-player, per-map stats from each one, and export everything to a single Excel (or CSV) file.

This scraper was built to source the match data behind the **Team Secret Valorant Tableau Dashboard** — the raw, uncleaned exports it produces are included in this repo and free for anyone to reuse.

![Python](https://img.shields.io/badge/Python-3.x-blue)
![cloudscraper](https://img.shields.io/badge/scraping-cloudscraper-orange)
![BeautifulSoup](https://img.shields.io/badge/parsing-BeautifulSoup4-yellow)

## Features

- **Batch scraping** — paste one listing URL (e.g. an event's match list) and the script automatically discovers and scrapes every individual match linked from it
- **Per-player, per-map granularity** — each row is one player's stats on one map, including team, agent(s) played, and which match/event it came from
- **Cloudflare-aware** — uses `cloudscraper` instead of plain `requests` to get past VLR.gg's anti-bot protection
- **Automatic export folder** — saves results to `Documents/VLR_Scraped_Stats` on any OS (Windows, macOS, Linux), with a timestamped filename so runs never overwrite each other
- **Graceful fallback** — exports to `.xlsx` via `openpyxl` if available, otherwise falls back to `.csv` automatically
- **Polite scraping** — small delay between match requests to avoid hammering the server

## Data Collected

Each row in the output contains:

| Column | Description |
|---|---|
| match_id | VLR match ID extracted from the URL |
| event | Tournament/event name |
| map_name | Map the stats were recorded on |
| team | Team the player belongs to |
| player | Player name |
| rating | VLR performance rating |
| acs | Average combat score |
| kills / deaths / assists | Raw stat totals |
| kd_diff | Kill/death differential |
| kast | KAST % |
| adr | Average damage per round |
| hs_pct | Headshot % |
| fk / fd | First kills / first deaths |
| agents | Agent(s) played |
| match_url | Source match URL |

## Requirements

- Python 3.8+
- `cloudscraper`
- `beautifulsoup4`
- `openpyxl` (optional — enables `.xlsx` export; the script falls back to `.csv` without it)

## Installation

```bash
git clone https://github.com/witchtechpp/Team-Secret-Tableau-Dashboard.git
cd Team-Secret-Tableau-Dashboard
pip install cloudscraper beautifulsoup4 openpyxl
```

## Usage

Run the script:

```bash
python VLRGGscraper.py
```

You'll be prompted for a URL:

```
1. Paste VLR Stats/Results URL:
```

Paste a VLR.gg page that links to multiple matches — for example an event's match list or results page. The script will:

1. Scan that page for all match links
2. Visit each match page and parse every player's stats, broken down by map
3. Compile everything into one export

Output is saved automatically to:

```
~/Documents/VLR_Scraped_Stats/VLR_Full_Export_<timestamp>.xlsx
```

(or `.csv` if `openpyxl` isn't installed)

## How It Works

1. **Link discovery** — the listing page's HTML is scanned for `<a>` tags whose `href` matches VLR's match URL pattern (`/<digits>/...vs...`)
2. **Per-match parsing** — for each match, the script reads the match header for the event name and both team names, then walks each map's stats table (`vm-stats-game` → `wf-table-inset`) row by row
3. **Per-player extraction** — for each table row, it pulls the player's name, agents played, and combined-side stat columns (rating, ACS, K/D/A, KAST, ADR, HS%, first kills/deaths)
4. **Export** — all collected rows are written to Excel (or CSV) with a consistent column schema

## Known Limitations

- **Tied to VLR.gg's current HTML structure** — the scraper relies on specific class names (`wf-table-inset`, `vm-stats-game`, `mod-both`, etc.). If VLR redesigns their site, the selectors will need updating.
- **Combined-side stats only** — pulls the "both sides" stat column, not the separate attacker/defender splits also shown on VLR.
- **Sequential, not parallel** — matches are scraped one at a time with a delay between requests, so large batches (entire events/seasons) can take a while.
- **No retry logic** — a failed request for a single match is skipped and logged to the console rather than retried.
- **Interactive only** — the script prompts for input rather than accepting CLI arguments, so it isn't currently designed for automation/scheduling without modification.

## Disclaimer

This tool is intended for personal, non-commercial use (e.g. building dashboards or doing personal analysis of match data). Be mindful of VLR.gg's terms of service and avoid scraping at a volume or frequency that could impact their servers.

## License

This project is available for personal and educational use.

## Author

**Shannen Anabo**
[GitHub Repo](https://github.com/witchtechpp/Team-Secret-Tableau-Dashboard) · [Tableau Dashboard](https://public.tableau.com/app/profile/shannen.anabo)
