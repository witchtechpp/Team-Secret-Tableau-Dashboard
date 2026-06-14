import sys
import re
import csv
import time
import os
from datetime import datetime
from pathlib import Path

try:
    import cloudscraper
    from bs4 import BeautifulSoup
except ImportError:
    print("Missing dependencies! Please run: pip install cloudscraper beautifulsoup4 openpyxl")
    sys.exit(1)

try:
    import openpyxl

    HAS_EXCEL = True
except ImportError:
    HAS_EXCEL = False

CSV_COLUMNS = [
    "match_id", "event", "map_name", "team", "player",
    "rating", "acs", "kills", "deaths", "assists", "kd_diff",
    "kast", "adr", "hs_pct", "fk", "fd", "agents", "match_url"
]


def get_match_id(url):
    m = re.search(r"vlr\.gg/(\d+)/", url)
    return m.group(1) if m else "unknown"


def parse_with_bs4(tr, team_name):
    cols = tr.find_all("td")
    if not cols or len(cols) < 10:
        return None

    # Extract Player Name
    player_div = cols[0].find("div", class_="text-of")
    player_name = player_div.get_text(strip=True) if player_div else "Unknown"

    # Extract Agents
    agent_imgs = cols[1].find_all("img")
    agents = ", ".join([img.get("title", img.get("alt", "")) for img in agent_imgs])

    def get_text_bs4(col):
        main_val = col.find("span", class_="mod-both")
        if main_val:
            return main_val.get_text(strip=True)
        return col.get_text(strip=True)

    try:
        return {
            "player": player_name,
            "team": team_name,
            "agents": agents,
            "rating": get_text_bs4(cols[2]),
            "acs": get_text_bs4(cols[3]),
            "kills": get_text_bs4(cols[4]).replace("/", "").strip(),
            "deaths": get_text_bs4(cols[5]).replace("/", "").strip(),
            "assists": get_text_bs4(cols[6]).replace("/", "").strip(),
            "kd_diff": get_text_bs4(cols[7]),
            "kast": get_text_bs4(cols[8]).replace("%", ""),
            "adr": get_text_bs4(cols[9]),
            "hs_pct": get_text_bs4(cols[10]).replace("%", ""),
            "fk": get_text_bs4(cols[11]),
            "fd": get_text_bs4(cols[12])
        }
    except Exception:
        return None


def scrape_match_page(scraper, url):
    """Downloads a match page and parses all teams found."""
    try:
        resp = scraper.get(url, timeout=20)
        soup = BeautifulSoup(resp.text, "html.parser")

        event_el = soup.find("a", class_="match-header-event")
        event = event_el.get_text(strip=True) if event_el else ""

        # Extract Team Names from the match header
        team_divs = soup.find_all("div", class_="wf-title-med")
        t1 = team_divs[0].get_text(strip=True) if len(team_divs) > 0 else "Team 1"
        t2 = team_divs[1].get_text(strip=True) if len(team_divs) > 1 else "Team 2"

        data_rows = []
        game_containers = soup.find_all("div", class_="vm-stats-game")

        for game in game_containers:
            map_name_div = game.find("div", class_="map")
            map_name = map_name_div.get_text(strip=True) if map_name_div else "Summary"
            map_name = re.sub(r'^\d+', '', map_name).strip()

            tables = game.find_all("table", class_="wf-table-inset")
            for i, table in enumerate(tables):
                # Identify which team this table belongs to
                current_team_name = t1 if i == 0 else t2

                rows = table.find("tbody").find_all("tr") if table.find("tbody") else []
                for tr in rows:
                    stats = parse_with_bs4(tr, current_team_name)
                    if stats:
                        stats.update({
                            "match_id": get_match_id(url),
                            "event": event,
                            "map_name": map_name,
                            "match_url": url
                        })
                        data_rows.append(stats)
        return data_rows
    except Exception as e:
        print(f"   [!] Failed to parse match {url}: {e}")
        return []


def main():
    print("=== VLR.GG UNIVERSAL BATCH SCRAPER ===")

    # 1. Prompt for URL
    stats_url = input("\n1. Paste VLR Stats/Results URL: ").strip()

    # 2. Setup Automatic Directory (Documents/VLR_Scraped_Stats)
    # Path.home() works on any computer (Windows, Mac, Linux)
    final_dir = Path.home() / "Documents" / "VLR_Scraped_Stats"

    try:
        final_dir.mkdir(parents=True, exist_ok=True)
        print(f"\n2. Files will be saved to: {final_dir}")
    except Exception as e:
        print(f"   [!] Could not create directory, saving to current folder instead. Error: {e}")
        final_dir = Path(".")

    scraper = cloudscraper.create_scraper()

    # --- STEP 1: Find all Match Links ---
    print(f"\n[Phase 1] Scanning page for match links...")
    try:
        list_resp = scraper.get(stats_url)
        list_soup = BeautifulSoup(list_resp.text, "html.parser")
        match_links = []
        for a in list_soup.find_all("a", href=True):
            href = a['href']
            # Match URLs contain digits and 'vs'
            if re.search(r"^/\d+/", href) and "vs" in href:
                full_link = "https://www.vlr.gg" + href
                if full_link not in match_links:
                    match_links.append(full_link)
    except Exception as e:
        print(f"Error getting match list: {e}")
        return

    if not match_links:
        print("No match links found.")
        return

    print(f"Found {len(match_links)} matches. Starting scrape...")

    # --- STEP 2: Scrape each match ---
    all_final_data = []
    for i, m_url in enumerate(match_links, 1):
        print(f"[{i}/{len(match_links)}] Parsing: {m_url}")
        match_stats = scrape_match_page(scraper, m_url)
        all_final_data.extend(match_stats)
        time.sleep(1.5)  # Slight delay to be polite to servers

    # --- STEP 3: Save results ---
    if not all_final_data:
        print(f"\n[!] No data was extracted. Check the URL or your connection.")
        return

    # Create a unique filename based on current time
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"VLR_Full_Export_{timestamp}.xlsx"
    xlsx_path = final_dir / filename

    if HAS_EXCEL:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "VLR Stats"
        ws.append(CSV_COLUMNS)
        for row in all_final_data:
            ws.append([row.get(col, "") for col in CSV_COLUMNS])

        try:
            wb.save(xlsx_path)
            print(f"\n>> SUCCESS! File saved at: {xlsx_path}")
        except PermissionError:
            print(f"\n[!] Error: Could not save file. Please make sure {filename} is not open in Excel.")
    else:
        # Fallback to CSV if openpyxl is missing
        csv_path = xlsx_path.with_suffix('.csv')
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
            writer.writeheader()
            writer.writerows(all_final_data)
        print(f"\n>> Excel module missing. Saved as CSV instead: {csv_path}")

    print(f"Successfully scraped {len(all_final_data)} total player-map rows.")
    input("\nPress Enter to close...")


if __name__ == "__main__":
    main()
