import csv
from datetime import datetime
import os
from scrapling.fetchers import Fetcher

# Define file paths
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(OUTPUT_DIR, "oil_geopolitical_timeline_2022_2026.csv")

# A curated list of major geopolitical events and India oil import details (2022-2026)
# mapped by date. Note: dates are formatted as YYYY-MM-DD.
GEOPOLITICAL_TIMELINE = {
    "2022-02-24": {
        "event": "Russia invades Ukraine",
        "category": "Sanctions & Supply Shock",
        "notes": "Brent oil spikes past $100/barrel. West plans bans on Russian energy."
    },
    "2022-03-08": {
        "event": "US bans imports of Russian oil and gas",
        "category": "Sanctions",
        "notes": "US implements immediate ban. UK and EU outline plans to phase out dependency."
    },
    "2022-05-30": {
        "event": "EU agrees ban on Russian oil imports",
        "category": "Sanctions",
        "notes": "EU leaders agree to cut 90% of Russian oil imports by end of the year."
    },
    "2022-12-05": {
        "event": "EU/G7 Price Cap ($60/bbl) on Russian crude takes effect",
        "category": "Sanctions & Indian Sourcing",
        "notes": "EU embargo on seaborne imports starts. India leverages steep discounts, boosting imports from Russia."
    },
    "2023-02-05": {
        "event": "EU/G7 Price Cap on Russian refined products takes effect",
        "category": "Sanctions",
        "notes": "Ban and price cap on products like diesel, gasoline, and jet fuel start."
    },
    "2023-04-02": {
        "event": "OPEC+ announces surprise production cuts",
        "category": "OPEC+ Policy",
        "notes": "Surprise voluntary cuts of ~1.16 million barrels per day spark price spike."
    },
    "2023-10-07": {
        "event": "Israel-Hamas war begins",
        "category": "Geopolitical Tension",
        "notes": "Oil prices jump on fears of conflict spreading to key Middle Eastern oil producers."
    },
    "2023-11-19": {
        "event": "Red Sea shipping attacks by Houthi rebels begin",
        "category": "Transit Chokepoint Risk",
        "notes": "Hijack of the Galaxy Leader. Tankers reroute around Africa's Cape of Good Hope, raising shipping costs."
    },
    "2024-04-13": {
        "event": "Iran launches direct drone and missile strikes on Israel",
        "category": "Transit Chokepoint Risk",
        "notes": "Tensions spike in Middle East. Threat to Strait of Hormuz pushes oil risk premium higher."
    },
    "2024-06-02": {
        "event": "OPEC+ extends production cuts",
        "category": "OPEC+ Policy",
        "notes": "OPEC+ agrees to extend cuts but details plan to phase out some voluntary cuts by late 2024."
    },
    "2026-01-15": {
        "event": "West Asia combat risk escalates (Iran Strait risk)",
        "category": "Transit Chokepoint Risk",
        "notes": "Combat involving Iran revives Strait of Hormuz transit fears. Indian refiners pay a premium on Russian oil."
    },
    "2026-03-15": {
        "event": "Russian oil share in India's total imports surges to 33%",
        "category": "Indian Sourcing Shift",
        "notes": "India secures waivers and adjusts payment channels, maintaining Russian oil as top source."
    },
    "2026-05-20": {
        "event": "OPEC+ and IEA supply/demand paradox debate",
        "category": "OPEC+ Policy & Indian Sourcing",
        "notes": "IEA projects 2026 surplus; OPEC sees balance. Russia remains top supplier to India (~38.6% of shipments)."
    }
}

def main():
    print("Fetching Brent Crude historical oil prices from GitHub Raw using Scrapling...")
    
    url = "https://raw.githubusercontent.com/datasets/oil-prices/master/data/brent-daily.csv"
    response = Fetcher.get(url)
    
    if not response.body:
        print("Error: Could not retrieve oil price data from raw GitHub.")
        return

    content_str = response.body.decode("utf-8")
    lines = content_str.strip().split("\n")
    reader = csv.reader(lines)
    header = next(reader) # Date, Price
    
    raw_data = []
    for row in reader:
        if len(row) < 2:
            continue
        date_str, price_str = row[0], row[1]
        raw_data.append((date_str, price_str))
        
    print(f"Total raw records fetched: {len(raw_data)}")
    
    # Filter for the timeline: 2022-01-01 to 2026-05-23 (per research scope)
    start_date = datetime.strptime("2022-01-01", "%Y-%m-%d")
    end_date = datetime.strptime("2026-05-23", "%Y-%m-%d")
    
    filtered_data = []
    for date_str, price_str in raw_data:
        try:
            current_date = datetime.strptime(date_str, "%Y-%m-%d")
            if start_date <= current_date <= end_date:
                filtered_data.append([date_str, price_str])
        except ValueError:
            continue
            
    print(f"Records in scope (2022-2026): {len(filtered_data)}")
    
    # Clean and forward-fill market closed values (represented by "." or empty if any)
    last_valid_price = "80.0"  # default placeholder just in case
    for item in filtered_data:
        price = item[1].strip()
        if not price or price == ".":
            item[1] = last_valid_price
        else:
            last_valid_price = price
            
    # Merge with geopolitical events
    final_rows = []
    for date_str, price in filtered_data:
        event_info = GEOPOLITICAL_TIMELINE.get(date_str, {
            "event": "None",
            "category": "None",
            "notes": "No major sourcing change or conflict escalation on this date."
        })
        
        final_rows.append({
            "Date": date_str,
            "Brent_Crude_Price_USD": price,
            "Geopolitical_Event": event_info["event"],
            "Event_Impact_Category": event_info["category"],
            "India_Import_Notes": event_info["notes"]
        })
        
    # Write to CSV
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    fieldnames = ["Date", "Brent_Crude_Price_USD", "Geopolitical_Event", "Event_Impact_Category", "India_Import_Notes"]
    with open(CSV_PATH, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(final_rows)
        
    print(f"Successfully created: {CSV_PATH}")
    print(f"Total exported rows: {len(final_rows)}")

if __name__ == "__main__":
    main()
