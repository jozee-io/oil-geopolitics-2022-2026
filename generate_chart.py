import csv
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime
import os

# Define file paths
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(OUTPUT_DIR, "oil_geopolitical_timeline_2022_2026.csv")
CHART_PATH = os.path.join(OUTPUT_DIR, "oil_geopolitical_timeline_chart.png")

# Key events to annotate on the chart
ANNOTATIONS = [
    ("2022-02-24", "Russia invades Ukraine\n(Brent spikes past $100)"),
    ("2022-12-05", "EU/G7 G7 Price Cap ($60)"),
    ("2023-10-07", "Israel-Hamas war begins"),
    ("2023-11-19", "Red Sea shipping attacks"),
    ("2024-04-13", "Iran drone strikes Israel"),
    ("2026-01-15", "West Asia Strait Risk\n(Iran combat risk rises)"),
    ("2026-03-15", "India Russian import share (33%)")
]

def main():
    if not os.path.exists(CSV_PATH):
        print(f"Error: {CSV_PATH} not found. Please run the scraper first.")
        return

    dates = []
    prices = []
    
    # Read the scraped CSV data
    with open(CSV_PATH, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            try:
                date_val = datetime.strptime(row["Date"], "%Y-%m-%d")
                price_val = float(row["Brent_Crude_Price_USD"])
                dates.append(date_val)
                prices.append(price_val)
            except ValueError:
                continue

    print(f"Loaded {len(dates)} price points for charting.")

    # Setup style: clean, modern off-white background with subtle grids
    plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
    
    fig, ax = plt.subplots(figsize=(15, 8), dpi=300)
    fig.patch.set_facecolor('#fafafb')
    ax.set_facecolor('#ffffff')
    
    # Plot Brent Crude Price line
    ax.plot(dates, prices, color='#0f172a', linewidth=2.0, label='Brent Crude Price (USD/bbl)', alpha=0.9)
    
    # Format axes
    ax.set_title("Brent Crude Oil Price & Key Geopolitical Event Milestones (2022-2026)", fontsize=16, fontweight='bold', pad=20, color='#1e293b')
    ax.set_xlabel("Timeline", fontsize=12, fontweight='semibold', labelpad=10, color='#334155')
    ax.set_ylabel("Brent Crude Spot Price (USD/bbl)", fontsize=12, fontweight='semibold', labelpad=10, color='#334155')
    
    # Set date format on x-axis
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
    ax.xaxis.set_minor_locator(mdates.MonthLocator(interval=3))
    
    # Grid customization
    ax.grid(True, which='major', linestyle='--', linewidth=0.7, color='#e2e8f0')
    ax.grid(True, which='minor', linestyle=':', linewidth=0.5, color='#f1f5f9')
    
    # Y-axis limits
    ax.set_ylim(min(prices) - 5, max(prices) + 15)
    
    # Add annotations for key events
    for date_str, text in ANNOTATIONS:
        event_date = datetime.strptime(date_str, "%Y-%m-%d")
        
        # Find closest price for the event date to place the annotation pointer
        if event_date in dates:
            idx = dates.index(event_date)
            event_price = prices[idx]
        else:
            # Fallback to closest available date
            closest_date = min(dates, key=lambda d: abs(d - event_date))
            idx = dates.index(closest_date)
            event_price = prices[idx]
            
        # Draw indicator dot
        ax.plot(event_date, event_price, marker='o', markersize=8, color='#dc2626', markeredgecolor='#ffffff', markeredgewidth=1.5, zorder=5)
        
        # Offset logic for labels to avoid overlaps
        # Alternate text positions above and below the line
        if date_str in ["2022-02-24", "2023-10-07", "2024-04-13", "2026-03-15"]:
            xytext_offset = (0, 35)
        else:
            xytext_offset = (0, -45)
            
        ax.annotate(
            text,
            xy=(event_date, event_price),
            xytext=xytext_offset,
            textcoords='offset points',
            arrowprops=dict(arrowstyle="-|>", connectionstyle="arc3,rad=0.1", color='#ef4444', lw=1.2),
            fontsize=9.5,
            fontweight='bold',
            color='#1e293b',
            bbox=dict(boxstyle="round,pad=0.3", fc='#ffffff', ec='#cbd5e1', lw=0.8, alpha=0.95),
            ha='center'
        )

    plt.tight_layout()
    plt.savefig(CHART_PATH, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    
    print(f"Chart successfully saved to: {CHART_PATH}")

if __name__ == "__main__":
    main()
