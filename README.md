# Geopolitical Conflicts & The Crude Oil Economy (2022–2026)

This repository contains the dataset, scripts, and slide deliverables for the **Geopolitical Oil Research** project. 

The project investigates crude oil not simply as a physical commodity, but as a geopolitical transmission mechanism—analyzing how G7 sanctions, Middle East chokepoint transit risks (Strait of Hormuz, Bab al-Mandab), and OPEC+ decisions impact global Brent Crude benchmarks, refining crack spreads, and corporate hedging strategies.

---

## Project Objectives

1. **Quantify Geopolitical Risk Premiums:** Distinguish between actual physical supply shocks (e.g., initial Ukraine disruptions) and sentiment-driven risk premiums (e.g., direct Iran-Israel drone strikes).
2. **Trace Transmission Channels:** Document the cost transmission through maritime logistics (increased ton-miles around the Cape of Good Hope, freight rates, and hull insurance).
3. **Analyze Sourcing Pragmatism (The India Angle):** Examine how emerging economies (such as India) leveraged G7 sanction waivers to import discounted Russian Urals crude (reaching ~33-38.6% of India's import mix in early 2026) to manage domestic inflation.
4. **Corporate Margin Analysis:** Evaluate fuel cost transmission and hedging strategies inside energy-intensive sectors (with a focus on the Aviation industry).

---

## File Structure

*   📊 **[oil_geopolitical_timeline_2022_2026.csv](oil_geopolitical_timeline_2022_2026.csv)**: The complete daily dataset (1,108 rows) containing Brent crude closing prices aligned with mapped geopolitical events, categories, and Indian import statistics from January 2022 to May 23, 2026.
*   📈 **[oil_geopolitical_timeline_chart.png](oil_geopolitical_timeline_chart.png)**: An annotated high-resolution line chart highlighting major geopolitical milestones directly on the price line.
*   🖥️ **[Geopolitical_Oil_Research_Presentation.pptx](Geopolitical_Oil_Research_Presentation.pptx)**: A professional, light-themed 12-slide presentation deck (16:9 widescreen format) containing complete bullet points, embedded chart assets, and native presenter speaker notes.
*   🐍 **[oil_research_scraper.py](oil_research_scraper.py)**: Python script utilizing the **Scrapling** static fetcher to gather Brent prices and merge them with our curated geopolitical timeline database.
*   🐍 **[generate_chart.py](generate_chart.py)**: Python script using `matplotlib` to render the styled, annotated line chart.
*   🐍 **[create_presentation.py](create_presentation.py)**: Python script using `python-pptx` to compile the slide deck programmatically.
*   🏫 **[iit_roorkee_logo.png](iit_roorkee_logo.png)**: The official IIT Roorkee logo watermark used in the presentation layout.

---

## Installation & Setup

To run the scraper and recreate the dataset or charts locally, set up the environment:

1. **Install Dependencies:**
   ```bash
   pip install "scrapling[all]" python-pptx matplotlib
   ```

2. **Install Scrapling Browsers:**
   ```bash
   scrapling install
   ```

3. **Run the Code:**
   * Run the scraper:
     ```bash
     python oil_research_scraper.py
     ```
   * Regenerate the chart:
     ```bash
     python generate_chart.py
     ```
   * Compile the PowerPoint file:
     ```bash
     python create_presentation.py
     ```
