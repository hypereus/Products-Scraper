# Westside Menswear Price Tracker

A Python-based web scraper that monitors product listings and price changes for Westside's menswear collection. The scraper automatically tracks new products, price updates, and maintains historical data in both SQLite and Excel formats.

## Features

- **Automated Web Scraping**: Uses Playwright to scrape product data from Westside's menswear section
- **Price Change Tracking**: Monitors both current and previous prices, detecting when products go on sale
- **Duplicate Detection**: Uses MD5 hashing to prevent duplicate entries
- **SQLite Database**: Stores product data with timestamps for first seen, last updated, and last seen
- **Excel Export**: Generates detailed Excel reports with product data and scrape metadata
- **Smart Updates**: Only updates database records when prices actually change
- **Progress Tracking**: Real-time progress bars for scraping and data processing

## Project Structure

```
.
├── main.py                    # Main scraper script
├── websites.json              # Configuration file for scraping targets
├── retail_data.db            # SQLite database (auto-generated)
└── Westside-Menswear.xlsx    # Excel export (auto-generated)
```

## Requirements

- Python 3.7+
- playwright
- pandas
- tqdm
- openpyxl

## Installation

1. Clone this repository:
```bash
git clone <repository-url>
cd westside-scraper
```

2. Install required packages:
```bash
pip install playwright pandas tqdm openpyxl
```

3. Install Playwright browsers:
```bash
playwright install chromium
```

4. Create a `websites.json` configuration file:
```json
{
  "Westside": {
    "url": "https://www.westside.com/...",
    "container": "css-selector-for-product-cards",
    "selectors": {
      "link": "a"
    }
  }
}
```

## Usage

Run the scraper:
```bash
python main.py
```

The script will:
1. Initialize the SQLite database (if not exists)
2. Scrape product data from Westside
3. Process and clean the data
4. Update the database with new/changed products
5. Export results to Excel

## Database Schema

The `Apparels` table includes:

| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER | Primary key |
| product_hash | TEXT | Unique MD5 hash of product attributes |
| Brand | TEXT | Product brand name |
| Item | TEXT | Product type (e.g., Shirt, Trouser) |
| Descriptor | TEXT | Product description |
| Colour | TEXT | Product color |
| IsNew | TEXT | Whether product is marked as "New" |
| Current_Price | REAL | Current selling price |
| Previous_Price | REAL | Previous/original price |
| Link | TEXT | Product URL |
| FirstSeen | TIMESTAMP | When product was first discovered |
| LastUpdated | TIMESTAMP | When product data was last modified |
| LastSeen | TIMESTAMP | When product was last seen in a scrape |

## Data Processing

The scraper handles:

- **Price Parsing**: Extracts current and previous prices from various formats
- **Product Title Parsing**: Separates brand, item type, descriptor, and color
- **New Product Detection**: Identifies items marked as "New"
- **Multi-pack Items**: Correctly handles "Pack of X" products
- **Color Variations**: Parses compound colors (e.g., "Dark Blue", "Light Grey")

## Excel Output

The Excel file contains two sheets:

1. **Products**: Complete product listing with all attributes and timestamps
2. **Scrape Metadata**: Historical log of all scrape operations including:
   - Timestamp
   - Scrape duration
   - Products found
   - New products added
   - Updated products
   - Total database size

## Scraping Strategy

The scraper uses a scroll-based approach to load dynamic content:
- Performs 60 scroll iterations
- Waits 1 second between scrolls for content to load
- Ensures all products are loaded before extraction

## Output Summary

After each run, you'll see:
```
--- Database Save Summary ---
Products before scrape: X
Products after scrape: Y
New products added: Z
Existing products with price updates: A
Existing products unchanged: B
Total products scraped: C
```

## Notes

- The scraper uses WAL mode for better SQLite performance
- Duplicate detection happens at both the scrape level and database level
- Products not seen in recent scrapes are still retained with their LastSeen timestamp
- The script is configured for headless operation but can be run with UI for debugging

## License

MIT License - feel free to modify and use for your projects.

## Disclaimer

This tool is for educational purposes. Always respect website terms of service and robots.txt when scraping. Consider implementing rate limiting and respectful scraping practices.
