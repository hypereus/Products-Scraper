import json
import sqlite3
from playwright.sync_api import sync_playwright
import pandas as pd
from tqdm import tqdm
import hashlib
from datetime import datetime

DB_FILE = 'retail_data.db'
excel_file = 'Westside-Menswear.xlsx'


# ==================== Database Functions ====================

def get_db_connection():
    """Get a database connection with proper timeout and WAL mode"""
    conn = sqlite3.connect(DB_FILE, timeout=30.0, isolation_level='DEFERRED')
    conn.execute('PRAGMA journal_mode=WAL')
    return conn


def initialize_database():
    """Create database table with proper schema or add missing columns"""
    conn = get_db_connection()
    cursor = conn.cursor()

    # cursor.execute('''
    #     DROP TABLE IF EXISTS Apparels
    # ''')
    # Create table if it doesn't exist
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Apparels (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_hash TEXT UNIQUE NOT NULL,
            Brand TEXT,
            Item TEXT,
            Descriptor TEXT,
            Colour TEXT,
            IsNew TEXT,
            Current_Price REAL,
            Previous_Price REAL,
            Link TEXT UNIQUE,
            FirstSeen TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            LastUpdated TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            LastSeen TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            CONSTRAINT Uniq_Product UNIQUE (Brand, Item, Descriptor, Colour)
        )
    ''')

    conn.commit()
    conn.close()

    print("Database initialized - table 'Apparels' created or reset!")


def get_total_products_in_db():
    """Get the total count of products in database"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT COUNT(*) FROM Apparels')
    count = cursor.fetchone()[0]
    conn.close()
    return count


def generate_product_hash(product):
    """Generate unique MD5 hash for product based on key attributes"""
    unique_string = f"{product['Brand']}|{product['Item']}|{product['Descriptor']}|{product['Colour']}"
    return hashlib.md5(unique_string.encode()).hexdigest()


def product_exists(product_hash):
    """Check if product exists in database"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT id FROM Apparels WHERE product_hash = ?', (product_hash,))
    exists = cursor.fetchone() is not None
    conn.close()
    return exists


def get_product_prices(product_hash):
    """Get current prices for existing product"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT Current_Price, Previous_Price FROM Apparels WHERE product_hash = ?', (product_hash,))
    result = cursor.fetchone()
    conn.close()
    return result if result else (None, None)


def update_existing_product(product_hash, product):
    """Update existing product's price and timestamp"""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE Apparels 
        SET Current_Price = ?, 
            Previous_Price = ?,
            IsNew = ?,
            LastUpdated = CURRENT_TIMESTAMP,
            LastSeen = CURRENT_TIMESTAMP
        WHERE product_hash = ?
    ''', (product['Current Price'], product['Previous Price'], product['IsNew'], product_hash))
    conn.commit()
    conn.close()


def save_to_db(data, metadata):
    """Save products to database, adding new or updating changed prices"""
    conn = get_db_connection()
    cursor = conn.cursor()

    products_before = get_total_products_in_db()

    new_products = 0
    updated_products = 0
    unchanged_products = 0
    skipped_duplicates = 0
    processed_hashes = set()

    for product in tqdm(data, desc="Saving to database "):
        product_hash = generate_product_hash(product)

        # Skip duplicates within this scrape
        if product_hash in processed_hashes:
            skipped_duplicates += 1
            continue

        processed_hashes.add(product_hash)

        if product_exists(product_hash):
            old_current_price, old_previous_price = get_product_prices(product_hash)

            if (old_current_price != product['Current Price'] or
                    old_previous_price != product['Previous Price']):
                conn.commit()
                update_existing_product(product_hash, product)
                updated_products += 1
            else:
                cursor.execute('UPDATE Apparels SET LastSeen = CURRENT_TIMESTAMP WHERE product_hash = ?', (product_hash,))
                unchanged_products += 1
        else:
            try:
                cursor.execute('''
                    INSERT INTO Apparels 
                    (product_hash, Brand, Item, Descriptor, Colour, IsNew, 
                     Current_Price, Previous_Price, Link)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (product_hash, product['Brand'], product['Item'],
                      product['Descriptor'], product['Colour'], product['IsNew'],
                      product['Current Price'], product['Previous Price'], product['Link']))
                new_products += 1
            except sqlite3.IntegrityError:
                skipped_duplicates += 1

    conn.commit()
    conn.close()

    products_after = get_total_products_in_db()

    print(f"\n--- Database Save Summary ---")
    print(f"Products before scrape: {products_before}")
    print(f"Products after scrape: {products_after}")
    print(f"New products added: {new_products}")
    print(f"Existing products with price updates: {updated_products}")
    print(f"Existing products unchanged: {unchanged_products}")
    if skipped_duplicates > 0:
        print(f"Duplicate products in scrape (skipped): {skipped_duplicates}")
    print(f"Total products scraped: {len(data)}")

    complete_metadata = {
        'Timestamp': metadata['scrape_start'].strftime('%Y-%m-%d %H:%M:%S'),
        'Scrape Duration (in S)': metadata['scrape_duration_seconds'],
        'Products Before Scrape': products_before,
        'Products After Scrape': products_after,
        'New Products Added': new_products,
        'Previous Products Updated': updated_products,
        'Products Scraped': len(data),
        'Unique Products': products_after - products_before
    }

    # Fetch latest data from DB including timestamps for Excel export
    export_query = """
        SELECT 
            Brand, Item, Descriptor, Colour, IsNew, 
            Current_Price as "Current Price", 
            Previous_Price as "Previous Price", 
            Link,
            FirstSeen, LastUpdated, LastSeen
        FROM Apparels 
        ORDER BY Brand, Item, Descriptor, Colour DESC
    """
    db_data = load_data_into_pandas(export_query)
    save_to_excel(db_data, complete_metadata)

    return {
        'new_products': new_products,
        'updated_products': updated_products,
        'unchanged_products': unchanged_products,
        'skipped_duplicates': skipped_duplicates
    }


def load_data_into_pandas(sql_query="SELECT * FROM Apparels"):
    """Load data from database into pandas DataFrame"""
    try:
        conn = get_db_connection()
        df = pd.read_sql_query(sql_query, conn)
        conn.close()
        #print(f"Successfully loaded {len(df)} records into DataFrame.")
        return df
    except sqlite3.Error as e:
        print(f"Error accessing database: {e}")
        return pd.DataFrame()


# ==================== Excel Export Functions ====================

def save_to_excel(data, metadata):
    """Save products and metadata to Excel with multiple sheets"""
    products_df = pd.DataFrame(data)

    # Load existing metadata if available
    try:
        with pd.ExcelFile(excel_file) as xls:
            existing_metadata_df = pd.read_excel(xls, sheet_name='Scrape Metadata')
        new_metadata_df = pd.DataFrame([metadata])
        metadata_df = pd.concat([existing_metadata_df, new_metadata_df], ignore_index=True)
    except (FileNotFoundError, ValueError):
        metadata_df = pd.DataFrame([metadata])

    with pd.ExcelWriter(excel_file, engine='openpyxl', mode='w') as writer:
        products_df.to_excel(writer, sheet_name='Products', index=False)
        metadata_df.to_excel(writer, sheet_name='Scrape Metadata', index=False)

    print(f"Data exported to Excel with {len(products_df)} products!")
    print(f"Scrape metadata appended - Total scrapes recorded: {len(metadata_df)}")


# ==================== Data Cleaning Functions ====================

def clean_product_cards(product_cards, attribute_link):
    """Extract text data from product card elements"""
    cleaned_cards = []
    product_urls = set()

    for card in tqdm(product_cards, desc="Cleaning Products "):
        product = {}
        raw_text = card.inner_text().split('\n')
        clean_text = [item for item in raw_text if item.strip()]

        if clean_text[0] == 'New':
            product['IsNew'] = 'True'
            product['Brand'] = clean_text[1]
            result = clean_product_title(clean_text[2])

            product['Item'] = result['Item']
            product['Descriptor'] = result['Descriptor']
            product['Colour'] = result['Colour']
            product['Previous Price'], product['Current Price'] = clean_product_price(clean_text[3])
            product['Link'] = card.locator(attribute_link).get_attribute('href')
        else:
            product['IsNew'] = 'False'
            product['Brand'] = clean_text[0]
            result = clean_product_title(clean_text[1])

            product['Item'] = result['Item']
            product['Descriptor'] = result['Descriptor']
            product['Colour'] = result['Colour']
            product['Previous Price'], product['Current Price'] = clean_product_price(clean_text[2])
            product['Link'] = card.locator(attribute_link).get_attribute('href')

        product_urls.add(product['Link'])

        cleaned_cards.append(product)

    return cleaned_cards, product_urls


def clean_product_price(price_string):
    """Parse price string to extract current and previous prices"""
    price = price_string.split()

    def parse_price(price_str):
        """Helper to parse individual price"""
        if ',' in price_str:
            parts = price_str.split('.')
            rupees_parts = parts[0].split(',')
            rupees = int(rupees_parts[0]) * 1000 + int(rupees_parts[1])
            paise = int(parts[-1])
            return rupees + paise * 0.01 + 1
        else:
            parts = price_str.split('.')
            return int(parts[0]) + int(parts[-1]) * 0.01 + 1

    if len(price) == 4:
        previous_price = parse_price(price[-1])
        current_price = parse_price(price[-3])
    else:
        previous_price = parse_price(price[-1])
        current_price = previous_price

    return previous_price, current_price


def clean_product_title(title):
    """Parse product title to extract brand, item, descriptor, and colour"""
    item = {}
    first_colour = ["Dark", "Dusty", "Light", "Plain"]
    words = title.split()

    # Determine item type
    if words[-3] == "Pack" and words[-2] == "of":
        item["Item"] = words[-5]
        item_offset = 5
    elif words[-2] in ["Polo", "Track"]:
        item["Item"] = f"{words[-2]} {words[-1]}"
        item_offset = 2
    else:
        item["Item"] = words[-1]
        item_offset = 1

    # Determine brand
    if words[0] == "WES":
        item["Brand"] = f"{words[0]} {words[1]}"
        start_idx = 2
    else:
        item["Brand"] = words[0]
        start_idx = 1

    # Determine colour and descriptor
    if words[start_idx] in first_colour:
        item["Colour"] = f"{words[start_idx]} {words[start_idx + 1]}"
        item["Descriptor"] = " ".join(words[start_idx + 2:len(words) - item_offset])
    else:
        item["Colour"] = words[start_idx]
        item["Descriptor"] = " ".join(words[start_idx + 1:len(words) - item_offset])

    return item


# ==================== Web Scraping Functions ====================

def run_scraper(site_name):
    """Main scraping function using Playwright"""
    # Load configuration
    with open('websites.json', 'r') as f:
        config = json.load(f)[site_name]

    scrape_start = datetime.now()

    with sync_playwright() as p:
        browser = p.chromium.launch()
        #browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        )
        page = context.new_page()
        page.goto(config['url'], wait_until='load')

        # Scroll to load all products
        for _ in tqdm(range(60), desc="Scrolling page "):
            page.mouse.wheel(0, 15000)
            page.wait_for_timeout(1000)
            page.wait_for_load_state('load')

        page.wait_for_load_state('load')
        product_cards = page.locator(config['container']).all()

        print(f"Products Scraped : {len(product_cards)}")

        data_list, product_urls = clean_product_cards(product_cards, config['selectors']['link'])

        print(f"Number of unique products: {len(product_urls)}")
        browser.close()

        scrape_end = datetime.now()
        scrape_duration = (scrape_end - scrape_start).total_seconds()

        return data_list, {
            'scrape_start': scrape_start,
            'scrape_end': scrape_end,
            'scrape_duration_seconds': round(scrape_duration, 2),
            'products_found': len(product_cards),
            'unique_urls': len(product_urls)
        }


# ==================== Main Execution ====================

if __name__ == "__main__":

    initialize_database()
    scraped_data, scrape_metadata = run_scraper('Westside')
    save_stats = save_to_db(scraped_data, scrape_metadata)

    query = """
        SELECT 
            Brand, Item, Descriptor, Colour, IsNew, 
            Current_Price, Previous_Price, Link,
            FirstSeen, LastUpdated, LastSeen
        FROM 
            Apparels 
        ORDER BY Brand, Item, Descriptor DESC
    """

    westside_df = load_data_into_pandas(query)
    print("\n--- Westside Products in Database ---")
    print(westside_df)
