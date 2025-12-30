import json
import sqlite3
from playwright.sync_api import sync_playwright
import pandas as pd
from tqdm import tqdm

DB_FILE = 'retail_data.db'  # Ensure this matches the file name you used in the scraper.py

def load_data_into_pandas(query="SELECT * FROM Apparels"):
    """
    Connects to the SQLite DB, executes a query, and loads the results into a Pandas DataFrame.

    :param query: The SQL query to execute. Defaults to selecting all columns from the 'products' table.
    :return: A Pandas DataFrame containing the query results.
    """
    try:
        # 1. Connect to the database file
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        # Method 1: Drop table if exists (recommended)
        cursor.execute("DROP TABLE IF EXISTS products")
        # 2. Use read_sql_query to execute the SQL and load results directly into a DataFrame
        df = pd.read_sql_query(query, conn)

        # 3. Close the connection
        conn.close()

        print(f"Successfully loaded {len(df)} records into a DataFrame.")
        return df

    except sqlite3.Error as e:
        print(f"An error occurred while accessing the database: {e}")
        return pd.DataFrame()  # Return an empty DataFrame on error

# Save to Database
def save_to_db(data):
    df = pd.DataFrame(data)
    # Simple SQLite database
    conn = sqlite3.connect('retail_data.db')
    df.to_sql('Apparels', conn, if_exists='replace', index=False) #
    df.to_excel('Westside-Menswear.xlsx', index=False)
    conn.close()
    print("Data saved to database!")

def clean_product_cards(product_cards):
    done_product_cards = []
    for i in tqdm(range(len(product_cards))):
        done_product = {}
        product = product_cards[i].inner_text().split('\n')
        clean_product = [ item for item in product if item.strip() ]
        #print(f"{len(clean_product)}")
        if clean_product[0] == 'New':
            done_product['IsNew'] = 'True'
            done_product['Brand'] = clean_product[1]
            done_product['Title'] = clean_product[2]
            done_product['Price'] = clean_product[3]
        else:
            done_product['IsNew'] = 'False'
            done_product['Brand'] = clean_product[0]
            done_product['Title'] = clean_product[1]
            done_product['Price'] = clean_product[2]
        #print(f'{done_product}')
        done_product_cards.append(done_product)
    return done_product_cards


def clean_product_price(clean_product_price):
    price = clean_product_price.split()
    if len(price) == 4:
        if ',' in price[-1]:
            real_price = 1 + int(price[-1].split('.')[-1]) * 0.01 + int(price[-1].split('.')[0].split(',')[0]) * 1000 + int(
                price[-1].split('.')[0].split(',')[1])
        else:
            real_price = 1 + int(price[-1].split('.')[-1]) * 0.01 + int(price[-1].split('.')[0])

        if ',' in price[-3]:
            current_price = 1 + int(price[-3].split('.')[-1]) * 0.01 + int(price[-3].split('.')[0].split(',')[0]) * 1000 + int(
                price[-3].split('.')[0].split(',')[1])
        else:
            current_price = 1 + int(price[-3].split('.')[-1]) * 0.01 + int(price[-3].split('.')[0])
    else:
        if ',' in price[-1]:
            real_price = 1 + int(price[-1].split('.')[-1]) * 0.01 + int(
                price[-1].split('.')[0].split(',')[0]) * 1000 + int(
                price[-1].split('.')[0].split(',')[1])
        else:
            real_price = 1 + int(price[-1].split('.')[-1]) * 0.01 + int(price[-1].split('.')[0])
        current_price = real_price
    return real_price, current_price

def clean_product_link(product_card_link):
    #print(f"Found {len(product_card_link)} links, Now cleaning Links!")
    print(f"Cleaning Product Card {product_card_link}!")
    link = product_card_link.get_attribute('href')
    print(f"{link}")
    return link


def clean_link(links):
    print(f"Found {len(links)} links, Now cleaning Links!")
    link_list = []
    for i in tqdm(range(len(links))):
        #print(f"Cleaning Link {i}!")
        temp_link = links[i]
        link = temp_link.get_attribute('href')
        #print(f"{link}")
        link_list.append(link)
    return link_list

def clean_price(prices):
    print(f"Found {len(prices)} prices, Now cleaning Prices!")
    price_list = []
    for i in tqdm(range(len(prices))):
        #print(f"Cleaning Price {i}!")
        #print(f"{prices[i].split()}")
        price = prices[i].split()[-1]
        if ',' in price:
            real_price = 1 + int(price.split('.')[-1])*0.01 + int(price.split('.')[0].split(',')[0])*1000 + int(price.split('.')[0].split(',')[1])
        else:
            real_price = 1 + int(price.split('.')[-1])*0.01 + int(price.split('.')[0])
        price_list.append(real_price)
    return price_list

def clean_product_title(product_card_title):
    item = {}
    new_title = product_card_title.split()
    if new_title[-2] == "of" and new_title[-3] == "Pack": #Underwear Condition
        item["Item"] = new_title[-5]
        if new_title[0] == "WES": #WES Condition
            item["Brand"] = new_title[0] + " " + new_title[1]
            if new_title[2] in ["Dark","Dusty","Light", "Plain"] : #Light or Dark Colour Condition
                item["Colour"] = new_title[2] + " " + new_title[3]
                item["Descriptor"] = new_title[4]
                for j in range(5, len(new_title)):
                    item["Descriptor"] += " " + new_title[j]
            else:
                item["Colour"] = new_title[2]
                item["Descriptor"] = new_title[3]
                for j in range(4, len(new_title)):
                    item["Descriptor"] += " " + new_title[j]
        else:
            item["Brand"] = new_title[0]
            if new_title[1] == "Dark" or new_title[1] == "Light":
                item["Colour"] = new_title[1] + " " + new_title[2]
                item["Descriptor"] = new_title[3]
                for j in range(4, len(new_title)):
                    item["Descriptor"] += " " + new_title[j]

            else:
                item["Colour"] = new_title[1]
                item["Descriptor"] = new_title[2]
                for j in range(3, len(new_title)):
                    item["Descriptor"] += " " + new_title[j]
    else:
        item["Item"] = new_title[-1]
        if new_title[0] == "WES":
            item["Brand"] = new_title[0] + " " + new_title[1]
            if new_title[2] == "Dark" or new_title[2] == "Light":
                item["Colour"] = new_title[2] + " " + new_title[3]
                item["Descriptor"] = new_title[4]
                for j in range(5, len(new_title) - 1):
                    item["Descriptor"] += " " + new_title[j]
            else:
                item["Colour"] = new_title[2]
                item["Descriptor"] = new_title[3]
                for j in range(4, len(new_title) - 1):
                    item["Descriptor"] += " " + new_title[j]
        else:
            item["Brand"] = new_title[0]
            if new_title[1] == "Dark" or new_title[1] == "Light":
                item["Colour"] = new_title[1] + " " + new_title[2]
                item["Descriptor"] = new_title[3]
                for j in range(4, len(new_title) - 1):
                    item["Descriptor"] += " " + new_title[j]

            else:
                item["Colour"] = new_title[1]
                item["Descriptor"] = new_title[2]
                for j in range(3, len(new_title) - 1):
                    item["Descriptor"] += " " + new_title[j]
        #print(f"Item {i} : {item}\n")
    return item #List of Dictionaries

def clean_title(titles):
    print(f"Found {len(titles)} titles, Now cleaning Titles!")
    new_titles = []
    items = []
    item = {}
    for i in tqdm(range(len(titles))):
        #print(f"Cleaning Title {i}!")
        new_titles.append(titles[i].split())
        if new_titles[i][-2] == "of" and new_titles[i][-3] == "Pack": #Underwear Condition
            item["Item"] = new_titles[i][-5]
            if new_titles[i][0] == "WES": #WES Condition
                item["Brand"] = new_titles[i][0] + " " + new_titles[i][1]
                if new_titles[i][2] == "Dark" or new_titles[i][2] == "Light": #Light or Dark Colour Condition
                    item["Colour"] = new_titles[i][2] + " " + new_titles[i][3]
                    item["Descriptor"] = new_titles[i][4]
                    for j in range(5, len(new_titles[i])):
                        item["Descriptor"] += " " + new_titles[i][j]
                else:
                    item["Colour"] = new_titles[i][2]
                    item["Descriptor"] = new_titles[i][3]
                    for j in range(4, len(new_titles[i])):
                        item["Descriptor"] += " " + new_titles[i][j]
            else:
                item["Brand"] = new_titles[i][0]
                if new_titles[i][1] == "Dark" or new_titles[i][1] == "Light":
                    item["Colour"] = new_titles[i][1] + " " + new_titles[i][2]
                    item["Descriptor"] = new_titles[i][3]
                    for j in range(4, len(new_titles[i])):
                        item["Descriptor"] += " " + new_titles[i][j]

                else:
                    item["Colour"] = new_titles[i][1]
                    item["Descriptor"] = new_titles[i][2]
                    for j in range(3, len(new_titles[i])):
                        item["Descriptor"] += " " + new_titles[i][j]
        else:
            item["Item"] = new_titles[i][-1]
            if new_titles[i][0] == "WES":
                item["Brand"] = new_titles[i][0] + " " + new_titles[i][1]
                if new_titles[i][2] == "Dark" or new_titles[i][2] == "Light":
                    item["Colour"] = new_titles[i][2] + " " + new_titles[i][3]
                    item["Descriptor"] = new_titles[i][4]
                    for j in range(5, len(new_titles[i]) - 1):
                        item["Descriptor"] += " " + new_titles[i][j]
                else:
                    item["Colour"] = new_titles[i][2]
                    item["Descriptor"] = new_titles[i][3]
                    for j in range(4, len(new_titles[i]) - 1):
                        item["Descriptor"] += " " + new_titles[i][j]
            else:
                item["Brand"] = new_titles[i][0]
                if new_titles[i][1] == "Dark" or new_titles[i][1] == "Light":
                    item["Colour"] = new_titles[i][1] + " " + new_titles[i][2]
                    item["Descriptor"] = new_titles[i][3]
                    for j in range(4, len(new_titles[i]) - 1):
                        item["Descriptor"] += " " + new_titles[i][j]

                else:
                    item["Colour"] = new_titles[i][1]
                    item["Descriptor"] = new_titles[i][2]
                    for j in range(3, len(new_titles[i]) - 1):
                        item["Descriptor"] += " " + new_titles[i][j]
            #print(f"Item {i} : {item}\n")
        items.append(item)
        item = {}
    return items #List of Dictionaries

def run_scraper(site_name):
    # 1. Load Config
    with open('websites.json', 'r') as f:
        config = json.load(f)[site_name]

    data_list = []

    with sync_playwright() as p:
        # 2. Launch Browser (Headless=False lets you watch it work!)

        #browser = p.chromium.launch()
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        )
        page = browser.new_page()
        page.goto(config['url'], wait_until='domcontentloaded')

        # 3. Handle Dynamic Content (Scroll to bottom)
        # Retail sites often load items only when you scroll
        for _ in tqdm(range(60)):  # Scroll 60 times (make this configurable)
            page.mouse.wheel(0, 15000)
            page.wait_for_timeout(1000)  # Wait for content to load

        product_cards = page.locator(config['container']).all()
        print(f"Found {len(product_cards)} products")
        #product_titles = clean_title(page.locator(config['selectors']['title']).all_inner_texts())
        #print(f"{product_titles}\n")

        #product_prices = clean_price(page.locator(config['selectors']['price']).all_inner_texts())
        #print(f"{product_prices}\n")

        product_links = clean_link(page.locator(config['selectors']['link']).all())
        #print(f"{product_links}")

        #print(f"Found {len(product_cards)} products. Extracting...")
        cpc = clean_product_cards(product_cards)
        #print(f'{cpc}')
        print(f"")
        for i in tqdm(range(len(product_cards))):
            #print(f"Extracting Product #{i}")
            product = {}

            #print(f"{product_cards[i].inner_text().split('\n')}")

            product['Brand'] = cpc[i]['Brand']
            product_title = clean_product_title(cpc[i]['Title'])
            #product['Unclean Brand'] = product_title['Brand']
            product['Colour'] = product_title['Colour']
            product['Descriptor'] = product_title['Descriptor']
            product['Item'] = product_title['Item']
            product['IsNew'] = cpc[i]['IsNew']
            product_price = clean_product_price(cpc[i]['Price'])
            product['Current Price'] = product_price[1]
            product['Previous Price'] = product_price[0]
            product['Unclean Link'] = clean_product_link(product_cards[i].locator("a.wizzy-result-product-item"))
            print(f"{product['Unclean Link']}")
            product['Link'] = product_links[i]
            data_list.append(product)

        browser.close()
        return data_list


# Run it
scraped_data = run_scraper('Westside')
save_to_db(scraped_data)

# --- Example Usage ---

# Load all data from the 'products' table
#all_products_df = load_data_into_pandas()
#print("\n--- First 5 rows of all data ---")
#print(all_products_df.head())

# --- Example of a Specific Query ---
specific_query = """
    SELECT 
        *
    FROM 
        Apparels 
"""

westside_deals_df = load_data_into_pandas(specific_query)
print("\n--- Westside Products ---")
print(westside_deals_df)