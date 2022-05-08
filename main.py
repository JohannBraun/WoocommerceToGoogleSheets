from woocommerce import API
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
import gspread
import os
from pprint import pprint
import gspread_formatting

# ---------- GOOGLE SHEETS API SETUP ----------- <
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'keys.json'

credentials = None
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# The ID of our spreadsheet. Get from googlesheets
SAMPLE_SPREADSHEET_ID = os.environ.get("SAMPLE_SPREADSHEET_ID")
# Create gspread service account with credentials from keys
gc = gspread.service_account(filename='keys.json')

# max row in your google sheet. if you have more products, you need more rows
MAX_ROWS = 500
# name of the newly created sheet in which the products will be seen
ROOT_SHEET = "products"
# choose your own titles for the corresponding data that you pull.
HEADER_ROW = ["id",
              "Art.-Nr.",
              "# Lager",
              "Verfügbarkeit",
              "Name",
              "Preis",
              "UVP",
              "Permalink",
              "Beschreibung",
              "Mini-Beschr.",
              "Kurz-Beschr.",
              "Bild-src"]
try:
    # Create googlesheet
    gsheet = gc.open_by_key(SAMPLE_SPREADSHEET_ID)
except HttpError as err:
    print(err)

# ---------- WOOCOMMERCE SETUP ----------- <
wcapi_consumer_key = os.environ.get("wcapi_consumer_key")
wcapi_consumer_secret = os.environ.get("wcapi_consumer_secret")

wcapi = API(
    url="https://www.ecology-water.eu",
    consumer_key=wcapi_consumer_key,
    consumer_secret=wcapi_consumer_secret,
    wp_api=True,
    version="wc/v3",
    query_string_auth=True,
    timeout=50
)


# ---------- CREATE SHEET ----------- <
def create_new_sheet():
    new_sheet = gsheet.add_worksheet(title=ROOT_SHEET, rows=MAX_ROWS, cols="20")
    new_sheet.append_row(HEADER_ROW)
    return new_sheet


# ---------- FORMAT HEADER ----------- <
def format_header(sheet):
    sheet.format("A1:K1", {
        "wrapStrategy": "WRAP",
        "textFormat": {
            "fontSize": 12,
            "bold": True
        }
    })

    gspread_formatting.set_frozen(sheet, rows=1)
    gspread_formatting.set_row_height(sheet, f'1:{MAX_ROWS}', 42)
    gspread_formatting.set_column_widths(sheet, [("E", 600), ("H:J", 200), ("A", 60), ("C", 60)])


# ---------- Get products from woocommerce, push into googlesheets ----------- <
def get_from_woocommerce():
    cur_sheet.batch_clear([f'A2:J{MAX_ROWS}'])
    product_data = {
        "per_page": 50,
        "page": 1
    }
    # the while loop runs as long as there are still products to pull
    while True:
        response = wcapi.get('products', params=product_data)
        products = response.json()
        products_formatted = []
        for product in products:
            try:
                product_img_src = product["images"][0]["src"]
            except IndexError:
                product_img_src = "invalid picture src"

            # we dont want every single information out of woocommerce, in the following we filter for the
            # data that we want. remember to name your row headers accordingly
            # you can print(product) to see which variables woocommerce gives you
            product_category_values = [
                product["id"],
                product["sku"],
                product["stock_quantity"],
                product["stock_status"],
                product["name"],
                product["price"],
                product["regular_price"],
                product["permalink"],
                product["description"],
                product["mini_desc"],
                product["short_description"],
                product_img_src
            ]
            products_formatted.append(product_category_values)

        cur_sheet.append_rows(values=products_formatted)
        product_data["page"] += 1
        print(response.status_code)

        # After searching through all products, next page is empty, i.e the lenght is 0, so we can stop searching
        if len(products) == 0:
            format_products()
            break


# Format body
def format_products():
    fmt = gspread_formatting.CellFormat(
        horizontalAlignment='CENTER',
        wrapStrategy='WRAP'
    )
    gspread_formatting.format_cell_range(cur_sheet, f'G2:J{MAX_ROWS}', fmt)


# function to push the data that we have in our sheet back to woocommerce
# like that we can make changes in gsheets and see them directly on the page
def push_to_woocommerce():
    products_in_sheet = cur_sheet.get_all_records()
    for product in products_in_sheet:
        data = {
            "regular_price": str(product["UVP"]),
            "stock_quantity": product["Auf Lager"],
            "description": product["Beschreibung"],
            "short_description": product["Kurz-Beschr."],
            "mini_desc": product["Mini-Beschr."],
            "name": product["Name"],
            "price": product["Preis"],
        }
        print(wcapi.put(f"products/{product['id']}", data).json())


try:
    cur_sheet = create_new_sheet()
except gspread.exceptions.APIError as err:
    cur_sheet = gsheet.worksheet(ROOT_SHEET)

format_header(cur_sheet)
get_from_woocommerce()
# pprint(products_in_sheet)
# push_to_woocommerce()
