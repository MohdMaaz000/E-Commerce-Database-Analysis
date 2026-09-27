"""
Sales Performance Dashboard - Complete Project Generator
=========================================================
Generates:
  1. 12 monthly sales CSV files with 10,000+ transactions
  2. customers.csv and products.csv lookup files
  3. Excel workbook with Dashboard, Pivot Analysis, KPI Calculations,
     XLOOKUP enrichment, charts, and professional formatting

Data includes realistic quality issues (nulls, duplicates, inconsistent
casing, date format variations) so the Power Query cleaning step is
meaningful.
"""

import csv
import os
import random
import math
from datetime import datetime, timedelta
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.chart import (
    BarChart, LineChart, PieChart, Reference, BarChart3D
)
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
EXCEL_DIR = os.path.join(BASE_DIR, "excel")
DOCS_DIR = os.path.join(BASE_DIR, "docs")

YEAR = 2025
SEED = 42
random.seed(SEED)

# ---------------------------------------------------------------------------
# REFERENCE DATA
# ---------------------------------------------------------------------------

REGIONS = ["North", "South", "East", "West"]

STATES_BY_REGION = {
    "North": ["Delhi", "Punjab", "Haryana", "Uttar Pradesh", "Rajasthan"],
    "South": ["Karnataka", "Tamil Nadu", "Kerala", "Andhra Pradesh", "Telangana"],
    "East": ["West Bengal", "Bihar", "Odisha", "Jharkhand", "Assam"],
    "West": ["Maharashtra", "Gujarat", "Goa", "Madhya Pradesh", "Chhattisgarh"],
}

CITIES_BY_STATE = {
    "Delhi": ["New Delhi", "Dwarka", "Rohini"],
    "Punjab": ["Chandigarh", "Amritsar", "Ludhiana"],
    "Haryana": ["Gurgaon", "Faridabad", "Panipat"],
    "Uttar Pradesh": ["Lucknow", "Noida", "Agra"],
    "Rajasthan": ["Jaipur", "Udaipur", "Jodhpur"],
    "Karnataka": ["Bangalore", "Mysore", "Hubli"],
    "Tamil Nadu": ["Chennai", "Coimbatore", "Madurai"],
    "Kerala": ["Kochi", "Trivandrum", "Calicut"],
    "Andhra Pradesh": ["Hyderabad", "Visakhapatnam", "Vijayawada"],
    "Telangana": ["Hyderabad", "Warangal", "Nizamabad"],
    "West Bengal": ["Kolkata", "Howrah", "Durgapur"],
    "Bihar": ["Patna", "Gaya", "Muzaffarpur"],
    "Odisha": ["Bhubaneswar", "Cuttack", "Rourkela"],
    "Jharkhand": ["Ranchi", "Jamshedpur", "Dhanbad"],
    "Assam": ["Guwahati", "Silchar", "Dibrugarh"],
    "Maharashtra": ["Mumbai", "Pune", "Nagpur"],
    "Gujarat": ["Ahmedabad", "Surat", "Vadodara"],
    "Goa": ["Panaji", "Margao", "Vasco"],
    "Madhya Pradesh": ["Bhopal", "Indore", "Jabalpur"],
    "Chhattisgarh": ["Raipur", "Bhilai", "Bilaspur"],
}

SEGMENTS = ["Consumer", "Corporate", "Home Office", "Small Business"]

CATEGORIES = {
    "Electronics": {
        "Smartphones": [
            ("Galaxy Pro X1", "TechVista", 18000, 28000),
            ("iPhone Lite Z", "FruitTech", 25000, 42000),
            ("Pixel Nova 7", "QuantumMobile", 15000, 24000),
            ("OnePlus Ace R", "SwiftComm", 12000, 19000),
        ],
        "Laptops": [
            ("UltraBook Pro 15", "ByteForce", 35000, 62000),
            ("ThinkPad Edge X", "NovaSys", 28000, 48000),
            ("MacAir Clone Z", "FruitTech", 40000, 72000),
            ("ChromeBook Flex", "QuantumMobile", 15000, 25000),
        ],
        "Tablets": [
            ("Tab Ultra S9", "TechVista", 12000, 22000),
            ("iPad Mini Clone", "FruitTech", 18000, 32000),
            ("SlimPad Pro", "ByteForce", 8000, 15000),
        ],
        "Accessories": [
            ("Wireless Mouse MX", "PeripheralCo", 500, 1800),
            ("Mech Keyboard K70", "PeripheralCo", 2000, 5500),
            ("USB-C Hub 7in1", "SwiftComm", 800, 2200),
            ("Webcam HD Pro", "NovaSys", 1500, 3500),
            ("Bluetooth Speaker Z", "SoundWave", 1200, 3000),
        ],
    },
    "Clothing": {
        "Men's Wear": [
            ("Classic Oxford Shirt", "UrbanThreads", 400, 1500),
            ("Slim Fit Chinos", "DenimCraft", 600, 2000),
            ("Wool Blend Blazer", "EliteWear", 1500, 4500),
            ("Cotton Polo Tee", "UrbanThreads", 300, 900),
        ],
        "Women's Wear": [
            ("Floral Maxi Dress", "StyleHaven", 500, 2200),
            ("Denim Jacket Classic", "DenimCraft", 800, 2800),
            ("Silk Blouse Premium", "EliteWear", 700, 2500),
            ("Yoga Pants Flex", "ActiveZone", 400, 1200),
        ],
        "Footwear": [
            ("Running Shoe AirMax", "StridePro", 1500, 4000),
            ("Leather Loafer Elite", "EliteWear", 1200, 3500),
            ("Canvas Sneaker V2", "UrbanThreads", 600, 1800),
        ],
    },
    "Home & Kitchen": {
        "Appliances": [
            ("Air Fryer ProChef", "HomeEssentials", 2000, 5500),
            ("Blender TurboMix", "KitchenPro", 800, 2200),
            ("Coffee Maker Brew+", "HomeEssentials", 1500, 4000),
            ("Toaster Oven Deluxe", "KitchenPro", 1200, 3200),
        ],
        "Furniture": [
            ("Ergonomic Office Chair", "ComfortLiving", 4000, 12000),
            ("Standing Desk Pro", "ComfortLiving", 8000, 18000),
            ("Bookshelf Walnut 5T", "WoodCraft", 2500, 6500),
        ],
        "Decor": [
            ("LED Desk Lamp Smart", "BrightHome", 600, 1800),
            ("Wall Art Canvas Set", "ArtisanDecor", 800, 2500),
            ("Scented Candle Trio", "ArtisanDecor", 200, 700),
        ],
    },
    "Sports & Outdoors": {
        "Fitness": [
            ("Yoga Mat Premium", "ActiveZone", 300, 900),
            ("Dumbbell Set 20kg", "IronFit", 1500, 3500),
            ("Resistance Band Kit", "ActiveZone", 200, 600),
            ("Fitness Tracker V3", "TechVista", 1800, 4500),
        ],
        "Outdoor Gear": [
            ("Camping Tent 4P", "WildTrail", 3000, 8000),
            ("Hiking Backpack 50L", "WildTrail", 1500, 4000),
            ("Trekking Poles Carbon", "WildTrail", 800, 2200),
        ],
    },
    "Books & Stationery": {
        "Books": [
            ("Python Programming 101", "EduPress", 150, 600),
            ("Data Science Handbook", "EduPress", 200, 800),
            ("Business Strategy Guide", "LeaderBooks", 180, 650),
            ("Fiction Bestseller Pack", "NovelHouse", 250, 900),
        ],
        "Stationery": [
            ("Premium Notebook A5", "WriteWell", 80, 250),
            ("Pen Set Executive", "WriteWell", 200, 600),
            ("Desk Organizer Pro", "OfficeMate", 300, 850),
        ],
    },
}

SALES_CHANNELS = ["Online", "Retail Store", "Marketplace", "Wholesale"]
PAYMENT_METHODS = ["Credit Card", "Debit Card", "UPI", "Net Banking", "COD", "Wallet"]
ORDER_STATUSES = ["Delivered", "Delivered", "Delivered", "Delivered",
                  "Shipped", "Shipped", "Processing", "Cancelled", "Returned"]

# Intentional dirty region variants for data quality issues
DIRTY_REGIONS = {
    "North": ["North", "north", "NORTH", "North ", " North"],
    "South": ["South", "south", "SOUTH", "South ", " South"],
    "East": ["East", "east", "EAST", "East ", " East"],
    "West": ["West", "west", "WEST", "West ", " West"],
}

DIRTY_CATEGORIES = {
    "Electronics": ["Electronics", "electronics", "ELECTRONICS", "Electrnics"],
    "Clothing": ["Clothing", "clothing", "CLOTHING", "Clothng"],
    "Home & Kitchen": ["Home & Kitchen", "home & kitchen", "Home and Kitchen",
                       "Home & kitchen"],
    "Sports & Outdoors": ["Sports & Outdoors", "sports & outdoors",
                          "Sports and Outdoors", "Sports & outdoors"],
    "Books & Stationery": ["Books & Stationery", "books & stationery",
                           "Books and Stationery", "Books & stationery"],
}

DATE_FORMATS = [
    "%Y-%m-%d",
    "%d-%m-%Y",
    "%m/%d/%Y",
    "%d/%m/%Y",
    "%Y/%m/%d",
]


# ---------------------------------------------------------------------------
# HELPER FUNCTIONS
# ---------------------------------------------------------------------------

def build_products():
    """Build flat product list with IDs."""
    products = []
    pid = 1001
    for cat, subcats in CATEGORIES.items():
        for subcat, items in subcats.items():
            for (name, brand, cost_low, price_high) in items:
                std_cost = round(random.uniform(cost_low, cost_low * 1.3), 2)
                products.append({
                    "Product_ID": f"P-{pid:04d}",
                    "Product_Name": name,
                    "Category": cat,
                    "Subcategory": subcat,
                    "Brand": brand,
                    "Standard_Cost": std_cost,
                    "_price_range": (cost_low, price_high),
                })
                pid += 1
    return products


def build_customers(n=500):
    """Build customer list with IDs."""
    first_names = [
        "Aarav", "Vivaan", "Aditya", "Vihaan", "Arjun", "Reyansh", "Sai",
        "Arnav", "Dhruv", "Kabir", "Ananya", "Diya", "Myra", "Sara", "Ira",
        "Aanya", "Navya", "Ishita", "Priya", "Riya", "Rohit", "Amit",
        "Sneha", "Pooja", "Neha", "Raj", "Vikram", "Manoj", "Kiran", "Deepa",
        "Suresh", "Ramesh", "Meena", "Kavita", "Sunita", "Ajay", "Vijay",
        "Sanjay", "Rekha", "Geeta", "Mohan", "Sohan", "Rahul", "Nitin",
        "Preeti", "Swati", "Ankita", "Manish", "Rajesh", "Nisha",
    ]
    last_names = [
        "Sharma", "Verma", "Gupta", "Singh", "Kumar", "Patel", "Shah",
        "Joshi", "Reddy", "Nair", "Iyer", "Pillai", "Das", "Bose",
        "Chatterjee", "Mukherjee", "Banerjee", "Deshmukh", "Patil", "Kulkarni",
        "Mehta", "Jain", "Agarwal", "Mishra", "Pandey", "Dubey", "Tiwari",
        "Saxena", "Kapoor", "Malhotra",
    ]

    customers = []
    used_names = set()
    for i in range(1, n + 1):
        while True:
            name = f"{random.choice(first_names)} {random.choice(last_names)}"
            if name not in used_names:
                used_names.add(name)
                break

        region = random.choice(REGIONS)
        state = random.choice(STATES_BY_REGION[region])
        city = random.choice(CITIES_BY_STATE[state])
        segment = random.choice(SEGMENTS)

        customers.append({
            "Customer_ID": f"C-{i:04d}",
            "Customer_Name": name,
            "Customer_Segment": segment,
            "City": city,
            "State": state,
            "Region": region,
        })
    return customers


def generate_transactions(products, customers, year=2025, target_rows=10500):
    """Generate monthly sales transactions with realistic data quality issues."""
    monthly_data = defaultdict(list)
    order_id = 100001

    # Seasonal weights (higher in Q4, lower in Q1)
    month_weights = {
        1: 0.7, 2: 0.65, 3: 0.75, 4: 0.8, 5: 0.85, 6: 0.9,
        7: 0.95, 8: 1.0, 9: 1.05, 10: 1.1, 11: 1.3, 12: 1.5,
    }
    total_weight = sum(month_weights.values())

    for month in range(1, 13):
        month_rows = int(target_rows * month_weights[month] / total_weight)
        # Ensure at least 500 rows per month
        month_rows = max(month_rows, 500)

        for _ in range(month_rows):
            product = random.choice(products)
            customer = random.choice(customers)

            # Generate order date within the month
            if month == 12:
                max_day = 31
            elif month in [4, 6, 9, 11]:
                max_day = 30
            elif month == 2:
                max_day = 28
            else:
                max_day = 31

            day = random.randint(1, max_day)
            order_date = datetime(year, month, day)

            # Introduce date format inconsistencies (5% of rows)
            if random.random() < 0.05:
                date_str = order_date.strftime(random.choice(DATE_FORMATS))
            else:
                date_str = order_date.strftime("%Y-%m-%d")

            region = customer["Region"]
            state = customer["State"]

            # Dirty region names (8% of rows)
            if random.random() < 0.08:
                region_val = random.choice(DIRTY_REGIONS[region])
            else:
                region_val = region

            # Sales channel
            channel = random.choice(SALES_CHANNELS)

            # Quantity
            qty = random.choices(
                [1, 2, 3, 4, 5, 6, 7, 8, 10],
                weights=[30, 25, 15, 10, 8, 5, 3, 2, 2],
                k=1
            )[0]

            # Unit price with some variance
            price_low, price_high = product["_price_range"]
            unit_price = round(random.uniform(price_low * 1.2, price_high), 2)

            # Discount (0-30%, with some blanks)
            if random.random() < 0.06:
                discount = ""  # blank discount - data quality issue
                discount_val = 0
            else:
                discount_val = random.choices(
                    [0, 0.05, 0.10, 0.15, 0.20, 0.25, 0.30],
                    weights=[35, 15, 20, 12, 10, 5, 3],
                    k=1
                )[0]
                discount = discount_val

            # Calculate financials
            sales_amount = round(qty * unit_price * (1 - discount_val), 2)
            cost = round(qty * product["Standard_Cost"], 2)
            profit = round(sales_amount - cost, 2)

            payment = random.choice(PAYMENT_METHODS)
            status = random.choice(ORDER_STATUSES)

            row = {
                "Order_ID": f"ORD-{order_id}",
                "Order_Date": date_str,
                "Customer_ID": customer["Customer_ID"],
                "Product_ID": product["Product_ID"],
                "Region": region_val,
                "State": state,
                "Sales_Channel": channel,
                "Quantity": qty,
                "Unit_Price": unit_price,
                "Discount": discount,
                "Sales_Amount": sales_amount,
                "Cost": cost,
                "Profit": profit,
                "Payment_Method": payment,
                "Order_Status": status,
            }

            # Introduce missing values (2% chance for certain fields)
            if random.random() < 0.02:
                row["Customer_ID"] = ""
            if random.random() < 0.015:
                row["Payment_Method"] = ""

            monthly_data[month].append(row)
            order_id += 1

    # Add ~50 duplicate rows scattered across months
    for _ in range(50):
        month = random.randint(1, 12)
        if monthly_data[month]:
            dup = random.choice(monthly_data[month]).copy()
            monthly_data[month].append(dup)

    return monthly_data


# ---------------------------------------------------------------------------
# CSV WRITERS
# ---------------------------------------------------------------------------

SALES_FIELDS = [
    "Order_ID", "Order_Date", "Customer_ID", "Product_ID", "Region",
    "State", "Sales_Channel", "Quantity", "Unit_Price", "Discount",
    "Sales_Amount", "Cost", "Profit", "Payment_Method", "Order_Status",
]

MONTH_NAMES = {
    1: "jan", 2: "feb", 3: "mar", 4: "apr", 5: "may", 6: "jun",
    7: "jul", 8: "aug", 9: "sep", 10: "oct", 11: "nov", 12: "dec",
}


def write_sales_csvs(monthly_data):
    """Write monthly CSV files."""
    os.makedirs(DATA_DIR, exist_ok=True)
    total = 0
    for month, rows in sorted(monthly_data.items()):
        fname = f"sales_{MONTH_NAMES[month]}.csv"
        path = os.path.join(DATA_DIR, fname)
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=SALES_FIELDS)
            writer.writeheader()
            writer.writerows(rows)
        total += len(rows)
        print(f"  ✓ {fname} — {len(rows):,} rows")
    print(f"  Total transaction rows: {total:,}")
    return total


def write_customers_csv(customers):
    """Write customers lookup CSV."""
    path = os.path.join(DATA_DIR, "customers.csv")
    fields = ["Customer_ID", "Customer_Name", "Customer_Segment",
              "City", "State", "Region"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(customers)
    print(f"  ✓ customers.csv — {len(customers)} rows")


def write_products_csv(products):
    """Write products lookup CSV."""
    path = os.path.join(DATA_DIR, "products.csv")
    fields = ["Product_ID", "Product_Name", "Category", "Subcategory",
              "Brand", "Standard_Cost"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for p in products:
            row = {k: v for k, v in p.items() if k in fields}
            writer.writerow(row)
    print(f"  ✓ products.csv — {len(products)} rows")


# ---------------------------------------------------------------------------
# CLEAN & CONSOLIDATE DATA (simulating Power Query output)
# ---------------------------------------------------------------------------

def clean_region(val):
    """Standardize region names."""
    if not val:
        return "Unknown"
    val = val.strip().title()
    if val in REGIONS:
        return val
    return "Unknown"


def clean_consolidated_data(monthly_data, products, customers):
    """
    Produce the 'clean' consolidated dataset that would come out of
    Power Query.  We do the same transformations in Python so the
    Excel workbook contains correct, analysis-ready data.
    """
    prod_map = {p["Product_ID"]: p for p in products}
    cust_map = {c["Customer_ID"]: c for c in customers}

    seen_ids = set()
    clean_rows = []

    for month in range(1, 13):
        for row in monthly_data[month]:
            oid = row["Order_ID"]
            # Remove duplicates
            if oid in seen_ids:
                continue
            seen_ids.add(oid)

            # Skip rows with blank order ID
            if not oid:
                continue

            # Clean region
            region = clean_region(row["Region"])

            # Parse date
            date_str = row["Order_Date"]
            parsed_date = None
            for fmt in DATE_FORMATS:
                try:
                    parsed_date = datetime.strptime(date_str.strip(), fmt)
                    break
                except (ValueError, AttributeError):
                    continue
            if parsed_date is None:
                continue  # skip unparseable dates

            # Clean discount
            disc = row["Discount"]
            if disc == "" or disc is None:
                disc_val = 0.0
            else:
                disc_val = float(disc)

            # Quantity
            try:
                qty = int(row["Quantity"])
            except (ValueError, TypeError):
                qty = 1

            # Unit Price
            try:
                unit_price = float(row["Unit_Price"])
            except (ValueError, TypeError):
                continue

            # Recalculate financials
            sales_amount = round(qty * unit_price * (1 - disc_val), 2)
            cost = float(row["Cost"]) if row["Cost"] else 0
            profit = round(sales_amount - cost, 2)
            margin = round(profit / sales_amount * 100, 2) if sales_amount > 0 else 0

            # Customer enrichment
            cid = row["Customer_ID"]
            cust = cust_map.get(cid, {})
            cust_name = cust.get("Customer_Name", "Unknown")
            cust_segment = cust.get("Customer_Segment", "Unknown")

            # Product enrichment
            pid = row["Product_ID"]
            prod = prod_map.get(pid, {})
            prod_name = prod.get("Product_Name", "Unknown")
            category = prod.get("Category", "Unknown")
            subcategory = prod.get("Subcategory", "Unknown")
            brand = prod.get("Brand", "Unknown")

            # Payment
            payment = row["Payment_Method"] if row["Payment_Method"] else "Unknown"

            # Status
            status = row["Order_Status"] if row["Order_Status"] else "Unknown"

            state = row["State"] if row["State"] else "Unknown"
            channel = row["Sales_Channel"] if row["Sales_Channel"] else "Unknown"

            clean_rows.append({
                "Order_ID": oid,
                "Order_Date": parsed_date,
                "Month": parsed_date.strftime("%b"),
                "Month_Num": parsed_date.month,
                "Quarter": f"Q{(parsed_date.month - 1) // 3 + 1}",
                "Customer_ID": cid if cid else "Unknown",
                "Customer_Name": cust_name,
                "Customer_Segment": cust_segment,
                "Product_ID": pid,
                "Product_Name": prod_name,
                "Category": category,
                "Subcategory": subcategory,
                "Brand": brand,
                "Region": region,
                "State": state,
                "Sales_Channel": channel,
                "Quantity": qty,
                "Unit_Price": unit_price,
                "Discount": disc_val,
                "Sales_Amount": sales_amount,
                "Cost": cost,
                "Profit": profit,
                "Profit_Margin_%": margin,
                "Payment_Method": payment,
                "Order_Status": status,
            })

    # Sort by date
    clean_rows.sort(key=lambda r: r["Order_Date"])
    return clean_rows


# ---------------------------------------------------------------------------
# EXCEL WORKBOOK BUILDER
# ---------------------------------------------------------------------------

# Color palette — professional dark blue theme
CLR_PRIMARY = "1B2A4A"       # Dark navy
CLR_SECONDARY = "2E86AB"     # Steel blue
CLR_ACCENT = "F18F01"        # Amber
CLR_SUCCESS = "2D6A4F"       # Forest green
CLR_DANGER = "D62828"        # Red
CLR_BG_LIGHT = "F0F4F8"      # Light gray-blue
CLR_BG_DARK = "1B2A4A"       # Dark navy
CLR_WHITE = "FFFFFF"
CLR_TEXT_DARK = "1A1A2E"
CLR_BORDER = "D1D9E6"
CLR_KPI_BG = "E8F4FD"
CLR_HEADER_BG = "1B2A4A"
CLR_ALT_ROW = "F7FAFC"
CLR_CHART_COLORS = ["2E86AB", "F18F01", "2D6A4F", "D62828", "6C63FF",
                     "E76F51", "264653", "E9C46A", "F4A261", "2A9D8F"]

THIN_BORDER = Border(
    left=Side(style="thin", color=CLR_BORDER),
    right=Side(style="thin", color=CLR_BORDER),
    top=Side(style="thin", color=CLR_BORDER),
    bottom=Side(style="thin", color=CLR_BORDER),
)


def style_header_row(ws, row, max_col, fill_color=CLR_HEADER_BG):
    """Apply header styling to a row."""
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = Font(name="Calibri", size=10, bold=True, color=CLR_WHITE)
        cell.fill = PatternFill(start_color=fill_color, end_color=fill_color,
                                fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = THIN_BORDER


def style_data_rows(ws, start_row, end_row, max_col):
    """Apply alternating row styling."""
    for r in range(start_row, end_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name="Calibri", size=9)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="center")
            if (r - start_row) % 2 == 1:
                cell.fill = PatternFill(start_color=CLR_ALT_ROW,
                                        end_color=CLR_ALT_ROW,
                                        fill_type="solid")


def write_kpi_card(ws, row, col, title, value, fmt="₹#,##0",
                   width=3, color=CLR_SECONDARY):
    """Write a KPI card block on the dashboard."""
    # Title cell
    title_cell = ws.cell(row=row, column=col, value=title)
    title_cell.font = Font(name="Calibri", size=9, bold=True, color=CLR_WHITE)
    title_cell.fill = PatternFill(start_color=color, end_color=color,
                                  fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.border = THIN_BORDER
    # Merge title across width
    if width > 1:
        ws.merge_cells(start_row=row, start_column=col,
                       end_row=row, end_column=col + width - 1)

    # Value cell
    val_cell = ws.cell(row=row + 1, column=col, value=value)
    val_cell.font = Font(name="Calibri", size=16, bold=True, color=CLR_TEXT_DARK)
    val_cell.fill = PatternFill(start_color=CLR_KPI_BG, end_color=CLR_KPI_BG,
                                fill_type="solid")
    val_cell.alignment = Alignment(horizontal="center", vertical="center")
    val_cell.number_format = fmt
    val_cell.border = THIN_BORDER
    if width > 1:
        ws.merge_cells(start_row=row + 1, start_column=col,
                       end_row=row + 1, end_column=col + width - 1)


def create_workbook(clean_data, products, customers, monthly_data):
    """Build the full Excel workbook."""
    wb = Workbook()

    # -----------------------------------------------------------------------
    # SHEET 1: Dashboard  (created last in terms of content, but first tab)
    # -----------------------------------------------------------------------
    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    ws_dash.sheet_properties.tabColor = CLR_PRIMARY

    # -----------------------------------------------------------------------
    # SHEET 2: Power_Query_Output  (clean consolidated data)
    # -----------------------------------------------------------------------
    ws_pq = wb.create_sheet("Power_Query_Output")
    ws_pq.sheet_properties.tabColor = CLR_SECONDARY

    pq_headers = [
        "Order_ID", "Order_Date", "Month", "Month_Num", "Quarter",
        "Customer_ID", "Customer_Name", "Customer_Segment",
        "Product_ID", "Product_Name", "Category", "Subcategory", "Brand",
        "Region", "State", "Sales_Channel",
        "Quantity", "Unit_Price", "Discount",
        "Sales_Amount", "Cost", "Profit", "Profit_Margin_%",
        "Payment_Method", "Order_Status",
    ]

    for c, h in enumerate(pq_headers, 1):
        ws_pq.cell(row=1, column=c, value=h)

    for r, row_data in enumerate(clean_data, 2):
        for c, h in enumerate(pq_headers, 1):
            val = row_data[h]
            cell = ws_pq.cell(row=r, column=c, value=val)
            if h == "Order_Date" and isinstance(val, datetime):
                cell.number_format = "YYYY-MM-DD"
            elif h in ("Unit_Price", "Sales_Amount", "Cost", "Profit"):
                cell.number_format = '#,##0.00'
            elif h == "Discount":
                cell.number_format = '0%'
            elif h == "Profit_Margin_%":
                cell.number_format = '0.00'

    max_pq_row = len(clean_data) + 1
    style_header_row(ws_pq, 1, len(pq_headers))
    style_data_rows(ws_pq, 2, min(max_pq_row, 5000), len(pq_headers))

    # Auto-fit approximate column widths
    col_widths_pq = {
        1: 14, 2: 12, 3: 6, 4: 5, 5: 5, 6: 10, 7: 20, 8: 16,
        9: 10, 10: 24, 11: 18, 12: 16, 13: 14, 14: 8, 15: 16, 16: 14,
        17: 8, 18: 10, 19: 8, 20: 12, 21: 10, 22: 10, 23: 10, 24: 14, 25: 12,
    }
    for c, w in col_widths_pq.items():
        ws_pq.column_dimensions[get_column_letter(c)].width = w

    # Freeze panes
    ws_pq.freeze_panes = "A2"

    print(f"  ✓ Power_Query_Output — {len(clean_data):,} rows")

    # -----------------------------------------------------------------------
    # SHEET 3: Lookup_Data  (Products + Customers on same sheet)
    # -----------------------------------------------------------------------
    ws_lookup = wb.create_sheet("Lookup_Data")
    ws_lookup.sheet_properties.tabColor = "6C63FF"

    # Products table
    ws_lookup.cell(row=1, column=1, value="PRODUCT LOOKUP TABLE")
    ws_lookup.cell(row=1, column=1).font = Font(name="Calibri", size=12,
                                                 bold=True, color=CLR_PRIMARY)
    prod_headers = ["Product_ID", "Product_Name", "Category", "Subcategory",
                    "Brand", "Standard_Cost"]
    for c, h in enumerate(prod_headers, 1):
        ws_lookup.cell(row=2, column=c, value=h)
    style_header_row(ws_lookup, 2, len(prod_headers))

    for r, p in enumerate(products, 3):
        ws_lookup.cell(row=r, column=1, value=p["Product_ID"])
        ws_lookup.cell(row=r, column=2, value=p["Product_Name"])
        ws_lookup.cell(row=r, column=3, value=p["Category"])
        ws_lookup.cell(row=r, column=4, value=p["Subcategory"])
        ws_lookup.cell(row=r, column=5, value=p["Brand"])
        ws_lookup.cell(row=r, column=6, value=p["Standard_Cost"])
        ws_lookup.cell(row=r, column=6).number_format = '#,##0.00'

    prod_end = 2 + len(products)
    style_data_rows(ws_lookup, 3, prod_end, len(prod_headers))

    # Customers table (starting at column 9)
    cust_start_col = 9
    ws_lookup.cell(row=1, column=cust_start_col, value="CUSTOMER LOOKUP TABLE")
    ws_lookup.cell(row=1, column=cust_start_col).font = Font(
        name="Calibri", size=12, bold=True, color=CLR_PRIMARY
    )
    cust_headers = ["Customer_ID", "Customer_Name", "Customer_Segment",
                    "City", "State", "Region"]
    for c, h in enumerate(cust_headers):
        ws_lookup.cell(row=2, column=cust_start_col + c, value=h)
    style_header_row(ws_lookup, 2, cust_start_col + len(cust_headers) - 1)
    # Fix: only style lookup header cols
    for col in range(cust_start_col, cust_start_col + len(cust_headers)):
        cell = ws_lookup.cell(row=2, column=col)
        cell.font = Font(name="Calibri", size=10, bold=True, color=CLR_WHITE)
        cell.fill = PatternFill(start_color=CLR_HEADER_BG, end_color=CLR_HEADER_BG,
                                fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = THIN_BORDER

    for r, cust in enumerate(customers, 3):
        ws_lookup.cell(row=r, column=cust_start_col, value=cust["Customer_ID"])
        ws_lookup.cell(row=r, column=cust_start_col + 1, value=cust["Customer_Name"])
        ws_lookup.cell(row=r, column=cust_start_col + 2, value=cust["Customer_Segment"])
        ws_lookup.cell(row=r, column=cust_start_col + 3, value=cust["City"])
        ws_lookup.cell(row=r, column=cust_start_col + 4, value=cust["State"])
        ws_lookup.cell(row=r, column=cust_start_col + 5, value=cust["Region"])

    cust_end = 2 + len(customers)

    # Column widths for lookup
    for c in range(1, 7):
        ws_lookup.column_dimensions[get_column_letter(c)].width = 18
    for c in range(cust_start_col, cust_start_col + 6):
        ws_lookup.column_dimensions[get_column_letter(c)].width = 18

    ws_lookup.freeze_panes = "A3"

    print(f"  ✓ Lookup_Data — {len(products)} products, {len(customers)} customers")

    # -----------------------------------------------------------------------
    # SHEET 4: XLOOKUP_Demo
    # -----------------------------------------------------------------------
    ws_xlookup = wb.create_sheet("XLOOKUP_Demo")
    ws_xlookup.sheet_properties.tabColor = CLR_ACCENT

    ws_xlookup.cell(row=1, column=1, value="XLOOKUP DEMONSTRATION")
    ws_xlookup.cell(row=1, column=1).font = Font(name="Calibri", size=14,
                                                   bold=True, color=CLR_PRIMARY)

    ws_xlookup.cell(row=2, column=1,
                    value="These formulas use XLOOKUP to enrich transaction data "
                          "by looking up Product and Customer details.")
    ws_xlookup.cell(row=2, column=1).font = Font(name="Calibri", size=10,
                                                   italic=True, color="666666")

    xl_headers = [
        "Order_ID", "Product_ID", "Customer_ID",
        "Product_Name (XLOOKUP)", "Category (XLOOKUP)",
        "Subcategory (XLOOKUP)", "Brand (XLOOKUP)",
        "Customer_Name (XLOOKUP)", "Customer_Segment (XLOOKUP)",
        "Region (XLOOKUP)",
    ]
    for c, h in enumerate(xl_headers, 1):
        ws_xlookup.cell(row=4, column=c, value=h)
    style_header_row(ws_xlookup, 4, len(xl_headers))

    # Write first 200 rows of data + XLOOKUP formulas
    sample_size = min(200, len(clean_data))
    for i in range(sample_size):
        r = i + 5
        row_data = clean_data[i]
        ws_xlookup.cell(row=r, column=1, value=row_data["Order_ID"])
        ws_xlookup.cell(row=r, column=2, value=row_data["Product_ID"])
        ws_xlookup.cell(row=r, column=3, value=row_data["Customer_ID"])

        # XLOOKUP formulas referencing Lookup_Data sheet
        # Product_Name
        ws_xlookup.cell(row=r, column=4).value = (
            f'=XLOOKUP(B{r},Lookup_Data!A$3:A${prod_end},'
            f'Lookup_Data!B$3:B${prod_end},"Not Found")'
        )
        # Category
        ws_xlookup.cell(row=r, column=5).value = (
            f'=XLOOKUP(B{r},Lookup_Data!A$3:A${prod_end},'
            f'Lookup_Data!C$3:C${prod_end},"Not Found")'
        )
        # Subcategory
        ws_xlookup.cell(row=r, column=6).value = (
            f'=XLOOKUP(B{r},Lookup_Data!A$3:A${prod_end},'
            f'Lookup_Data!D$3:D${prod_end},"Not Found")'
        )
        # Brand
        ws_xlookup.cell(row=r, column=7).value = (
            f'=XLOOKUP(B{r},Lookup_Data!A$3:A${prod_end},'
            f'Lookup_Data!E$3:E${prod_end},"Not Found")'
        )
        # Customer_Name
        ws_xlookup.cell(row=r, column=8).value = (
            f'=XLOOKUP(C{r},Lookup_Data!I$3:I${cust_end},'
            f'Lookup_Data!J$3:J${cust_end},"Not Found")'
        )
        # Customer_Segment
        ws_xlookup.cell(row=r, column=9).value = (
            f'=XLOOKUP(C{r},Lookup_Data!I$3:I${cust_end},'
            f'Lookup_Data!K$3:K${cust_end},"Not Found")'
        )
        # Region
        ws_xlookup.cell(row=r, column=10).value = (
            f'=XLOOKUP(C{r},Lookup_Data!I$3:I${cust_end},'
            f'Lookup_Data!N$3:N${cust_end},"Not Found")'
        )

    for c in range(1, 11):
        ws_xlookup.column_dimensions[get_column_letter(c)].width = 22
    ws_xlookup.freeze_panes = "A5"

    # Add note about Power Query merge vs XLOOKUP
    note_row = sample_size + 7
    ws_xlookup.cell(row=note_row, column=1,
                    value="NOTE: XLOOKUP vs Power Query Merge")
    ws_xlookup.cell(row=note_row, column=1).font = Font(
        name="Calibri", size=12, bold=True, color=CLR_DANGER
    )
    ws_xlookup.cell(row=note_row + 1, column=1,
                    value="• XLOOKUP is ideal for ad-hoc lookups and small datasets "
                          "where formulas are preferred.")
    ws_xlookup.cell(row=note_row + 2, column=1,
                    value="• Power Query Merge is better for large datasets (10,000+ rows) "
                          "as it runs during refresh, not on every cell recalculation.")
    ws_xlookup.cell(row=note_row + 3, column=1,
                    value="• In this project, the main Power_Query_Output sheet uses "
                          "a merge approach (simulated here via Python consolidation).")
    ws_xlookup.cell(row=note_row + 4, column=1,
                    value="• This XLOOKUP_Demo sheet demonstrates the formula approach "
                          "for comparison and portfolio demonstration purposes.")
    for dr in range(note_row + 1, note_row + 5):
        ws_xlookup.cell(row=dr, column=1).font = Font(name="Calibri", size=10,
                                                        color="444444")

    print(f"  ✓ XLOOKUP_Demo — {sample_size} formula rows")

    # -----------------------------------------------------------------------
    # SHEET 5: KPI_Calculations
    # -----------------------------------------------------------------------
    ws_kpi = wb.create_sheet("KPI_Calculations")
    ws_kpi.sheet_properties.tabColor = CLR_SUCCESS

    ws_kpi.cell(row=1, column=1, value="KEY PERFORMANCE INDICATORS")
    ws_kpi.cell(row=1, column=1).font = Font(name="Calibri", size=14,
                                               bold=True, color=CLR_PRIMARY)

    # Calculate KPIs from clean data
    total_sales = sum(r["Sales_Amount"] for r in clean_data)
    total_profit = sum(r["Profit"] for r in clean_data)
    total_cost = sum(r["Cost"] for r in clean_data)
    total_orders = len(clean_data)
    total_units = sum(r["Quantity"] for r in clean_data)
    avg_order_value = total_sales / total_orders if total_orders > 0 else 0
    profit_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0
    avg_discount = sum(r["Discount"] for r in clean_data) / total_orders

    kpi_items = [
        ("Total Sales", total_sales, "₹#,##0"),
        ("Total Profit", total_profit, "₹#,##0"),
        ("Total Cost", total_cost, "₹#,##0"),
        ("Profit Margin (%)", profit_margin, "0.00\"%\""),
        ("Total Orders", total_orders, "#,##0"),
        ("Total Units Sold", total_units, "#,##0"),
        ("Average Order Value", avg_order_value, "₹#,##0.00"),
        ("Average Discount", avg_discount, "0.00%"),
    ]

    kpi_headers = ["KPI Metric", "Value"]
    for c, h in enumerate(kpi_headers, 1):
        ws_kpi.cell(row=3, column=c, value=h)
    style_header_row(ws_kpi, 3, 2)

    for i, (label, val, fmt) in enumerate(kpi_items, 4):
        ws_kpi.cell(row=i, column=1, value=label)
        ws_kpi.cell(row=i, column=1).font = Font(name="Calibri", size=11,
                                                   bold=True, color=CLR_TEXT_DARK)
        cell = ws_kpi.cell(row=i, column=2, value=val)
        cell.number_format = fmt
        cell.font = Font(name="Calibri", size=11, color=CLR_TEXT_DARK)
        cell.border = THIN_BORDER
        ws_kpi.cell(row=i, column=1).border = THIN_BORDER

    ws_kpi.column_dimensions["A"].width = 24
    ws_kpi.column_dimensions["B"].width = 20

    # -----------------------------------------------------------------------
    # AGGREGATION TABLES for Pivot Analysis + Charts
    # -----------------------------------------------------------------------

    # --- Sales by Region ---
    region_sales = defaultdict(lambda: {"Sales": 0, "Profit": 0, "Orders": 0, "Units": 0})
    for r in clean_data:
        region_sales[r["Region"]]["Sales"] += r["Sales_Amount"]
        region_sales[r["Region"]]["Profit"] += r["Profit"]
        region_sales[r["Region"]]["Orders"] += 1
        region_sales[r["Region"]]["Units"] += r["Quantity"]

    # --- Sales by Category ---
    cat_sales = defaultdict(lambda: {"Sales": 0, "Profit": 0, "Orders": 0})
    for r in clean_data:
        cat_sales[r["Category"]]["Sales"] += r["Sales_Amount"]
        cat_sales[r["Category"]]["Profit"] += r["Profit"]
        cat_sales[r["Category"]]["Orders"] += 1

    # --- Sales by Month ---
    month_sales = defaultdict(lambda: {"Sales": 0, "Profit": 0, "Orders": 0})
    for r in clean_data:
        mk = r["Month_Num"]
        month_sales[mk]["Sales"] += r["Sales_Amount"]
        month_sales[mk]["Profit"] += r["Profit"]
        month_sales[mk]["Orders"] += 1

    # --- Top 10 Products by Sales ---
    prod_sales = defaultdict(lambda: {"Sales": 0, "Profit": 0, "Qty": 0})
    for r in clean_data:
        prod_sales[r["Product_Name"]]["Sales"] += r["Sales_Amount"]
        prod_sales[r["Product_Name"]]["Profit"] += r["Profit"]
        prod_sales[r["Product_Name"]]["Qty"] += r["Quantity"]

    top10_by_sales = sorted(prod_sales.items(), key=lambda x: x[1]["Sales"],
                            reverse=True)[:10]
    top10_by_profit = sorted(prod_sales.items(), key=lambda x: x[1]["Profit"],
                             reverse=True)[:10]

    # --- Sales by Customer Segment ---
    seg_sales = defaultdict(lambda: {"Sales": 0, "Profit": 0, "Orders": 0})
    for r in clean_data:
        seg_sales[r["Customer_Segment"]]["Sales"] += r["Sales_Amount"]
        seg_sales[r["Customer_Segment"]]["Profit"] += r["Profit"]
        seg_sales[r["Customer_Segment"]]["Orders"] += 1

    # --- Sales by Channel ---
    chan_sales = defaultdict(lambda: {"Sales": 0, "Profit": 0, "Orders": 0})
    for r in clean_data:
        chan_sales[r["Sales_Channel"]]["Sales"] += r["Sales_Amount"]
        chan_sales[r["Sales_Channel"]]["Profit"] += r["Profit"]
        chan_sales[r["Sales_Channel"]]["Orders"] += 1

    # -----------------------------------------------------------------------
    # SHEET 6: Pivot_Analysis
    # -----------------------------------------------------------------------
    ws_pivot = wb.create_sheet("Pivot_Analysis")
    ws_pivot.sheet_properties.tabColor = "E76F51"

    ws_pivot.cell(row=1, column=1, value="PIVOT TABLE ANALYSIS")
    ws_pivot.cell(row=1, column=1).font = Font(name="Calibri", size=14,
                                                 bold=True, color=CLR_PRIMARY)

    current_row = 3

    # --- Table 1: Sales & Profit by Region ---
    ws_pivot.cell(row=current_row, column=1,
                  value="1. Total Sales & Profit by Region")
    ws_pivot.cell(row=current_row, column=1).font = Font(
        name="Calibri", size=12, bold=True, color=CLR_SECONDARY
    )
    current_row += 1
    reg_headers = ["Region", "Total Sales", "Total Profit", "Total Orders",
                   "Total Units", "Profit Margin %"]
    for c, h in enumerate(reg_headers, 1):
        ws_pivot.cell(row=current_row, column=c, value=h)
    style_header_row(ws_pivot, current_row, len(reg_headers))
    current_row += 1

    reg_data_start = current_row
    for region in sorted(region_sales.keys()):
        d = region_sales[region]
        margin = (d["Profit"] / d["Sales"] * 100) if d["Sales"] > 0 else 0
        ws_pivot.cell(row=current_row, column=1, value=region)
        ws_pivot.cell(row=current_row, column=2, value=round(d["Sales"], 2))
        ws_pivot.cell(row=current_row, column=2).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=3, value=round(d["Profit"], 2))
        ws_pivot.cell(row=current_row, column=3).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=4, value=d["Orders"])
        ws_pivot.cell(row=current_row, column=4).number_format = '#,##0'
        ws_pivot.cell(row=current_row, column=5, value=d["Units"])
        ws_pivot.cell(row=current_row, column=5).number_format = '#,##0'
        ws_pivot.cell(row=current_row, column=6, value=round(margin, 2))
        ws_pivot.cell(row=current_row, column=6).number_format = '0.00"%"'
        current_row += 1
    reg_data_end = current_row - 1
    style_data_rows(ws_pivot, reg_data_start, reg_data_end, len(reg_headers))

    current_row += 2

    # --- Table 2: Sales by Category ---
    ws_pivot.cell(row=current_row, column=1,
                  value="2. Sales by Product Category")
    ws_pivot.cell(row=current_row, column=1).font = Font(
        name="Calibri", size=12, bold=True, color=CLR_SECONDARY
    )
    current_row += 1
    cat_headers_list = ["Category", "Total Sales", "Total Profit", "Orders",
                        "Profit Margin %"]
    for c, h in enumerate(cat_headers_list, 1):
        ws_pivot.cell(row=current_row, column=c, value=h)
    style_header_row(ws_pivot, current_row, len(cat_headers_list))
    current_row += 1

    cat_data_start = current_row
    for cat in sorted(cat_sales.keys()):
        d = cat_sales[cat]
        margin = (d["Profit"] / d["Sales"] * 100) if d["Sales"] > 0 else 0
        ws_pivot.cell(row=current_row, column=1, value=cat)
        ws_pivot.cell(row=current_row, column=2, value=round(d["Sales"], 2))
        ws_pivot.cell(row=current_row, column=2).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=3, value=round(d["Profit"], 2))
        ws_pivot.cell(row=current_row, column=3).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=4, value=d["Orders"])
        ws_pivot.cell(row=current_row, column=4).number_format = '#,##0'
        ws_pivot.cell(row=current_row, column=5, value=round(margin, 2))
        ws_pivot.cell(row=current_row, column=5).number_format = '0.00"%"'
        current_row += 1
    cat_data_end = current_row - 1
    style_data_rows(ws_pivot, cat_data_start, cat_data_end, len(cat_headers_list))

    current_row += 2

    # --- Table 3: Monthly Sales ---
    ws_pivot.cell(row=current_row, column=1, value="3. Monthly Sales & Profit")
    ws_pivot.cell(row=current_row, column=1).font = Font(
        name="Calibri", size=12, bold=True, color=CLR_SECONDARY
    )
    current_row += 1
    month_headers = ["Month", "Total Sales", "Total Profit", "Orders",
                     "MoM Growth %"]
    for c, h in enumerate(month_headers, 1):
        ws_pivot.cell(row=current_row, column=c, value=h)
    style_header_row(ws_pivot, current_row, len(month_headers))
    current_row += 1

    month_names_list = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                        "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    month_data_start = current_row
    prev_sales = None
    for m in range(1, 13):
        d = month_sales[m]
        growth = 0
        if prev_sales and prev_sales > 0:
            growth = ((d["Sales"] - prev_sales) / prev_sales) * 100
        ws_pivot.cell(row=current_row, column=1, value=month_names_list[m - 1])
        ws_pivot.cell(row=current_row, column=2, value=round(d["Sales"], 2))
        ws_pivot.cell(row=current_row, column=2).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=3, value=round(d["Profit"], 2))
        ws_pivot.cell(row=current_row, column=3).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=4, value=d["Orders"])
        ws_pivot.cell(row=current_row, column=4).number_format = '#,##0'
        ws_pivot.cell(row=current_row, column=5, value=round(growth, 2))
        ws_pivot.cell(row=current_row, column=5).number_format = '0.00"%"'
        prev_sales = d["Sales"]
        current_row += 1
    month_data_end = current_row - 1
    style_data_rows(ws_pivot, month_data_start, month_data_end, len(month_headers))

    # Conditional formatting for MoM Growth
    green_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA",
                             fill_type="solid")
    red_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA",
                           fill_type="solid")
    ws_pivot.conditional_formatting.add(
        f"E{month_data_start}:E{month_data_end}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=green_fill)
    )
    ws_pivot.conditional_formatting.add(
        f"E{month_data_start}:E{month_data_end}",
        CellIsRule(operator="lessThan", formula=["0"], fill=red_fill)
    )

    current_row += 2

    # --- Table 4: Top 10 Products by Sales ---
    ws_pivot.cell(row=current_row, column=1,
                  value="4. Top 10 Products by Sales")
    ws_pivot.cell(row=current_row, column=1).font = Font(
        name="Calibri", size=12, bold=True, color=CLR_SECONDARY
    )
    current_row += 1
    top_headers = ["Rank", "Product Name", "Total Sales", "Total Profit",
                   "Quantity Sold"]
    for c, h in enumerate(top_headers, 1):
        ws_pivot.cell(row=current_row, column=c, value=h)
    style_header_row(ws_pivot, current_row, len(top_headers))
    current_row += 1

    top10_sales_start = current_row
    for rank, (name, d) in enumerate(top10_by_sales, 1):
        ws_pivot.cell(row=current_row, column=1, value=rank)
        ws_pivot.cell(row=current_row, column=2, value=name)
        ws_pivot.cell(row=current_row, column=3, value=round(d["Sales"], 2))
        ws_pivot.cell(row=current_row, column=3).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=4, value=round(d["Profit"], 2))
        ws_pivot.cell(row=current_row, column=4).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=5, value=d["Qty"])
        ws_pivot.cell(row=current_row, column=5).number_format = '#,##0'
        current_row += 1
    top10_sales_end = current_row - 1
    style_data_rows(ws_pivot, top10_sales_start, top10_sales_end, len(top_headers))

    current_row += 2

    # --- Table 5: Top 10 Products by Profit ---
    ws_pivot.cell(row=current_row, column=1,
                  value="5. Top 10 Products by Profit")
    ws_pivot.cell(row=current_row, column=1).font = Font(
        name="Calibri", size=12, bold=True, color=CLR_SECONDARY
    )
    current_row += 1
    for c, h in enumerate(top_headers, 1):
        ws_pivot.cell(row=current_row, column=c, value=h)
    style_header_row(ws_pivot, current_row, len(top_headers))
    current_row += 1

    top10_profit_start = current_row
    for rank, (name, d) in enumerate(top10_by_profit, 1):
        ws_pivot.cell(row=current_row, column=1, value=rank)
        ws_pivot.cell(row=current_row, column=2, value=name)
        ws_pivot.cell(row=current_row, column=3, value=round(d["Sales"], 2))
        ws_pivot.cell(row=current_row, column=3).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=4, value=round(d["Profit"], 2))
        ws_pivot.cell(row=current_row, column=4).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=5, value=d["Qty"])
        ws_pivot.cell(row=current_row, column=5).number_format = '#,##0'
        current_row += 1
    top10_profit_end = current_row - 1
    style_data_rows(ws_pivot, top10_profit_start, top10_profit_end, len(top_headers))

    current_row += 2

    # --- Table 6: Sales by Customer Segment ---
    ws_pivot.cell(row=current_row, column=1,
                  value="6. Sales by Customer Segment")
    ws_pivot.cell(row=current_row, column=1).font = Font(
        name="Calibri", size=12, bold=True, color=CLR_SECONDARY
    )
    current_row += 1
    seg_hdr = ["Customer Segment", "Total Sales", "Total Profit", "Orders",
               "Profit Margin %"]
    for c, h in enumerate(seg_hdr, 1):
        ws_pivot.cell(row=current_row, column=c, value=h)
    style_header_row(ws_pivot, current_row, len(seg_hdr))
    current_row += 1

    seg_data_start = current_row
    for seg in sorted(seg_sales.keys()):
        d = seg_sales[seg]
        margin = (d["Profit"] / d["Sales"] * 100) if d["Sales"] > 0 else 0
        ws_pivot.cell(row=current_row, column=1, value=seg)
        ws_pivot.cell(row=current_row, column=2, value=round(d["Sales"], 2))
        ws_pivot.cell(row=current_row, column=2).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=3, value=round(d["Profit"], 2))
        ws_pivot.cell(row=current_row, column=3).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=4, value=d["Orders"])
        ws_pivot.cell(row=current_row, column=4).number_format = '#,##0'
        ws_pivot.cell(row=current_row, column=5, value=round(margin, 2))
        ws_pivot.cell(row=current_row, column=5).number_format = '0.00"%"'
        current_row += 1
    seg_data_end = current_row - 1
    style_data_rows(ws_pivot, seg_data_start, seg_data_end, len(seg_hdr))

    current_row += 2

    # --- Table 7: Sales by Sales Channel ---
    ws_pivot.cell(row=current_row, column=1,
                  value="7. Sales by Sales Channel")
    ws_pivot.cell(row=current_row, column=1).font = Font(
        name="Calibri", size=12, bold=True, color=CLR_SECONDARY
    )
    current_row += 1
    chan_hdr = ["Sales Channel", "Total Sales", "Total Profit", "Orders",
                "Profit Margin %"]
    for c, h in enumerate(chan_hdr, 1):
        ws_pivot.cell(row=current_row, column=c, value=h)
    style_header_row(ws_pivot, current_row, len(chan_hdr))
    current_row += 1

    chan_data_start = current_row
    for ch in sorted(chan_sales.keys()):
        d = chan_sales[ch]
        margin = (d["Profit"] / d["Sales"] * 100) if d["Sales"] > 0 else 0
        ws_pivot.cell(row=current_row, column=1, value=ch)
        ws_pivot.cell(row=current_row, column=2, value=round(d["Sales"], 2))
        ws_pivot.cell(row=current_row, column=2).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=3, value=round(d["Profit"], 2))
        ws_pivot.cell(row=current_row, column=3).number_format = '₹#,##0'
        ws_pivot.cell(row=current_row, column=4, value=d["Orders"])
        ws_pivot.cell(row=current_row, column=4).number_format = '#,##0'
        ws_pivot.cell(row=current_row, column=5, value=round(margin, 2))
        ws_pivot.cell(row=current_row, column=5).number_format = '0.00"%"'
        current_row += 1
    chan_data_end = current_row - 1
    style_data_rows(ws_pivot, chan_data_start, chan_data_end, len(chan_hdr))

    for c in range(1, 7):
        ws_pivot.column_dimensions[get_column_letter(c)].width = 20

    ws_pivot.freeze_panes = "A2"

    print("  ✓ Pivot_Analysis — 7 analysis tables created")
    print(f"  ✓ KPI_Calculations — {len(kpi_items)} KPIs")

    # -----------------------------------------------------------------------
    # SHEET 7: Raw_Data (first 1000 rows of raw data for reference)
    # -----------------------------------------------------------------------
    ws_raw = wb.create_sheet("Raw_Data")
    ws_raw.sheet_properties.tabColor = "999999"

    ws_raw.cell(row=1, column=1,
                value="RAW DATA SAMPLE (Before Cleaning)")
    ws_raw.cell(row=1, column=1).font = Font(name="Calibri", size=12,
                                               bold=True, color=CLR_DANGER)
    ws_raw.cell(row=2, column=1,
                value="This sheet shows a sample of raw transactional data "
                      "BEFORE Power Query cleaning. Notice data quality issues.")
    ws_raw.cell(row=2, column=1).font = Font(name="Calibri", size=10,
                                               italic=True, color="666666")

    for c, h in enumerate(SALES_FIELDS, 1):
        ws_raw.cell(row=3, column=c, value=h)
    style_header_row(ws_raw, 3, len(SALES_FIELDS))

    # Write raw rows (first 1000 from first few months)
    raw_rows = []
    for m in range(1, 13):
        raw_rows.extend(monthly_data[m])
    raw_sample = raw_rows[:1000]

    for r, row_data in enumerate(raw_sample, 4):
        for c, field in enumerate(SALES_FIELDS, 1):
            ws_raw.cell(row=r, column=c, value=row_data[field])

    for c in range(1, len(SALES_FIELDS) + 1):
        ws_raw.column_dimensions[get_column_letter(c)].width = 14

    ws_raw.freeze_panes = "A4"
    print(f"  ✓ Raw_Data — {len(raw_sample)} sample rows")

    # -----------------------------------------------------------------------
    # DASHBOARD — Charts & KPI Cards
    # -----------------------------------------------------------------------
    print("\n  Building Dashboard...")

    # Set dashboard background
    for r in range(1, 55):
        for c in range(1, 22):
            cell = ws_dash.cell(row=r, column=c)
            cell.fill = PatternFill(start_color=CLR_BG_LIGHT,
                                    end_color=CLR_BG_LIGHT,
                                    fill_type="solid")

    # Dashboard Title
    ws_dash.merge_cells("A1:U1")
    title_cell = ws_dash.cell(row=1, column=1,
                               value="📊  SALES PERFORMANCE DASHBOARD  —  FY 2025")
    title_cell.font = Font(name="Calibri", size=20, bold=True, color=CLR_WHITE)
    title_cell.fill = PatternFill(start_color=CLR_PRIMARY, end_color=CLR_PRIMARY,
                                  fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.row_dimensions[1].height = 45

    # Subtitle
    ws_dash.merge_cells("A2:U2")
    sub_cell = ws_dash.cell(row=2, column=1,
                             value="Executive Summary  |  All Regions  |  "
                                   "All Categories  |  Data Refreshed: "
                                   f"{datetime.now().strftime('%d-%b-%Y')}")
    sub_cell.font = Font(name="Calibri", size=10, italic=True, color=CLR_WHITE)
    sub_cell.fill = PatternFill(start_color=CLR_SECONDARY, end_color=CLR_SECONDARY,
                                fill_type="solid")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_dash.row_dimensions[2].height = 28

    # KPI Cards — Row 4-5
    ws_dash.row_dimensions[4].height = 24
    ws_dash.row_dimensions[5].height = 36

    kpi_positions = [
        (4, 2, "TOTAL SALES", total_sales, "₹#,##0", CLR_SECONDARY, 3),
        (4, 6, "TOTAL PROFIT", total_profit, "₹#,##0", CLR_SUCCESS, 3),
        (4, 10, "PROFIT MARGIN", profit_margin, '0.0"%"', CLR_ACCENT, 3),
        (4, 14, "TOTAL ORDERS", total_orders, "#,##0", "6C63FF", 3),
        (4, 17, "UNITS SOLD", total_units, "#,##0", CLR_DANGER, 2),
        (4, 20, "AVG ORDER", avg_order_value, "₹#,##0", "264653", 2),
    ]

    for (row, col, title, val, fmt, clr, w) in kpi_positions:
        write_kpi_card(ws_dash, row, col, title, val, fmt, w, clr)

    # Set column widths for dashboard
    for c in range(1, 22):
        ws_dash.column_dimensions[get_column_letter(c)].width = 7

    # -----------------------------------------------------------------------
    # CHART 1: Monthly Sales Trend (Line Chart)
    # -----------------------------------------------------------------------

    # Write chart data to hidden area (rows 58+)
    chart_data_row = 58

    # Monthly trend data
    ws_dash.cell(row=chart_data_row, column=1, value="Month")
    ws_dash.cell(row=chart_data_row, column=2, value="Sales")
    ws_dash.cell(row=chart_data_row, column=3, value="Profit")
    for i, m in enumerate(range(1, 13)):
        ws_dash.cell(row=chart_data_row + 1 + i, column=1,
                     value=month_names_list[m - 1])
        ws_dash.cell(row=chart_data_row + 1 + i, column=2,
                     value=round(month_sales[m]["Sales"], 0))
        ws_dash.cell(row=chart_data_row + 1 + i, column=3,
                     value=round(month_sales[m]["Profit"], 0))

    chart1 = LineChart()
    chart1.title = "Monthly Sales & Profit Trend"
    chart1.style = 10
    chart1.y_axis.title = "Amount (₹)"
    chart1.x_axis.title = "Month"
    chart1.height = 12
    chart1.width = 22

    cats = Reference(ws_dash, min_col=1, min_row=chart_data_row + 1,
                     max_row=chart_data_row + 12)
    sales_data = Reference(ws_dash, min_col=2, min_row=chart_data_row,
                           max_row=chart_data_row + 12)
    profit_data = Reference(ws_dash, min_col=3, min_row=chart_data_row,
                            max_row=chart_data_row + 12)

    chart1.add_data(sales_data, titles_from_data=True)
    chart1.add_data(profit_data, titles_from_data=True)
    chart1.set_categories(cats)

    # Style the lines
    chart1.series[0].graphicalProperties.line.width = 25000
    chart1.series[1].graphicalProperties.line.width = 25000
    chart1.series[0].graphicalProperties.line.solidFill = CLR_SECONDARY
    chart1.series[1].graphicalProperties.line.solidFill = CLR_SUCCESS

    ws_dash.add_chart(chart1, "B7")

    # -----------------------------------------------------------------------
    # CHART 2: Sales by Region (Bar Chart)
    # -----------------------------------------------------------------------
    region_row = chart_data_row + 15
    ws_dash.cell(row=region_row, column=1, value="Region")
    ws_dash.cell(row=region_row, column=2, value="Sales")
    ws_dash.cell(row=region_row, column=3, value="Profit")

    sorted_regions = sorted(region_sales.keys())
    for i, reg in enumerate(sorted_regions):
        ws_dash.cell(row=region_row + 1 + i, column=1, value=reg)
        ws_dash.cell(row=region_row + 1 + i, column=2,
                     value=round(region_sales[reg]["Sales"], 0))
        ws_dash.cell(row=region_row + 1 + i, column=3,
                     value=round(region_sales[reg]["Profit"], 0))

    chart2 = BarChart()
    chart2.type = "col"
    chart2.title = "Sales & Profit by Region"
    chart2.style = 10
    chart2.y_axis.title = "Amount (₹)"
    chart2.height = 12
    chart2.width = 13

    cats2 = Reference(ws_dash, min_col=1, min_row=region_row + 1,
                      max_row=region_row + len(sorted_regions))
    data2_s = Reference(ws_dash, min_col=2, min_row=region_row,
                        max_row=region_row + len(sorted_regions))
    data2_p = Reference(ws_dash, min_col=3, min_row=region_row,
                        max_row=region_row + len(sorted_regions))

    chart2.add_data(data2_s, titles_from_data=True)
    chart2.add_data(data2_p, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.series[0].graphicalProperties.solidFill = CLR_SECONDARY
    chart2.series[1].graphicalProperties.solidFill = CLR_SUCCESS

    ws_dash.add_chart(chart2, "L7")

    # -----------------------------------------------------------------------
    # CHART 3: Sales by Category (Horizontal Bar)
    # -----------------------------------------------------------------------
    cat_row = region_row + len(sorted_regions) + 3
    ws_dash.cell(row=cat_row, column=1, value="Category")
    ws_dash.cell(row=cat_row, column=2, value="Sales")

    sorted_cats = sorted(cat_sales.keys(), key=lambda x: cat_sales[x]["Sales"],
                         reverse=True)
    for i, cat in enumerate(sorted_cats):
        ws_dash.cell(row=cat_row + 1 + i, column=1, value=cat)
        ws_dash.cell(row=cat_row + 1 + i, column=2,
                     value=round(cat_sales[cat]["Sales"], 0))

    chart3 = BarChart()
    chart3.type = "bar"
    chart3.title = "Sales by Product Category"
    chart3.style = 10
    chart3.x_axis.title = "Amount (₹)"
    chart3.height = 12
    chart3.width = 13

    cats3 = Reference(ws_dash, min_col=1, min_row=cat_row + 1,
                      max_row=cat_row + len(sorted_cats))
    data3 = Reference(ws_dash, min_col=2, min_row=cat_row,
                      max_row=cat_row + len(sorted_cats))

    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)

    colors3 = [CLR_CHART_COLORS[i % len(CLR_CHART_COLORS)]
               for i in range(len(sorted_cats))]
    for i, c in enumerate(colors3):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = c
        chart3.series[0].data_points.append(pt)

    ws_dash.add_chart(chart3, "B23")

    # -----------------------------------------------------------------------
    # CHART 4: Top 10 Products (Horizontal Bar)
    # -----------------------------------------------------------------------
    top_row = cat_row + len(sorted_cats) + 3
    ws_dash.cell(row=top_row, column=1, value="Product")
    ws_dash.cell(row=top_row, column=2, value="Sales")

    for i, (name, d) in enumerate(top10_by_sales):
        ws_dash.cell(row=top_row + 1 + i, column=1, value=name)
        ws_dash.cell(row=top_row + 1 + i, column=2,
                     value=round(d["Sales"], 0))

    chart4 = BarChart()
    chart4.type = "bar"
    chart4.title = "Top 10 Products by Sales"
    chart4.style = 10
    chart4.x_axis.title = "Amount (₹)"
    chart4.height = 12
    chart4.width = 13

    cats4 = Reference(ws_dash, min_col=1, min_row=top_row + 1,
                      max_row=top_row + 10)
    data4 = Reference(ws_dash, min_col=2, min_row=top_row,
                      max_row=top_row + 10)

    chart4.add_data(data4, titles_from_data=True)
    chart4.set_categories(cats4)
    chart4.series[0].graphicalProperties.solidFill = CLR_ACCENT

    ws_dash.add_chart(chart4, "L23")

    # -----------------------------------------------------------------------
    # CHART 5: Profit by Region (Pie Chart)
    # -----------------------------------------------------------------------
    prof_reg_row = top_row + 13
    ws_dash.cell(row=prof_reg_row, column=1, value="Region")
    ws_dash.cell(row=prof_reg_row, column=2, value="Profit")

    for i, reg in enumerate(sorted_regions):
        ws_dash.cell(row=prof_reg_row + 1 + i, column=1, value=reg)
        ws_dash.cell(row=prof_reg_row + 1 + i, column=2,
                     value=round(region_sales[reg]["Profit"], 0))

    chart5 = PieChart()
    chart5.title = "Profit Distribution by Region"
    chart5.style = 10
    chart5.height = 12
    chart5.width = 13

    cats5 = Reference(ws_dash, min_col=1, min_row=prof_reg_row + 1,
                      max_row=prof_reg_row + len(sorted_regions))
    data5 = Reference(ws_dash, min_col=2, min_row=prof_reg_row,
                      max_row=prof_reg_row + len(sorted_regions))

    chart5.add_data(data5, titles_from_data=True)
    chart5.set_categories(cats5)
    chart5.dataLabels = DataLabelList()
    chart5.dataLabels.showPercent = True

    for i, c in enumerate(CLR_CHART_COLORS[:len(sorted_regions)]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = c
        chart5.series[0].data_points.append(pt)

    ws_dash.add_chart(chart5, "B39")

    # -----------------------------------------------------------------------
    # CHART 6: Sales by Customer Segment (Pie)
    # -----------------------------------------------------------------------
    seg_row = prof_reg_row + len(sorted_regions) + 3
    ws_dash.cell(row=seg_row, column=1, value="Segment")
    ws_dash.cell(row=seg_row, column=2, value="Sales")

    sorted_segs = sorted(seg_sales.keys())
    for i, seg in enumerate(sorted_segs):
        ws_dash.cell(row=seg_row + 1 + i, column=1, value=seg)
        ws_dash.cell(row=seg_row + 1 + i, column=2,
                     value=round(seg_sales[seg]["Sales"], 0))

    chart6 = PieChart()
    chart6.title = "Sales by Customer Segment"
    chart6.style = 10
    chart6.height = 12
    chart6.width = 13

    cats6 = Reference(ws_dash, min_col=1, min_row=seg_row + 1,
                      max_row=seg_row + len(sorted_segs))
    data6 = Reference(ws_dash, min_col=2, min_row=seg_row,
                      max_row=seg_row + len(sorted_segs))

    chart6.add_data(data6, titles_from_data=True)
    chart6.set_categories(cats6)
    chart6.dataLabels = DataLabelList()
    chart6.dataLabels.showPercent = True

    for i, c in enumerate(CLR_CHART_COLORS[:len(sorted_segs)]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = c
        chart6.series[0].data_points.append(pt)

    ws_dash.add_chart(chart6, "L39")

    # -----------------------------------------------------------------------
    # CHART 7: Sales Channel Performance (Bar)
    # -----------------------------------------------------------------------
    chan_row = seg_row + len(sorted_segs) + 3
    ws_dash.cell(row=chan_row, column=1, value="Channel")
    ws_dash.cell(row=chan_row, column=2, value="Sales")
    ws_dash.cell(row=chan_row, column=3, value="Profit")

    sorted_chans = sorted(chan_sales.keys(), key=lambda x: chan_sales[x]["Sales"],
                          reverse=True)
    for i, ch in enumerate(sorted_chans):
        ws_dash.cell(row=chan_row + 1 + i, column=1, value=ch)
        ws_dash.cell(row=chan_row + 1 + i, column=2,
                     value=round(chan_sales[ch]["Sales"], 0))
        ws_dash.cell(row=chan_row + 1 + i, column=3,
                     value=round(chan_sales[ch]["Profit"], 0))

    chart7 = BarChart()
    chart7.type = "col"
    chart7.title = "Sales Channel Performance"
    chart7.style = 10
    chart7.y_axis.title = "Amount (₹)"
    chart7.height = 12
    chart7.width = 22

    cats7 = Reference(ws_dash, min_col=1, min_row=chan_row + 1,
                      max_row=chan_row + len(sorted_chans))
    data7_s = Reference(ws_dash, min_col=2, min_row=chan_row,
                        max_row=chan_row + len(sorted_chans))
    data7_p = Reference(ws_dash, min_col=3, min_row=chan_row,
                        max_row=chan_row + len(sorted_chans))

    chart7.add_data(data7_s, titles_from_data=True)
    chart7.add_data(data7_p, titles_from_data=True)
    chart7.set_categories(cats7)
    chart7.series[0].graphicalProperties.solidFill = CLR_SECONDARY
    chart7.series[1].graphicalProperties.solidFill = CLR_SUCCESS

    ws_dash.add_chart(chart7, "B55")

    # Hide chart data rows
    for r in range(chart_data_row, chan_row + len(sorted_chans) + 2):
        ws_dash.row_dimensions[r].hidden = True

    # Dashboard footer
    ws_dash.merge_cells("A71:U71")
    footer = ws_dash.cell(row=71, column=1,
                           value="Dashboard created with Power Query + Pivot Tables + "
                                 "XLOOKUP  |  Data Source: Monthly CSV Files  |  "
                                 "Refresh: Data → Refresh All")
    footer.font = Font(name="Calibri", size=8, italic=True, color="888888")
    footer.alignment = Alignment(horizontal="center")

    # Set print area and gridlines off
    ws_dash.sheet_view.showGridLines = False

    # Reorder sheets: Dashboard first
    wb.move_sheet(ws_dash, offset=0)

    print("  ✓ Dashboard — 7 charts + 6 KPI cards")

    # -----------------------------------------------------------------------
    # Save
    # -----------------------------------------------------------------------
    os.makedirs(EXCEL_DIR, exist_ok=True)
    output_path = os.path.join(EXCEL_DIR, "Sales_Performance_Dashboard.xlsx")
    wb.save(output_path)
    print(f"\n  ✅ Workbook saved: {output_path}")

    return (total_sales, total_profit, profit_margin, total_orders,
            total_units, avg_order_value, region_sales, cat_sales,
            month_sales, seg_sales, chan_sales, top10_by_sales, top10_by_profit)


# ---------------------------------------------------------------------------
# POWER QUERY M CODE DOCUMENTATION
# ---------------------------------------------------------------------------

def write_power_query_docs():
    """Write Power Query M code documentation."""
    os.makedirs(DOCS_DIR, exist_ok=True)

    m_code = r'''// ============================================================
// Power Query M Code — Sales Data Import & Transformation
// ============================================================
// This query imports all monthly CSV files from the data/ folder,
// combines them, and performs data cleaning & transformation.
//
// HOW TO USE:
// 1. Open Excel → Data → Get Data → From File → From Folder
// 2. Navigate to the "data" folder in this project
// 3. Click "Combine & Transform Data"
// 4. In the Power Query Editor, replace the generated M code
//    with the code below (or use it as a reference)
// 5. Click "Close & Load"
// ============================================================

let
    // Step 1: Import all CSV files from the data folder
    Source = Folder.Files(".\data"),

    // Step 2: Filter to only CSV files (exclude customers.csv and products.csv)
    FilteredFiles = Table.SelectRows(Source, each
        Text.StartsWith([Name], "sales_") and Text.EndsWith([Name], ".csv")
    ),

    // Step 3: Read each CSV file content
    AddContent = Table.AddColumn(FilteredFiles, "ParsedCSV", each
        Csv.Document([Content], [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None])
    ),

    // Step 4: Expand the parsed CSV tables
    ExpandedTable = Table.ExpandTableColumn(AddContent, "ParsedCSV",
        {"Column1","Column2","Column3","Column4","Column5","Column6","Column7",
         "Column8","Column9","Column10","Column11","Column12","Column13",
         "Column14","Column15"}
    ),

    // Step 5: Promote first row as headers
    PromotedHeaders = Table.PromoteHeaders(ExpandedTable, [PromoteAllScalars=true]),

    // Step 6: Remove unnecessary file metadata columns
    RemovedColumns = Table.RemoveColumns(PromotedHeaders,
        {"Content", "Name", "Extension", "Date accessed",
         "Date modified", "Date created", "Attributes", "Folder Path"},
        MissingField.Ignore
    ),

    // Step 7: Clean blank/null values
    ReplacedNulls_Discount = Table.ReplaceValue(RemovedColumns, null, "0",
        Replacer.ReplaceValue, {"Discount"}),
    ReplacedNulls_Payment = Table.ReplaceValue(ReplacedNulls_Discount, null, "Unknown",
        Replacer.ReplaceValue, {"Payment_Method"}),
    ReplacedBlanks_Customer = Table.ReplaceValue(ReplacedNulls_Payment, "", "Unknown",
        Replacer.ReplaceText, {"Customer_ID"}),

    // Step 8: Remove duplicate Order_IDs
    RemovedDuplicates = Table.Distinct(ReplacedBlanks_Customer, {"Order_ID"}),

    // Step 9: Standardize Region names (trim whitespace + Title Case)
    CleanedRegion = Table.TransformColumns(RemovedDuplicates,
        {{"Region", each Text.Trim(Text.Proper(_)), type text}}
    ),

    // Step 10: Convert dates to proper Date type
    ConvertedDates = Table.TransformColumns(CleanedRegion,
        {{"Order_Date", each
            try Date.From(DateTime.From(_))
            otherwise try Date.FromText(_, [Format="yyyy-MM-dd"])
            otherwise try Date.FromText(_, [Format="dd-MM-yyyy"])
            otherwise try Date.FromText(_, [Format="MM/dd/yyyy"])
            otherwise null,
            type date
        }}
    ),

    // Step 11: Remove rows with null dates
    FilteredValidDates = Table.SelectRows(ConvertedDates,
        each [Order_Date] <> null
    ),

    // Step 12: Convert numeric columns to proper types
    ConvertedTypes = Table.TransformColumnTypes(FilteredValidDates, {
        {"Quantity", Int64.Type},
        {"Unit_Price", type number},
        {"Discount", type number},
        {"Sales_Amount", type number},
        {"Cost", type number},
        {"Profit", type number}
    }),

    // Step 13: Add calculated columns
    AddedMonth = Table.AddColumn(ConvertedTypes, "Month",
        each Date.MonthName([Order_Date]), type text),
    AddedMonthNum = Table.AddColumn(AddedMonth, "Month_Num",
        each Date.Month([Order_Date]), Int64.Type),
    AddedQuarter = Table.AddColumn(AddedMonthNum, "Quarter",
        each "Q" & Text.From(Date.QuarterOfYear([Order_Date])), type text),

    // Step 14: Calculate Revenue (recalculate to ensure consistency)
    AddedRevenue = Table.AddColumn(AddedQuarter, "Calculated_Revenue",
        each [Quantity] * [Unit_Price] * (1 - [Discount]), type number),

    // Step 15: Calculate Profit Margin
    AddedProfitMargin = Table.AddColumn(AddedRevenue, "Profit_Margin_%",
        each if [Sales_Amount] > 0
             then ([Profit] / [Sales_Amount]) * 100
             else 0,
        type number),

    // Final: Reorder columns for clean output
    ReorderedColumns = Table.ReorderColumns(AddedProfitMargin, {
        "Order_ID", "Order_Date", "Month", "Month_Num", "Quarter",
        "Customer_ID", "Product_ID", "Region", "State", "Sales_Channel",
        "Quantity", "Unit_Price", "Discount", "Sales_Amount",
        "Calculated_Revenue", "Cost", "Profit", "Profit_Margin_%",
        "Payment_Method", "Order_Status"
    })
in
    ReorderedColumns


// ============================================================
// ADDITIONAL QUERY: Product Lookup Merge
// ============================================================
// This query can be used to merge product details into the
// sales data instead of using XLOOKUP formulas.

/*
let
    Source = Excel.CurrentWorkbook(){[Name="SalesTable"]}[Content],
    ProductTable = Csv.Document(File.Contents(".\data\products.csv"),
        [Delimiter=",", Encoding=65001]),
    PromotedProducts = Table.PromoteHeaders(ProductTable),
    MergedProducts = Table.NestedJoin(Source, {"Product_ID"},
        PromotedProducts, {"Product_ID"}, "ProductLookup",
        JoinKind.LeftOuter),
    ExpandedProducts = Table.ExpandTableColumn(MergedProducts,
        "ProductLookup",
        {"Product_Name", "Category", "Subcategory", "Brand", "Standard_Cost"})
in
    ExpandedProducts
*/


// ============================================================
// ADDITIONAL QUERY: Customer Lookup Merge
// ============================================================
/*
let
    Source = Excel.CurrentWorkbook(){[Name="SalesTable"]}[Content],
    CustomerTable = Csv.Document(File.Contents(".\data\customers.csv"),
        [Delimiter=",", Encoding=65001]),
    PromotedCustomers = Table.PromoteHeaders(CustomerTable),
    MergedCustomers = Table.NestedJoin(Source, {"Customer_ID"},
        PromotedCustomers, {"Customer_ID"}, "CustomerLookup",
        JoinKind.LeftOuter),
    ExpandedCustomers = Table.ExpandTableColumn(MergedCustomers,
        "CustomerLookup",
        {"Customer_Name", "Customer_Segment", "City"})
in
    ExpandedCustomers
*/
'''

    with open(os.path.join(DOCS_DIR, "power_query_m_code.pq"), "w",
              encoding="utf-8") as f:
        f.write(m_code)

    print(f"  ✓ Power Query M code saved to docs/power_query_m_code.pq")


# ---------------------------------------------------------------------------
# PIVOT TABLE SETUP GUIDE
# ---------------------------------------------------------------------------

def write_pivot_slicer_guide():
    """Write guide for creating Pivot Tables and Slicers in Excel."""
    guide = """# Pivot Tables & Slicers Setup Guide

## Overview
The `Pivot_Analysis` sheet contains pre-calculated summary tables that mirror
what Pivot Tables would produce. To create **actual interactive Pivot Tables
with Slicers and Timeline**, follow these steps in Excel.

---

## Step 1: Create Pivot Tables from Power_Query_Output

1. Go to the **Power_Query_Output** sheet
2. Select any cell in the data range
3. Go to **Insert → PivotTable**
4. Choose "New Worksheet" or "Existing Worksheet"
5. Click OK

### Recommended Pivot Tables:

#### PT1: Total Sales by Region
- Rows: Region
- Values: Sum of Sales_Amount, Sum of Profit

#### PT2: Sales by Product Category
- Rows: Category
- Values: Sum of Sales_Amount, Sum of Profit, Count of Order_ID

#### PT3: Monthly Sales Trend
- Rows: Month (or Order_Date grouped by Month)
- Values: Sum of Sales_Amount, Sum of Profit

#### PT4: Top 10 Products
- Rows: Product_Name
- Values: Sum of Sales_Amount
- Sort: Descending by Sales_Amount
- Filter: Top 10

#### PT5: Sales by Customer Segment
- Rows: Customer_Segment
- Values: Sum of Sales_Amount, Sum of Profit

#### PT6: Sales Channel Performance
- Rows: Sales_Channel
- Values: Sum of Sales_Amount, Sum of Profit, Count of Order_ID

#### PT7: Monthly Growth
- Rows: Month_Num
- Values: Sum of Sales_Amount
- Add Calculated Field: Growth = (Current - Previous) / Previous

---

## Step 2: Add Slicers

1. Click on any Pivot Table
2. Go to **PivotTable Analyze → Insert Slicer**
3. Select these fields:
   - ☑ Region
   - ☑ Category
   - ☑ Customer_Segment
   - ☑ Sales_Channel
   - ☑ Order_Status
4. Click OK
5. Position slicers on the Dashboard sheet

### Connect Slicers to Multiple Pivot Tables:
1. Right-click on a slicer
2. Select **Report Connections...**
3. Check ALL Pivot Tables that should respond to this slicer
4. Click OK

---

## Step 3: Add Timeline Slicer

1. Click on any Pivot Table
2. Go to **PivotTable Analyze → Insert Timeline**
3. Select **Order_Date**
4. Click OK
5. Set the timeline to show **Months**
6. Position on the Dashboard sheet

### Connect Timeline to Multiple Pivot Tables:
1. Right-click on the Timeline
2. Select **Report Connections...**
3. Check ALL relevant Pivot Tables
4. Click OK

---

## Step 4: Create Pivot Charts

For each Pivot Table:
1. Click on the Pivot Table
2. Go to **PivotTable Analyze → PivotChart**
3. Choose the appropriate chart type:
   - Monthly Trend → Line Chart
   - Sales by Region → Clustered Bar
   - Sales by Category → Horizontal Bar
   - Top 10 Products → Horizontal Bar
   - Profit by Region → Pie Chart
   - Customer Segment → Pie Chart
   - Sales Channel → Clustered Column

4. Move charts to the Dashboard sheet
5. Resize and align professionally

---

## Step 5: Dashboard Layout

Arrange on the Dashboard sheet:
```
Row 1:      Title Bar
Row 2:      Subtitle / Date Range
Row 4-5:    KPI Cards (already created)
Row 7-22:   Monthly Trend (left) | Sales by Region (right)
Row 23-38:  Sales by Category (left) | Top 10 Products (right)
Row 39-54:  Profit by Region (left) | Customer Segment (right)
Row 55-70:  Sales Channel Performance (full width)
```

Place Slicers in a sidebar or above the charts.

---

## Tips

- Use **Slicer Styles** (Slicer tab) to match the dashboard color scheme
- Set Slicer columns to 1 or 2 for compact layout
- Use **number formatting** on Pivot Table values (₹ format)
- Remove gridlines on the Dashboard sheet (View → uncheck Gridlines)
- Set a consistent font (Calibri recommended)
"""

    with open(os.path.join(DOCS_DIR, "pivot_slicer_setup_guide.md"), "w",
              encoding="utf-8") as f:
        f.write(guide)

    print(f"  ✓ Pivot & Slicer guide saved to docs/pivot_slicer_setup_guide.md")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    print("=" * 60)
    print("  SALES PERFORMANCE DASHBOARD — Project Generator")
    print("=" * 60)

    # Create directories
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(EXCEL_DIR, exist_ok=True)
    os.makedirs(DOCS_DIR, exist_ok=True)
    os.makedirs(os.path.join(BASE_DIR, "screenshots"), exist_ok=True)

    # Generate reference data
    print("\n📦 Generating reference data...")
    products = build_products()
    customers = build_customers(500)

    # Generate transactions
    print("\n📝 Generating transaction data...")
    monthly_data = generate_transactions(products, customers, YEAR, 10500)

    # Write CSV files
    print("\n💾 Writing CSV files...")
    total_rows = write_sales_csvs(monthly_data)
    write_customers_csv(customers)
    write_products_csv(products)

    # Clean & consolidate
    print("\n🧹 Cleaning & consolidating data...")
    clean_data = clean_consolidated_data(monthly_data, products, customers)
    print(f"  Clean rows after dedup + cleaning: {len(clean_data):,}")

    # Create Excel workbook
    print("\n📊 Building Excel workbook...")
    stats = create_workbook(clean_data, products, customers, monthly_data)

    # Write documentation
    print("\n📄 Writing documentation...")
    write_power_query_docs()
    write_pivot_slicer_guide()

    # Summary
    (total_sales, total_profit, profit_margin, total_orders,
     total_units, avg_order_value, region_sales, cat_sales,
     month_sales, seg_sales, chan_sales, top10_sales, top10_profit) = stats

    print("\n" + "=" * 60)
    print("  ✅ PROJECT GENERATION COMPLETE")
    print("=" * 60)
    print(f"\n  📊 Dataset Summary:")
    print(f"     Total Raw Rows:    {total_rows:>10,}")
    print(f"     Clean Rows:        {len(clean_data):>10,}")
    print(f"     Products:          {len(products):>10,}")
    print(f"     Customers:         {len(customers):>10,}")
    print(f"\n  💰 KPI Summary:")
    print(f"     Total Sales:       ₹{total_sales:>14,.0f}")
    print(f"     Total Profit:      ₹{total_profit:>14,.0f}")
    print(f"     Profit Margin:     {profit_margin:>14.1f}%")
    print(f"     Total Orders:      {total_orders:>14,}")
    print(f"     Total Units:       {total_units:>14,}")
    print(f"     Avg Order Value:   ₹{avg_order_value:>14,.0f}")

    print(f"\n  📁 Files Created:")
    print(f"     data/sales_jan.csv → sales_dec.csv  (12 files)")
    print(f"     data/customers.csv")
    print(f"     data/products.csv")
    print(f"     excel/Sales_Performance_Dashboard.xlsx")
    print(f"     docs/power_query_m_code.pq")
    print(f"     docs/pivot_slicer_setup_guide.md")

    return stats, clean_data, products, customers


if __name__ == "__main__":
    stats, clean_data, products, customers = main()
