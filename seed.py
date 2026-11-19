"""Makes Sunnyside Farmers Market's data, three files in three formats: our own made-up data. Ready to use.

Run it with: python seed.py

It writes, in this folder:
- sales.csv: one row for each product a vendor sold on a Saturday of the 2026 season (April 4 to October 31):
  date, vendor_id, product, category, units, price. A few units were never reported, so they're blank.
- vendors.xlsx: the vendor list, an Excel workbook with a sheet for each year, 2025 and 2026. Each sheet has two
  title rows above the table: vendor_id, vendor_name, stall, category, joined.
- reviews.json: the market app's export of its reviews: each vendor, with a nested list of its reviews (stars,
  comment, date).

It uses a fixed random seed, so every run, on every computer, writes exactly the same data. The tests make their own
copy the same way. Writing the Excel workbook needs openpyxl, which requirements.txt installs.
"""
import csv
import json
import random
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook

SEED = 2026
SALES_COLUMNS = ["date", "vendor_id", "product", "category", "units", "price"]
VENDOR_COLUMNS = ["vendor_id", "vendor_name", "stall", "category", "joined"]

FIRST_SATURDAY = date(2026, 4, 4)
SATURDAYS = 31          # April 4 to October 31, 2026

# vendor_id, vendor_name, stall, category, joined, products: (product, price, usual units, first month, last month)
VENDORS = [
    ("V01", "Green Acre Greens", "A1", "produce", date(2019, 5, 4),
     [("lettuce", 3.0, 30, 4, 10), ("kale", 3.5, 18, 4, 10), ("spinach", 4.0, 20, 4, 10),
      ("radishes", 2.5, 15, 5, 7), ("cucumbers", 1.5, 30, 6, 9), ("green beans", 4.0, 20, 6, 9),
      ("tomatoes", 4.0, 40, 7, 10), ("sweet corn", 0.75, 90, 7, 9)]),
    ("V02", "Bee Happy Honey", "B3", "honey", date(2020, 6, 6),
     [("honey jar", 12.0, 22, 4, 10), ("honeycomb", 15.0, 6, 4, 10), ("honey sticks", 0.5, 80, 4, 10),
      ("creamed honey", 10.0, 10, 4, 10), ("beeswax wrap", 8.0, 8, 4, 10)]),
    ("V03", "Rise & Shine Bakery", "C2", "baked goods", date(2018, 5, 5),
     [("sourdough loaf", 8.0, 35, 4, 10), ("rye loaf", 7.5, 15, 4, 10), ("cinnamon roll", 4.0, 60, 4, 10),
      ("croissant", 3.75, 50, 4, 10), ("blueberry muffin", 3.5, 40, 4, 10), ("scone", 3.25, 30, 4, 10)]),
    ("V04", "Petal Patch Flowers", "A4", "flowers", date(2021, 5, 1),
     [("tulip bunch", 10.0, 14, 5, 5), ("peony bunch", 20.0, 8, 5, 6), ("mixed bouquet", 18.0, 16, 4, 10),
      ("zinnia bunch", 9.0, 15, 7, 10), ("sunflower bunch", 12.0, 18, 7, 9), ("potted mum", 12.0, 10, 9, 10)]),
    ("V05", "Hollow Oak Woodworks", "D1", "crafts", date(2022, 5, 7),
     [("cutting board", 35.0, 4, 4, 10), ("wooden spoon", 9.0, 10, 4, 10), ("coaster set", 18.0, 6, 4, 10),
      ("serving tray", 45.0, 2, 4, 10), ("birdhouse", 30.0, 3, 4, 10)]),
    ("V06", "Sunny Slope Orchard", "A2", "produce", date(2017, 5, 6),
     [("strawberries", 6.0, 45, 5, 6), ("cherries", 7.0, 30, 6, 7), ("peaches", 4.5, 50, 7, 8),
      ("plums", 4.0, 30, 8, 9), ("pears", 3.5, 30, 8, 10), ("apples", 3.0, 70, 9, 10), ("cider", 7.0, 25, 9, 10),
      ("apple butter", 8.0, 12, 4, 10)]),
    ("V07", "Little Loaf Bakery", "C4", "baked goods", date(2023, 5, 6),
     [("baguette", 5.0, 40, 4, 10), ("focaccia", 7.0, 20, 4, 10), ("banana bread", 9.0, 12, 4, 10),
      ("cookie", 2.0, 60, 4, 10), ("fruit pie", 22.0, 8, 6, 10)]),
    ("V08", "Wildflower Stems", "A5", "flowers", date(2024, 5, 4),
     [("wildflower bunch", 12.0, 15, 5, 9), ("lavender bundle", 8.0, 15, 6, 8), ("dahlia bunch", 15.0, 12, 8, 10),
      ("dried flower wreath", 35.0, 3, 4, 10)]),
    ("V09", "Riverbend Farm", "B1", "produce", date(2018, 6, 2),
     [("eggs (dozen)", 6.0, 30, 4, 10), ("onions", 2.0, 30, 4, 10), ("carrots", 3.0, 30, 4, 10),
      ("beets", 3.0, 18, 6, 10), ("potatoes", 2.5, 40, 6, 10), ("zucchini", 2.0, 35, 6, 9),
      ("winter squash", 4.0, 25, 9, 10)]),
    ("V10", "Clay & Kiln Pottery", "D3", "crafts", date(2021, 6, 5),
     [("mug", 28.0, 6, 4, 10), ("bowl", 32.0, 4, 4, 10), ("plate", 30.0, 4, 4, 10), ("vase", 40.0, 3, 4, 10),
      ("spoon rest", 12.0, 6, 4, 10)]),
    ("V11", "Golden Comb Apiary", "B4", "honey", date(2025, 5, 3),
     [("honey jar", 11.0, 18, 4, 10), ("honey sticks", 0.5, 60, 4, 10), ("beeswax candle", 14.0, 9, 4, 10),
      ("lip balm", 4.0, 20, 4, 10)]),
    ("V12", "Maple Lane Market Garden", "B2", "produce", date(2026, 4, 4),
     [("salad greens", 4.0, 25, 4, 10), ("microgreens", 5.0, 18, 4, 10), ("herb bundle", 3.0, 20, 4, 10),
      ("basil plant", 6.0, 10, 5, 7), ("cherry tomatoes", 5.0, 30, 7, 10), ("peppers", 1.5, 40, 7, 10)]),
    ("V14", "Stone Hearth Bread", "C1", "baked goods", date(2024, 4, 20),
     [("whole wheat loaf", 7.0, 20, 4, 10), ("pretzel", 3.0, 40, 4, 10), ("bagel", 2.5, 45, 4, 10),
      ("pumpkin bread", 9.0, 12, 9, 10)]),
    ("V15", "Hillside Herb Farm", "B5", "produce", date(2022, 4, 23),
     [("herb plant", 5.0, 15, 4, 7), ("garlic", 1.5, 40, 7, 10), ("shallots", 3.0, 15, 7, 10),
      ("dried herbs", 6.0, 12, 4, 10), ("mint bundle", 2.5, 18, 5, 9)]),
    ("V16", "Knot & Needle", "D4", "crafts", date(2025, 4, 19),
     [("knit hat", 25.0, 4, 4, 10), ("scarf", 30.0, 3, 9, 10), ("tote bag", 15.0, 6, 4, 10),
      ("dish cloth set", 10.0, 8, 4, 10)]),
]

# The 2025 season's list: the vendors who had joined by then, plus one who has since left.
LEFT_IN_2025 = ("V13", "Old Mill Candles", "D2", "crafts", date(2020, 5, 2))

COMMENTS = {
    5: ["Best at the market!", "Worth the early trip.", "Always fresh and friendly.", "We come back every week.",
        "Ask what's new this week.", "Perfect for a picnic."],
    4: ["Very good, a little pricey.", "Great, but sold out by ten.", "Lovely, the line was long.",
        "Good choice for a gift."],
    3: ["Fine, nothing special.", "Good, but smaller than last year.", "Okay, the stall was hard to find."],
    2: ["Sold out when we got there.", "Not as fresh as we hoped."],
    1: ["Closed early without a sign."],
}


def saturdays():
    return [FIRST_SATURDAY + timedelta(weeks=week) for week in range(SATURDAYS)]


def make_sales(rng):
    """Every row of sales.csv, in date order, then in vendor order."""
    rows = []
    for day in saturdays():
        for vendor_id, _, _, category, _, products in VENDORS:
            if rng.random() < 0.03:   # a vendor skips about one Saturday in thirty
                continue
            for product, price, usual, first_month, last_month in products:
                if not first_month <= day.month <= last_month:
                    continue
                units = max(1, round(rng.gauss(usual, usual * 0.25)))
                rows.append({"date": day.isoformat(), "vendor_id": vendor_id, "product": product,
                             "category": category, "units": str(units), "price": f"{price:.2f}"})
    # A few vendors forgot to report how many they sold: the units stay blank.
    for row in rng.sample(rows, 12):
        row["units"] = ""
    return rows


def make_reviews(rng):
    """The app's export: every vendor, with a nested list of its reviews. A vendor new this year has none yet."""
    days = saturdays()
    export = []
    for vendor_id, vendor_name, _, _, joined, _ in VENDORS:
        count = 0 if joined.year == 2026 else rng.randint(4, 11)
        reviews = []
        for _ in range(count):
            stars = rng.choices([5, 4, 3, 2, 1], weights=[45, 30, 14, 7, 4])[0]
            reviews.append({"stars": stars, "comment": rng.choice(COMMENTS[stars]),
                            "date": rng.choice(days).isoformat()})
        reviews.sort(key=lambda review: review["date"])
        export.append({"vendor_id": vendor_id, "vendor_name": vendor_name, "reviews": reviews})
    return export


def write_sales(path):
    """Writes sales.csv to path (a file name or a Path) and returns the Path."""
    path = Path(path)
    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=SALES_COLUMNS)
        writer.writeheader()
        writer.writerows(make_sales(random.Random(SEED)))
    return path


def write_vendors(path):
    """Writes vendors.xlsx to path: a sheet for 2025 and one for 2026, each with two title rows. Returns the Path."""
    path = Path(path)
    book = Workbook()
    book.remove(book.active)
    seasons = {
        2025: [v[:5] for v in VENDORS if v[4].year <= 2025] + [LEFT_IN_2025],
        2026: [v[:5] for v in VENDORS],
    }
    for year, vendors in seasons.items():
        sheet = book.create_sheet(str(year))
        sheet.append([f"Sunnyside Farmers Market: vendors for the {year} season"])
        sheet.append(["Saturdays, 8 a.m. to 1 p.m., Sunnyside Park"])
        sheet.append(VENDOR_COLUMNS)
        for vendor in sorted(vendors):
            sheet.append(list(vendor))
        for row in sheet.iter_rows(min_row=4, min_col=5, max_col=5):
            row[0].number_format = "yyyy-mm-dd"
    book.save(path)
    return path


def write_reviews(path):
    """Writes reviews.json to path and returns the Path."""
    path = Path(path)
    with path.open("w", encoding="utf-8") as file:
        json.dump(make_reviews(random.Random(SEED + 1)), file, indent=2)
        file.write("\n")
    return path


def write_all(folder):
    """Writes all three files into folder and returns their Paths."""
    folder = Path(folder)
    return (write_sales(folder / "sales.csv"), write_vendors(folder / "vendors.xlsx"),
            write_reviews(folder / "reviews.json"))


if __name__ == "__main__":
    here = Path(__file__).parent
    write_all(here)
    rows = len(make_sales(random.Random(SEED)))
    reviews = sum(len(vendor["reviews"]) for vendor in make_reviews(random.Random(SEED + 1)))
    print(f"Wrote sales.csv: {rows} rows.")
    print("Wrote vendors.xlsx: sheets 2025 and 2026.")
    print(f"Wrote reviews.json: {len(VENDORS)} vendors, {reviews} reviews.")
