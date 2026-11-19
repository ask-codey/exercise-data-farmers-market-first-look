# Loads the market's three files with your five functions from market.py and prints what's in them, so you can see
# them at work. Ready to use. Make the data first with: python seed.py
# Run it with: python report.py
from market import load_reviews, load_sales, load_vendors, table_summary, top_categories

sales = load_sales("sales.csv")
summary = table_summary(sales)
print("1) sales rows:", summary["rows"])
print("2) sales columns:", summary["columns"])
print("3) number columns:", ", ".join(summary["numeric"]))
print("4) units missing:", summary["missing"]["units"])

top = top_categories(sales, 3)
print("5) top categories:", ", ".join(f"{name} ({count})" for name, count in top.items()))

print("6) vendors in 2026:", len(load_vendors("vendors.xlsx", 2026)))
print("7) vendors in 2025:", len(load_vendors("vendors.xlsx", 2025)))

reviews = load_reviews("reviews.json")
print("8) reviews:", len(reviews))
print("9) vendors with reviews:", reviews["vendor_id"].nunique())
print("10) average stars:", round(reviews["stars"].mean(), 2))
