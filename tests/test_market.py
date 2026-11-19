# Tests for market.py. Run them with: python -m pytest (in the exercise folder, with the virtual environment active).
# Each test calls one of your functions, on the files seed.py makes (in a temporary folder) or on a small file or
# table written in the test, and checks what comes back. A failing test says what it called, what it expected and
# what came back. You don't need to change this file.
import json
import math

import pandas as pd
from openpyxl import Workbook

from market import load_reviews, load_sales, load_vendors, table_summary, top_categories


def check(call, got, expected):
    """Fails with a message that says what the call returned and what it should return."""
    assert got == expected, f"{call} should give {expected!r}, but it gave {got!r}."


def values(column):
    """A column's values as a plain list, with each missing value shown as None."""
    return [None if isinstance(v, float) and math.isnan(v) else v for v in column.tolist()]


def counted(series):
    """A Series of counts as a list of (label, count) pairs, in its order."""
    return list(zip(series.index.tolist(), series.tolist()))


# load_sales

def test_1_load_sales_reads_the_file_as_1900_rows_and_6_columns(market_files):
    check("load_sales(path).shape", load_sales(market_files["sales"]).shape, (1900, 6))


def test_2_load_sales_keeps_the_columns_date_vendor_id_product_category_units_and_price(market_files):
    check("list(load_sales(path).columns)", list(load_sales(market_files["sales"]).columns),
          ["date", "vendor_id", "product", "category", "units", "price"])


def test_3_load_sales_reads_units_as_numbers_with_12_missing(market_files):
    sales = load_sales(market_files["sales"])
    check("load_sales(path)['units'].dtype", str(sales["units"].dtype), "float64")
    check("len(sales) - sales['units'].count()", int(len(sales) - sales["units"].count()), 12)


def test_4_load_sales_reads_the_file_it_is_given_2_rows_of_honey_jars(tmp_path):
    path = tmp_path / "two_rows.csv"
    path.write_text(
        "date,vendor_id,product,category,units,price\n"
        "2026-04-04,V02,honey jar,honey,20,12.00\n"
        "2026-04-04,V11,honey jar,honey,15,11.00\n",
        encoding="utf-8",
    )
    sales = load_sales(path)
    check("load_sales(path)['vendor_id'] for a file of 2 rows", values(sales["vendor_id"]), ["V02", "V11"])
    check("load_sales(path)['units'] for a file of 2 rows", values(sales["units"]), [20, 15])


# load_vendors

def test_5_load_vendors_2026_reads_15_rows_and_5_columns(market_files):
    check("load_vendors(path, 2026).shape", load_vendors(market_files["vendors"], 2026).shape, (15, 5))


def test_6_load_vendors_2026_skips_the_title_rows_and_starts_with_v01_green_acre_greens(market_files):
    vendors = load_vendors(market_files["vendors"], 2026)
    check("list(load_vendors(path, 2026).columns)", list(vendors.columns),
          ["vendor_id", "vendor_name", "stall", "category", "joined"])
    check("load_vendors(path, 2026)['vendor_name'][0]", vendors["vendor_name"][0], "Green Acre Greens")


def test_7_load_vendors_2025_has_old_mill_candles_v13_and_not_maple_lane_v12(market_files):
    ids = values(load_vendors(market_files["vendors"], 2025)["vendor_id"])
    check("'V13' in load_vendors(path, 2025)['vendor_id']", "V13" in ids, True)
    check("'V12' in load_vendors(path, 2025)['vendor_id']", "V12" in ids, False)


def test_8_load_vendors_takes_the_year_as_text_2026_too(market_files):
    check("load_vendors(path, '2026').shape", load_vendors(market_files["vendors"], "2026").shape, (15, 5))


def test_9_load_vendors_2025_reads_the_sheet_named_2025_when_it_is_the_second_tab(tmp_path):
    path = tmp_path / "two_tabs.xlsx"
    book = Workbook()
    book.remove(book.active)
    for year, vendor in [(2026, ["V12", "Maple Lane Market Garden", "B2", "produce"]),
                         (2025, ["V13", "Old Mill Candles", "D2", "crafts"])]:
        sheet = book.create_sheet(str(year))
        sheet.append([f"Vendors for the {year} season"])
        sheet.append(["Saturdays, 8 a.m. to 1 p.m."])
        sheet.append(["vendor_id", "vendor_name", "stall", "category"])
        sheet.append(vendor)
    book.save(path)
    check("load_vendors(path, 2025)['vendor_name'] for a workbook whose tabs are 2026, then 2025",
          values(load_vendors(path, 2025)["vendor_name"]), ["Old Mill Candles"])


# load_reviews

def test_10_load_reviews_gives_one_row_for_each_of_the_104_reviews(market_files):
    check("len(load_reviews(path))", len(load_reviews(market_files["reviews"])), 104)


def test_11_load_reviews_has_the_columns_stars_comment_date_and_vendor_id(market_files):
    check("sorted(load_reviews(path).columns)", sorted(load_reviews(market_files["reviews"]).columns),
          ["comment", "date", "stars", "vendor_id"])


def test_12_load_reviews_turns_v01_with_2_reviews_and_v02_with_1_into_3_rows(tmp_path):
    path = tmp_path / "three_reviews.json"
    path.write_text(json.dumps([
        {"vendor_id": "V01", "vendor_name": "Green Acre Greens", "reviews": [
            {"stars": 5, "comment": "Best at the market!", "date": "2026-05-09"},
            {"stars": 3, "comment": "Fine, nothing special.", "date": "2026-06-13"}]},
        {"vendor_id": "V02", "vendor_name": "Bee Happy Honey", "reviews": [
            {"stars": 4, "comment": "Good choice for a gift.", "date": "2026-07-11"}]},
        {"vendor_id": "V12", "vendor_name": "Maple Lane Market Garden", "reviews": []},
    ]), encoding="utf-8")
    reviews = load_reviews(path)
    check("load_reviews(path)['vendor_id'] for V01 with 2 reviews, V02 with 1 and V12 with none",
          values(reviews["vendor_id"]), ["V01", "V01", "V02"])
    check("load_reviews(path)['stars'] for V01 with 2 reviews, V02 with 1 and V12 with none",
          values(reviews["stars"]), [5, 3, 4])


def test_13_load_reviews_has_no_rows_for_maple_lane_v12_which_has_no_reviews(market_files):
    ids = values(load_reviews(market_files["reviews"])["vendor_id"])
    check("'V12' in load_reviews(path)['vendor_id']", "V12" in ids, False)
    check("number of different vendor_ids in load_reviews(path)", len(set(ids)), 14)


# table_summary

def test_14_table_summary_of_3_rows_with_1_of_the_units_missing():
    table = pd.DataFrame({"product": ["lettuce", "kale", "spinach"], "units": [21, None, 18]})
    check("table_summary(df) for product and units, 3 rows, 1 unit missing", table_summary(table),
          {"rows": 3, "columns": 2, "numeric": ["units"], "missing": {"product": 0, "units": 1}})


def test_15_table_summary_of_the_sales_gives_1900_rows_6_columns_and_numeric_units_and_price(market_files):
    summary = table_summary(load_sales(market_files["sales"]))
    check("table_summary(sales)['rows']", summary["rows"], 1900)
    check("table_summary(sales)['columns']", summary["columns"], 6)
    check("table_summary(sales)['numeric']", summary["numeric"], ["units", "price"])


def test_16_table_summary_of_the_sales_counts_12_missing_units_and_none_elsewhere(market_files):
    summary = table_summary(load_sales(market_files["sales"]))
    check("table_summary(sales)['missing']", summary["missing"],
          {"date": 0, "vendor_id": 0, "product": 0, "category": 0, "units": 12, "price": 0})


def test_17_table_summary_leaves_the_table_it_was_given_unchanged():
    table = pd.DataFrame({"product": ["lettuce", "kale"], "units": [21, None]})
    table_summary(table)
    check("the table's columns after table_summary(df)", list(table.columns), ["product", "units"])
    check("the table's units after table_summary(df)", values(table["units"]), [21.0, None])


# top_categories

def test_18_top_categories_2_of_produce_honey_produce_crafts_produce_honey_gives_produce_3_and_honey_2():
    sales = pd.DataFrame({"category": ["produce", "honey", "produce", "crafts", "produce", "honey"]})
    check("top_categories(sales, 2) as (category, count) pairs", counted(top_categories(sales, 2)),
          [("produce", 3), ("honey", 2)])


def test_19_top_categories_3_of_the_sales_gives_produce_654_baked_goods_415_and_crafts_399(market_files):
    top = top_categories(load_sales(market_files["sales"]), 3)
    check("top_categories(sales, 3) as (category, count) pairs", counted(top),
          [("produce", 654), ("baked goods", 415), ("crafts", 399)])
