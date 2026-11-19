"""Sunnyside Farmers Market: five functions for you to finish, to load the market's three files and sum them up.

The docstring under each function says what it should do, with examples. Replace each
raise NotImplementedError("Not built yet.") line with your own code.

The three load_ functions are given a path and return a DataFrame. table_summary and top_categories are given a
DataFrame, and the table they're given must stay exactly as it was.
"""
import json

import pandas as pd


def load_sales(path):
    """Read the sales CSV file at path into a DataFrame, with every row and column of the file.

    On the file seed.py makes: 1900 rows and 6 columns, date, vendor_id, product, category, units and price.
    """
    raise NotImplementedError("Not built yet.")


def load_vendors(path, year):
    """Read one season's vendor list from the Excel workbook at path into a DataFrame.

    The workbook has a sheet for each season, named by its year: '2025' and '2026'. year may be given as a number,
    such as 2026, or as text, such as '2026'; either way, read the sheet with that name. Each sheet has two title rows
    above the table, so the third row holds the column names: vendor_id, vendor_name, stall, category and joined.

    load_vendors(path, 2026) on the workbook seed.py makes: 15 rows and 5 columns, starting with V01, Green Acre Greens.
    """
    raise NotImplementedError("Not built yet.")


def load_reviews(path):
    """Read the app's reviews export, a JSON file at path, into a DataFrame with one row for each review.

    The file is a list of vendors, each with its vendor_id, vendor_name and a nested list, reviews, in which each
    review has stars, comment and date. Return one row for each review, with the columns stars, comment, date and
    vendor_id (the vendor_id of the vendor it belongs to). A vendor with no reviews has no rows.

    On the file seed.py makes: 104 rows.
    """
    raise NotImplementedError("Not built yet.")


def table_summary(df):
    """Return a dictionary that sums up the DataFrame df, with four keys:

    'rows': how many rows it has
    'columns': how many columns it has
    'numeric': a list of the names of its number columns, in the table's order
    'missing': a dictionary of each column's name and how many of its values are missing

    For a table with a product column and a units column, 3 rows, one of the units missing:
    {'rows': 3, 'columns': 2, 'numeric': ['units'], 'missing': {'product': 0, 'units': 1}}
    """
    raise NotImplementedError("Not built yet.")


def top_categories(sales, n):
    """Return a Series of the n categories that appear in the most rows of sales, with how many rows each has,
    most first: the category names are the index and the counts are the values.

    For a category column holding produce, honey, produce, crafts, produce and honey, top_categories(sales, 2)
    gives produce 3 and honey 2.
    """
    raise NotImplementedError("Not built yet.")
