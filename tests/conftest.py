# Shared set-up for the tests. You don't need to change this file.
# Each test that asks for market_files gets its own fresh copy of the market's three files, made by seed.py in a
# temporary folder, so the tests check your code against the same data even if you changed your own copies.
import pytest

from seed import write_all


@pytest.fixture
def market_files(tmp_path):
    """The paths of a freshly seeded sales.csv, vendors.xlsx and reviews.json in a temporary folder."""
    sales, vendors, reviews = write_all(tmp_path)
    return {"sales": sales, "vendors": vendors, "reviews": reviews}
