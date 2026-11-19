# Exercise: Farmers Market First Look

Exercise code for the lesson **Exercise: Farmers Market First Look** in the Python for Data Analysis course on [ask-codey.com](https://www.ask-codey.com).

Sunnyside Farmers Market's data arrives three ways: a season of sales as a CSV file, the vendor list as an Excel workbook with a sheet for each year, and the market app's reviews as nested JSON. You'll write five functions in `market.py` that load each file into a pandas table and sum a table up, and tests call each one to check what it returns.

## What's here

- `seed.py`: makes the market's three files, our own made-up data: `sales.csv` (1,900 rows), `vendors.xlsx` (sheets `2025` and `2026`) and `reviews.json` (15 vendors, 104 reviews). It uses a fixed random seed, so it writes the same data on every computer. Ready to use.
- `market.py`: the five functions, `load_sales`, `load_vendors`, `load_reviews`, `table_summary` and `top_categories`. Each one stops with `NotImplementedError: Not built yet.` for now. The docstring under each says what it should do, with examples. For you to finish.
- `explore.ipynb`: a notebook for looking at the three files and trying your functions as you write them. The tests don't read it. Ready to use, and yours to change.
- `report.py`: loads the three files with your five functions and prints what's in them, so you can see them at work. Ready to use.
- `tests/test_market.py`: tests that call your functions and check what comes back. You don't need to change this file.
- `tests/conftest.py`: gives each test a fresh copy of the three files, made by `seed.py` in a temporary folder, so the tests use the same data even if you changed your own copies. Ready to use.
- `requirements.txt`: the packages this exercise needs, each at an exact version: pandas, openpyxl (which pandas uses to read Excel files), ipykernel (so VS Code can run the notebook) and pytest. Ready to use.
- `pytest.ini`: tells pytest where the tests are and where to find `market.py` and `seed.py`. Ready to use.
- `.gitignore`: tells Git to leave out the virtual environment, the files Python and pytest make as they run, and the three data files, which `seed.py` makes again whenever they're needed. Ready to use.

## Before you start

The packages go in a virtual environment: a private set of packages for this folder only. In VS Code's terminal, in this folder, run these three commands.

On Windows (PowerShell):

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

On a Mac (zsh):

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

- The first command makes the virtual environment, in a folder named `.venv`.
- The second turns it on for this terminal: the prompt starts with `(.venv)` from then on.
- The third installs the packages `requirements.txt` lists into it.

*Virtual Environments and pip* explains all three. The `.venv` folder is in `.gitignore`, so Git leaves it out: these commands make it again on any computer.

If PowerShell answers that running scripts is disabled on this system, run this once, then run the second command again:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Each new terminal starts with the virtual environment off. If the prompt doesn't start with `(.venv)`, run the second command again. (VS Code often turns it on for you.)

Then make the data, once:

```bash
python seed.py
```

It prints three lines, one for each file it wrote. To look at the data, open `explore.ipynb`, choose the `.venv` kernel with **Select Kernel** (top right), and run the cells.

## Check your work

```bash
python -m pytest
```

At first every test fails: that is expected, because each function stops with a `NotImplementedError` for now. Each function you finish turns its tests green. A failing test's message says what it called, what it expected and what came back, for example *load_vendors(path, 2026).shape should give (15, 5), but it gave (17, 5).*

To see your functions at work on the whole market, run:

```bash
python report.py
```

## The rules

- Keep each function's name and parameters as they are.
- Load each file from the path the function is given, never from a name of your own.
- Don't change the table `table_summary` or `top_categories` is given.
- A sheet of the workbook is named by its year, as text: `'2026'`. A year given as the number `2026` still means the sheet named `2026`.
- Don't change the tests, `conftest.py`, `seed.py` or `report.py`.

## Tips

- Try each step in `explore.ipynb` first, on the real files, then move the lines that work into `market.py`.
- Look at the workbook's column names before you choose `skiprows`: names such as `Unnamed: 1` mean there are title rows above the table.
- `pd.json_normalize` takes Python data, not a file name, so read the file with `json.load` first.
- `describe()` sums up only the number columns, so its columns are the number columns.
- `len(df) - df.count()` gives the missing values in each column, and a Series' `to_dict()` turns it into a dictionary.
- After you change `market.py`, restart the notebook's kernel before you try it there, so the notebook imports your newest code.

## Hand it in

When it works, push your solution to a public repo of your own on GitHub, then paste that repo's address into **Submit your exercise** on the lesson page at [ask-codey.com](https://www.ask-codey.com). That's what completes the lesson. The lesson page shows the Git commands, step by step.
