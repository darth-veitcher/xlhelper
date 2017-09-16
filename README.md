# xlhelper
Excel Sheet to Python Dict converter using openpyxl

XL helper (`xlhelper`) is a small python module built to import a defined Excel
worksheet and convert the contained data into `Ordered Dictionaries` with the
keys and values corresponding to a defined header row and the values of the
rows.

This was heavily influenced by some work I did a few years ago and has since
been updated to use [generators](https://wiki.python.org/moin/Generators)
in order to remain as efficient as possible with memory usage when reading
large files.

Whilst the option remains to use something like
[pandas](https://pandas.pydata.org) I was after a minimal installation and
limited overhead.


Example:

| ID       | Product Name        | Modifier |
|----------|:--------------------|----------|
| 1        | Whizbang 5000       | Instant fame and fortune
| 2a       | Recursive Slingshot | +5 annoyance

Produces:

```python
>>> import os
>>> import xlhelper
>>> from pprint import pprint

>>> xlpath = os.path.join(os.getcwd(),
                    'examples',
                    'product listing.xlsx')

# Returns a standard dict (ordered alphanumerically)
>>> for i in xlhelper.sheet_to_dict(xlpath):
...   pprint(i)

{'ID': 1,
 'Modifier': 'Instant fame and fortune',
 'Product Name': 'Whizbang 5000'}
{'ID': '2a', 'Modifier': '+5 annoyance', 'Product Name': 'Recursive Slingshot'}

# Return an OrderedDict with original ordering from Excel file of columns
# maintained.
>>> pprint(xlhelper.sheet_to_dict(xlpath, keep_order=True))

OrderedDict([('ID', 1),
             ('Product Name', 'Whizbang 5000'),
             ('Modifier', 'Instant fame and fortune')])
OrderedDict([('ID', '2a'),
             ('Product Name', 'Recursive Slingshot'),
             ('Modifier', '+5 annoyance')])

```

## Installation
From github:
```bash
git clone https://github.com/darth-veitcher/xlhelper.git
cd xlhelper
pip install .
```
