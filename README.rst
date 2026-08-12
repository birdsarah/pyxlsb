pyxlsb
======

|PyPI|

``pyxlsb`` is an Excel 2007-2010 Binary Workbook (xlsb) parser for
Python. The library is currently extremely limited, but functional
enough for basic data extraction.

Install
-------

.. code:: sh

   pip install pyxlsb

Usage
-----

The module exposes an ``open_workbook(name)`` method (similar to Xlrd
and OpenPyXl) for opening XLSB files. The Workbook object representing
the file is returned.

.. code:: python

   from pyxlsb import open_workbook
   with open_workbook('Book1.xlsb') as wb:
       # Do stuff with wb

The Workbook object exposes a ``get_sheet(idx)`` method for retrieving a
Worksheet instance.

.. code:: python

   # Using the sheet index (1-based)
   with wb.get_sheet(1) as sheet:
       # Do stuff with sheet

   # Using the sheet name
   with wb.get_sheet('Sheet1') as sheet:
       # Do stuff with sheet

Tip: A ``sheets`` property containing the sheet names is available on
the Workbook instance.

The ``rows()`` method will hand out an iterator to read the worksheet
rows.

.. code:: python

   # You can use .rows(sparse=True) to skip empty rows
   for row in sheet.rows():
       print(row)
   # [Cell(r=0, c=0, v='TEXT', f=None, error=None),
   #  Cell(r=0, c=1, v=42.1337, f='SUM(A1:A5)', error=None)]

Formulas
--------

Each cell carries its formula in ``f``, as the text Excel would show in
the formula bar without the leading ``=``. A cell that holds no formula
has ``f`` set to ``None``.

.. code:: python

   for row in sheet.rows():
       for cell in row:
           if cell.f is not None:
               print(cell.r, cell.c, '=' + cell.f, '->', cell.v)
   # 1 1 =SUM(Data!A1:A5)*Rate -> 1234.5

Shared and array formulas are resolved for you: every cell of the block
reports its own text, with a shared formula's relative references shifted
to the cell it lands in.

Unreadable formulas
~~~~~~~~~~~~~~~~~~~

By default a formula that cannot be parsed raises ``FormulaError``, naming
the offending cell. Pass ``formula_errors='store'`` to keep reading
instead, which records the failure on the cell it happened in:

.. code:: python

   with open_workbook('Book1.xlsb', formula_errors='store') as wb:
       with wb.get_sheet(1) as sheet:
           for row in sheet.rows():
               for cell in row:
                   if cell.error is not None:
                       print('could not read', cell.r, cell.c, cell.error)

Either way, ``f`` only ever holds a real formula. A cell whose formula
could not be read has ``f=None`` **and** ``error`` set, so the test for a
genuinely formula-free cell stays ``cell.f is None and cell.error is
None``, and a broken formula is never quietly reported as a hardcoded
value. The cached value in ``v`` is unaffected either way.

Parsing costs time on formula-heavy workbooks, and reading a whole
workbook's formulas requires the defined-name and external-reference
tables. Pass ``parse_formulas=False`` to skip all of it, in which case
``f`` is always ``None``:

.. code:: python

   with open_workbook('Book1.xlsb', parse_formulas=False) as wb:
       # Values only, as in pyxlsb 1.0.11 and earlier

The whitespace an author typed inside a formula is preserved, so the text
round-trips byte for byte against what Excel itself writes.

Known gaps: references into *other workbooks* render with a numeric
``[n]`` book prefix rather than the linked file name; and where a legacy
formula relies on implicit intersection, Excel adds an explicit
``_xlfn.SINGLE(...)`` wrapper when it exports to ``.xlsx`` that is not
present in the ``.xlsb`` and so is not reported here.

As a sense of scale, parsing adds roughly 7 µs per formula, so a workbook
containing a million of them costs a few tens of seconds on top of the
time to read its values.

Do note that dates will appear as floats. You must use the
``convert_date(date)`` method from the ``pyxlsb`` module to turn them
into ``datetime`` instances.

.. code:: python

   from pyxlsb import convert_date
   print(convert_date(41235.45578))
   # datetime.datetime(2012, 11, 22, 10, 56, 19)

.. |PyPI| image:: https://img.shields.io/pypi/v/pyxlsb.svg
   :target: https://pypi.python.org/pypi/pyxlsb
