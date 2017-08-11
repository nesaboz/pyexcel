Sheet
==========


Random access
-----------------

To randomly access a cell of :class:`~pyexcel.Sheet` instance, two
syntax are available::

    sheet[row, column]

or::

    sheet['A1']

The former syntax is handy when you know the row and column numbers.
The latter syntax is introduced to help you convert the excel column header
such as "AX" to integer numbers.

Suppose you have the following data, you can get value 5 by reader[2, 2].

.. pyexcel-table::

   ---pyexcel:example data---
   Example,X,Y,Z
   a,1,2,3
   b,4,5,6
   c,7,8,9


Here is the example code showing how you can randomly access a cell:

.. testcode::
   :hide:

   >>> data = [
   ...     ['Example', 'X', 'Y', 'Z'],
   ...     ['a', 1, 2, 3],
   ...     ['b', 4, 5, 6],
   ...     ['c', 7, 8, 9]]
   >>> s = pyexcel.Sheet(data)
   >>> s.save_as("example.xls")

.. testcode::

   >>> sheet = pyexcel.get_sheet(file_name="example.xls")
   >>> sheet.content
   +---------+---+---+---+
   | Example | X | Y | Z |
   +---------+---+---+---+
   | a       | 1 | 2 | 3 |
   +---------+---+---+---+
   | b       | 4 | 5 | 6 |
   +---------+---+---+---+
   | c       | 7 | 8 | 9 |
   +---------+---+---+---+
   >>> print(sheet[2, 2])
   5
   >>> print(sheet["C3"])
   5
   >>> sheet[3, 3] = 10
   >>> print(sheet[3, 3])
   10


.. note::

   In order to set a value to a cell, please use
   sheet[row_index, column_index] = new_value


**Random access to rows and columns**

.. testcode::
   :hide:

   >>> sheet[1, 0] = str(sheet[1, 0])
   >>> str(sheet[1,0])
   'a'
   >>> sheet[0, 2] = str(sheet[0, 2])
   >>> sheet[0, 2]
   'Y'

Continue with previous excel file, you can access
row and column separately::

    >>> sheet.row[1]
    ['a', 1, 2, 3]
    >>> sheet.column[2]
    ['Y', 2, 5, 8]


**Use custom names instead of index**
Alternatively, it is possible to use the first row to
refer to each columns::

    >>> sheet.name_columns_by_row(0)
    >>> print(sheet[1, "Y"])
    5
    >>> sheet[1, "Y"] = 100
    >>> print(sheet[1, "Y"])
    100

You have noticed the row index has been changed. It is because
first row is taken as the column names, hence all rows after
the first row are shifted. Now accessing the columns are
changed too::

    >>> sheet.column['Y']
    [2, 100, 8]

Hence access the same cell, this statement also works::

    >>> sheet.column['Y'][1]
    100

Further more, it is possible to use first column to refer to each rows::

    >>> sheet.name_rows_by_column(0)

To access the same cell, we can use this line::

    >>> sheet.row["b"][1]
    100

For the same reason, the row index has been reduced by 1. Since we
have named columns and rows, it is possible to access the same cell
like this::

    >>> print(sheet["b", "Y"])
    100
    >>> sheet["b", "Y"] = 200
    >>> print(sheet["b", "Y"])
    200


**Play with data**

Suppose you have the following data in any of the supported
excel formats again:

.. pyexcel-table::

   ---pyexcel:data with columns---
   Column 1,Column 2,Column 3
   1,4,7
   2,5,8
   3,6,9

.. testcode::

   >>> sheet = pyexcel.get_sheet(file_name="example_series.xls",
   ...      name_columns_by_row=0)

.. testcode::
   :hide:

   >>> sheet.colnames = [ str(name) for name in sheet.colnames]

You can get headers::

    >>> print(list(sheet.colnames))
    ['Column 1', 'Column 2', 'Column 3']

You can use a utility function to get all in a dictionary::

    >>> sheet.to_dict()
    OrderedDict([('Column 1', [1, 4, 7]), ('Column 2', [2, 5, 8]), ('Column 3', [3, 6, 9])])

Maybe you want to get only the data without the column headers.
You can call :meth:`~pyexcel.Sheet.rows()` instead::

    >>> list(sheet.rows())
    [[1, 2, 3], [4, 5, 6], [7, 8, 9]]

You can get data from the bottom to the top one by
 calling :meth:`~pyexcel.Sheet.rrows()`::

    >>> list(sheet.rrows())
    [[7, 8, 9], [4, 5, 6], [1, 2, 3]]

You might want the data arranged vertically. You can call
:meth:`~pyexcel.Sheet.columns()`::

    >>> list(sheet.columns())
    [[1, 4, 7], [2, 5, 8], [3, 6, 9]]

You can get columns in reverse sequence as well by calling
:meth:`~pyexcel.Sheet.rcolumns()`::

    >>> list(sheet.rcolumns())
    [[3, 6, 9], [2, 5, 8], [1, 4, 7]]

Do you want to flatten the data? You can get the content in one
dimensional array. If you are interested in playing with one
dimensional enumeration, you can check out these functions
:meth:`~pyexcel.Sheet.enumerate`, :meth:`~pyexcel.Sheet.reverse`,
:meth:`~pyexcel.Sheet.vertical`, and :meth:`~pyexcel.Sheet.rvertical()`::

    >>> list(sheet.enumerate())
    [1, 2, 3, 4, 5, 6, 7, 8, 9]
    >>> list(sheet.reverse())
    [9, 8, 7, 6, 5, 4, 3, 2, 1]
    >>> list(sheet.vertical())
    [1, 4, 7, 2, 5, 8, 3, 6, 9]
    >>> list(sheet.rvertical())
    [9, 6, 3, 8, 5, 2, 7, 4, 1]


**attributes**

Attributes::

    >>> import pyexcel
    >>> content = "1,2,3\n3,4,5"
    >>> sheet = pyexcel.get_sheet(file_type="csv", file_content=content)
    >>> sheet.tsv
    '1\t2\t3\r\n3\t4\t5\r\n'
    >>> print(sheet.simple)
    csv:
    -  -  -
    1  2  3
    3  4  5
    -  -  -

What's more, you could as well set value to an attribute, for example::
    >>> import pyexcel
    >>> content = "1,2,3\n3,4,5"
    >>> sheet = pyexcel.Sheet()
    >>> sheet.csv = content
    >>> sheet.array
    [[1, 2, 3], [3, 4, 5]]

You can get the direct access to underneath stream object. In some situation,
it is desired::

    >>> stream = sheet.stream.tsv

The returned stream object has tsv formatted content for reading.


What you could further do is to set a memory stream of any supported file format
to a sheet. For example:

    >>> another_sheet = pyexcel.Sheet()
    >>> another_sheet.xls = sheet.xls
    >>> another_sheet.content
    +---+---+---+
    | 1 | 2 | 3 |
    +---+---+---+
    | 3 | 4 | 5 |
    +---+---+---+

Yet, it is possible assign a absolute url to an online excel file
to an instance of :class:`pyexcel.Sheet`.

**custom attributes**

You can pass on source specific parameters to getter and setter functions.

.. code-block:: python

    >>> content = "1-2-3\n3-4-5"
    >>> sheet = pyexcel.Sheet()
    >>> sheet.set_csv(content, delimiter="-")
    >>> sheet.csv
    '1,2,3\r\n3,4,5\r\n'
    >>> sheet.get_csv(delimiter="|")
    '1|2|3\r\n3|4|5\r\n'

Example::

    >>> import pyexcel as p
    >>> content = {'A': [[1]]}
    >>> b = p.get_book(bookdict=content)
    >>> b
    A:
    +---+
    | 1 |
    +---+
    >>> b[0].name
    'A'
    >>> b[0].name = 'B'
    >>> b
    B:
    +---+
    | 1 |
    +---+

