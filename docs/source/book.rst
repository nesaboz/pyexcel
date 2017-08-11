Book
=========


You access each cell via this syntax::

    book[sheet_index][row, column]

or::

    book["sheet_name"][row, column]

Suppose you have the following sheets:

.. pyexcel-table::

   ---pyexcel:Sheet 1---
   1,2,3
   4,5,6
   7,8,9
   ---pyexcel---
   ---pyexcel:Sheet 2---
   X,Y,Z
   1,2,3
   4,5,6
   ---pyexcel---
   ---pyexcel:Sheet 3---
   O,P,Q
   3,2,1
   4,3,2

.. testcode::
   :hide:

   >>> data = {
   ...      'Sheet 1':
   ...          [
   ...              [1.0, 2.0, 3.0],
   ...              [4.0, 5.0, 6.0],
   ...              [7.0, 8.0, 9.0]
   ...          ],
   ...      'Sheet 2':
   ...          [
   ...              ['X', 'Y', 'Z'],
   ...              [1.0, 2.0, 3.0],
   ...              [4.0, 5.0, 6.0]
   ...          ],
   ...      'Sheet 3':
   ...          [
   ...              ['O', 'P', 'Q'],
   ...              [3.0, 2.0, 1.0],
   ...              [4.0, 3.0, 2.0]
   ...          ]
   ...  }
   >>> book = pyexcel.Book(data)
   >>> book.save_as("example.xls")

And you can randomly access a cell in a sheet::

    >>> book = pyexcel.get_book(file_name="example.xls")
    >>> print(book["Sheet 1"][0,0])
    1
    >>> print(book[0][0,0]) # the same cell
    1

.. TIP::
  With pyexcel, you can regard single sheet reader as an
  two dimensional array and multi-sheet excel book reader
  as a ordered dictionary of two dimensional arrays.

**Write multiple sheet excel file**

Suppose you have previous data as a dictionary and you want to 
save it as multiple sheet excel file::

    >>> content = {
    ...     'Sheet 1':
    ...         [
    ...             [1.0, 2.0, 3.0],
    ...             [4.0, 5.0, 6.0],
    ...             [7.0, 8.0, 9.0]
    ...         ],
    ...     'Sheet 2':
    ...         [
    ...             ['X', 'Y', 'Z'],
    ...             [1.0, 2.0, 3.0],
    ...             [4.0, 5.0, 6.0]
    ...         ],
    ...     'Sheet 3':
    ...         [
    ...             ['O', 'P', 'Q'],
    ...             [3.0, 2.0, 1.0],
    ...             [4.0, 3.0, 2.0]
    ...         ]
    ... }
    >>> book = pyexcel.get_book(bookdict=content)
    >>> book.save_as("output.xls")

You shall get a xls file


**Read multiple sheet excel file**

Let's read the previous file back:

    >>> book = pyexcel.get_book(file_name="output.xls")
    >>> sheets = book.to_dict()
    >>> for name in sheets.keys():
    ...     print(name)
    Sheet 1
    Sheet 2
    Sheet 3

Get content
************

.. code-block:: python

    >>> book_dict = {
    ...      'Sheet 2':
    ...          [
    ...              ['X', 'Y', 'Z'],
    ...              [1.0, 2.0, 3.0],
    ...              [4.0, 5.0, 6.0]
    ...          ],
    ...      'Sheet 3':
    ...          [
    ...              ['O', 'P', 'Q'],
    ...              [3.0, 2.0, 1.0],
    ...              [4.0, 3.0, 2.0]
    ...          ],
    ...      'Sheet 1':
    ...          [
    ...              [1.0, 2.0, 3.0],
    ...              [4.0, 5.0, 6.0],
    ...              [7.0, 8.0, 9.0]
    ...          ]
    ...  }
    >>> book = pyexcel.get_book(bookdict=book_dict)
    >>> book
    Sheet 1:
    +-----+-----+-----+
    | 1.0 | 2.0 | 3.0 |
    +-----+-----+-----+
    | 4.0 | 5.0 | 6.0 |
    +-----+-----+-----+
    | 7.0 | 8.0 | 9.0 |
    +-----+-----+-----+
    Sheet 2:
    +-----+-----+-----+
    | X   | Y   | Z   |
    +-----+-----+-----+
    | 1.0 | 2.0 | 3.0 |
    +-----+-----+-----+
    | 4.0 | 5.0 | 6.0 |
    +-----+-----+-----+
    Sheet 3:
    +-----+-----+-----+
    | O   | P   | Q   |
    +-----+-----+-----+
    | 3.0 | 2.0 | 1.0 |
    +-----+-----+-----+
    | 4.0 | 3.0 | 2.0 |
    +-----+-----+-----+
    >>> print(book.rst)
    Sheet 1:
    =  =  =
    1  2  3
    4  5  6
    7  8  9
    =  =  =
    Sheet 2:
    ===  ===  ===
    X    Y    Z
    1.0  2.0  3.0
    4.0  5.0  6.0
    ===  ===  ===
    Sheet 3:
    ===  ===  ===
    O    P    Q
    3.0  2.0  1.0
    4.0  3.0  2.0
    ===  ===  ===

You can get the direct access to underneath stream object. In some situation,
it is desired.


.. code-block:: python

    >>> stream = sheet.stream.plain

The returned stream object has the content formatted in plain format
for further reading.


Set content
************

Surely, you could set content to an instance of :class:`pyexcel.Book`.

.. code-block:: python

    >>> other_book = pyexcel.Book()
    >>> other_book.bookdict = book_dict
    >>> print(other_book.plain)
    Sheet 1:
    1  2  3
    4  5  6
    7  8  9
    Sheet 2:
    X    Y    Z
    1.0  2.0  3.0
    4.0  5.0  6.0
    Sheet 3:
    O    P    Q
    3.0  2.0  1.0
    4.0  3.0  2.0

You can set via 'xls' attribute too.

.. code-block:: python

    >>> another_book = pyexcel.Book()
    >>> another_book.xls = other_book.xls
    >>> print(another_book.mediawiki)
    Sheet 1:
    {| class="wikitable" style="text-align: left;"
    |+ <!-- caption -->
    |-
    | align="right"| 1 || align="right"| 2 || align="right"| 3
    |-
    | align="right"| 4 || align="right"| 5 || align="right"| 6
    |-
    | align="right"| 7 || align="right"| 8 || align="right"| 9
    |}
    Sheet 2:
    {| class="wikitable" style="text-align: left;"
    |+ <!-- caption -->
    |-
    | X || Y || Z
    |-
    | 1 || 2 || 3
    |-
    | 4 || 5 || 6
    |}
    Sheet 3:
    {| class="wikitable" style="text-align: left;"
    |+ <!-- caption -->
    |-
    | O || P || Q
    |-
    | 3 || 2 || 1
    |-
    | 4 || 3 || 2
    |}
