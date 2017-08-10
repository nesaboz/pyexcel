
=============
API Reference
=============

.. currentmodule:: pyexcel
.. _api:


This is intended for users of pyexcel.

.. _signature-functions:

Signature functions
====================

.. _conversion-from:


Obtaining data from excel file
-------------------------------

It is believed that once a Python developer could easily operate on list,
dictionary and various mixture of both. This library provides four module level
functions to help you obtain excel data in those formats. Please refer to
"A list of module level functions", the first three functions operates on any
one sheet from an excel book and the fourth one returns all data in all sheets
in an excel book.

.. autosummary::
   :toctree: generated/

   get_array
   get_dict
   get_records
   get_book_dict

In cases where the excel data needs custom manipulations, a pyexcel user got a
few choices: one is to use :class:`~pyexcel.Sheet` and :class:`~pyexcel.Book`,
the other is to look for more sophisticated ones:

* Pandas, for numerical analysis
* Do-it-yourself

.. autosummary::
   :toctree: generated/

   get_book
   get_sheet

The following two variants of the data access function use generator and should work well with big data files. However, you will need to call :meth:`~pyexcel.free_resources` to make sure file handles are closed.


.. autosummary::
   :toctree: generated/

   iget_array
   iget_records
   free_resources

.. _conversion-to:

Saving data to excel file
--------------------------

.. autosummary::
   :toctree: generated/

   save_as
   save_book_as


The following functions would work with big data and will work every well
with :meth:`~pyexcel.iget_array` and :meth:`~pyexcel.iget_records`.

.. autosummary::
   :toctree: generated/

   isave_as
   isave_book_as

If you would only use these two functions to do format transcoding, you may enjoy a
speed boost using :meth:`~pyexcel.isave_as` and :meth:`~pyexcel.isave_book_as`,
because they use `yield` keyword and minimize memory footprint. However, you will
need to call :meth:`~pyexcel.free_resources` to make sure file handles are closed.
And :meth:`~pyexcel.save_as` and :meth:`~pyexcel.save_book_as` reads all data into
memory and **will make all rows the same width**.


Cookbook
==========

.. autosummary::
   :toctree: generated/

   merge_csv_to_a_book
   merge_all_to_a_book
   split_a_book
   extract_a_sheet_from_a_book

   
Book 
=====

Here's the entity relationship between Book, Sheet, Row and Column

.. image:: entity-relationship-diagram.png

Constructor
------------

.. autosummary::
   :toctree: generated/

   Book

Attribute
------------

.. autosummary::
   :toctree: generated/

   Book.number_of_sheets
   Book.sheet_names

Conversions
-------------

.. autosummary::
   :toctree: generated/

   Book.bookdict
   Book.url
   Book.csv
   Book.tsv
   Book.csvz
   Book.tsvz
   Book.xls
   Book.xlsm
   Book.xlsx
   Book.ods
   Book.stream

Save changes
-------------

.. autosummary::
   :toctree: generated/

   Book.save_as
   Book.save_to_memory
   Book.save_to_database

Sheet
=====


Constructor
-----------

.. autosummary::
   :toctree: generated/

   Sheet


Attributes
-----------

.. autosummary::
   :toctree: generated/

   Sheet.content
   Sheet.number_of_rows
   Sheet.number_of_columns
   Sheet.row_range
   Sheet.column_range

Iteration
-----------------

.. autosummary::
   :toctree: generated/

   Sheet.rows
   Sheet.rrows
   Sheet.columns
   Sheet.rcolumns
   Sheet.enumerate
   Sheet.reverse
   Sheet.vertical
   Sheet.rvertical


Cell access
------------------

.. autosummary::
   :toctree: generated/

   Sheet.cell_value
   Sheet.__getitem__

Row access
------------------

.. autosummary::
   :toctree: generated/

   Sheet.row_at
   Sheet.set_row_at
   Sheet.delete_rows
   Sheet.extend_rows

Column access
--------------

.. autosummary::
   :toctree: generated/

   Sheet.column_at
   Sheet.set_column_at
   Sheet.delete_columns
   Sheet.extend_columns


Data series
------------


Any column as row name
************************

.. autosummary::
   :toctree: generated/

   Sheet.name_columns_by_row
   Sheet.rownames
   Sheet.named_column_at
   Sheet.set_named_column_at
   Sheet.delete_named_column_at


Any row as column name
************************

.. autosummary::
   :toctree: generated/

   Sheet.name_rows_by_column
   Sheet.colnames
   Sheet.named_row_at
   Sheet.set_named_row_at
   Sheet.delete_named_row_at

   
Conversion
-------------

.. autosummary::
   :toctree: generated/

   Sheet.array
   Sheet.records
   Sheet.dict
   Sheet.url
   Sheet.csv
   Sheet.tsv
   Sheet.csvz
   Sheet.tsvz
   Sheet.xls
   Sheet.xlsm
   Sheet.xlsx
   Sheet.ods
   Sheet.stream


Formatting
------------------

.. autosummary::
   :toctree: generated/

   Sheet.format

Filtering
-----------

.. autosummary::
   :toctree: generated/

   Sheet.filter


Transformation
----------------

.. autosummary::
   :toctree: generated/

   Sheet.transpose
   Sheet.map
   Sheet.region
   Sheet.cut
   Sheet.paste
        
Save changes
--------------

.. autosummary::
   :toctree: generated/

   Sheet.save_as
   Sheet.save_to_memory
   Sheet.save_to_database
