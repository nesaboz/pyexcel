Quick start
================================================================================

Four data access functions
--------------------------------------------------------------------------------

.. testcode::
   :hide:

   >>> import pyexcel as p
   >>> content="""
   ... Coffees,Serving Size,Caffeine (mg)
   ... Starbucks Coffee Blonde Roast,venti(20 oz),475
   ... Dunkin' Donuts Coffee with Turbo Shot,large(20 oz.),398
   ... Starbucks Coffee Pike Place Roast,grande(16 oz.),310
   ... Panera Coffee Light Roast,regular(16 oz.),300
   ... """.strip()
   >>> sheet = p.get_sheet(file_content=content, file_type='csv')
   >>> sheet.save_as("your_file.xls")

Suppose you want to process the following excel data :

.. pyexcel-table::

   ---pyexcel:Top 5 coffeine drinks---
   Coffees,Serving Size,Caffeine (mg)
   Starbucks Coffee Blonde Roast,venti(20 oz),475
   Dunkin' Donuts Coffee with Turbo Shot,large(20 oz.),398
   Starbucks Coffee Pike Place Roast,grande(16 oz.),310
   Panera Coffee Light Roast,regular(16 oz.),300

Let's get a list of dictionary out from the xls file:
   
   >>> records = p.get_records(file_name="your_file.xls")
   >>> for record in records:
   ...     print("%s of %s has %s mg" % (
   ...         record['Serving Size'],
   ...         record['Coffees'],
   ...         record['Caffeine (mg)']))
   venti(20 oz) of Starbucks Coffee Blonde Roast has 475 mg
   large(20 oz.) of Dunkin' Donuts Coffee with Turbo Shot has 398 mg
   grande(16 oz.) of Starbucks Coffee Pike Place Roast has 310 mg
   regular(16 oz.) of Panera Coffee Light Roast has 300 mg


Instead, what if you have to use :meth:`pyexcel.get_array` to do the same:

   >>> for row in p.get_array(file_name="your_file.xls", start_row=1):
   ...     print("%s of %s has %s mg" % (
   ...         row[1],
   ...         row[0],
   ...         row[2]))
   venti(20 oz) of Starbucks Coffee Blonde Roast has 475 mg
   large(20 oz.) of Dunkin' Donuts Coffee with Turbo Shot has 398 mg
   grande(16 oz.) of Starbucks Coffee Pike Place Roast has 310 mg
   regular(16 oz.) of Panera Coffee Light Roast has 300 mg

where `start_row` skips the first row, which is the header row.
   
.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("your_file.xls")

Suppose you have a csv, xls, xlsx file as the following:

.. pyexcel-table::

   ---pyexcel:data with columns---
   Column 1,Column 2,Column 3
   1,4,7
   2,5,8
   3,6,9

.. testcode::
   :hide:

   >>> data = [
   ...      ["Column 1", "Column 2", "Column 3"],
   ...      [1, 2, 3],
   ...      [4, 5, 6],
   ...      [7, 8, 9]
   ...  ]
   >>> s = p.Sheet(data)
   >>> s.save_as("example_series.xls")


Now let's get a dictionary out from the spreadsheet:

.. code-block:: python
    
   >>> from pyexcel._compact import OrderedDict
   >>> my_dict = p.get_dict(file_name="example_series.xls", name_columns_by_row=0)
   >>> isinstance(my_dict, OrderedDict)
   True
   >>> for key, values in my_dict.items():
   ...     print({str(key): values})
   {'Column 1': [1, 4, 7]}
   {'Column 2': [2, 5, 8]}
   {'Column 3': [3, 6, 9]}

Please note that my_dict is an OrderedDict.

.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("example_series.xls")


How to obtain a dictionary from a multiple sheet book
-------------------------------------------------------

.. testcode::
   :hide:

   >>> a_dictionary_of_two_dimensional_arrays = {
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
   >>> data = OrderedDict()
   >>> data.update({"Sheet 1": a_dictionary_of_two_dimensional_arrays['Sheet 1']})
   >>> data.update({"Sheet 2": a_dictionary_of_two_dimensional_arrays['Sheet 2']})
   >>> data.update({"Sheet 3": a_dictionary_of_two_dimensional_arrays['Sheet 3']})
   >>> p.save_book_as(bookdict=data, dest_file_name="book.xls")

Suppose you have a multiple sheet book as the following:

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

Here is the code to obtain those sheets as a single dictionary::

   >>> import json
   >>> book_dict = p.get_book_dict(file_name="book.xls")
   >>> isinstance(book_dict, OrderedDict)
   True
   >>> for key, item in book_dict.items():
   ...     print(json.dumps({key: item}))
   {"Sheet 1": [[1, 2, 3], [4, 5, 6], [7, 8, 9]]}
   {"Sheet 2": [["X", "Y", "Z"], [1, 2, 3], [4, 5, 6]]}
   {"Sheet 3": [["O", "P", "Q"], [3, 2, 1], [4, 3, 2]]}

.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("book.xls")


Data export
---------------------------------------------

Suppose you have the following array::

   >>> data = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]

And here is the code to save it as an excel file ::

   >>> p.save_as(array=data, dest_file_name="example.xls")

Let's verify it::

    >>> p.get_sheet(file_name="example.xls")
    pyexcel_sheet1:
    +---+---+---+
    | 1 | 2 | 3 |
    +---+---+---+
    | 4 | 5 | 6 |
    +---+---+---+
    | 7 | 8 | 9 |
    +---+---+---+

.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("example.xls")



Suppose you have the following array::

   >>> data = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]

And here is the code to save it as an excel file ::

   >>> p.save_as(array=data,
   ...           dest_file_name="example.csv",
   ...           dest_delimiter=':')

Let's verify it::

   >>> with open("example.csv") as f:
   ...     for line in f.readlines():
   ...         print(line.rstrip())
   ...
   1:2:3
   4:5:6
   7:8:9

.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("example.csv")


Export data as a multi-sheet excel file
--------------------------------------------------------------------

Suppose you want to save the below dictionary to an excel file ::
  
   >>> a_dictionary_of_two_dimensional_arrays = {
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

Here is the code::

   >>> p.save_book_as(
   ...    bookdict=a_dictionary_of_two_dimensional_arrays,
   ...    dest_file_name="book.xls"
   ... )

If you want to preserve the order of sheets in your dictionary, you have to
pass on an ordered dictionary to the function itself. For example::

   >>> data = OrderedDict()
   >>> data.update({"Sheet 2": a_dictionary_of_two_dimensional_arrays['Sheet 2']})
   >>> data.update({"Sheet 1": a_dictionary_of_two_dimensional_arrays['Sheet 1']})
   >>> data.update({"Sheet 3": a_dictionary_of_two_dimensional_arrays['Sheet 3']})
   >>> p.save_book_as(bookdict=data, dest_file_name="book.xls")

Let's verify its order::

   >>> book_dict = p.get_book_dict(file_name="book.xls")
   >>> for key, item in book_dict.items():
   ...     print(json.dumps({key: item}))
   {"Sheet 2": [["X", "Y", "Z"], [1, 2, 3], [4, 5, 6]]}
   {"Sheet 1": [[1, 2, 3], [4, 5, 6], [7, 8, 9]]}
   {"Sheet 3": [["O", "P", "Q"], [3, 2, 1], [4, 3, 2]]}

Please notice that "Sheet 2" is the first item in the *book_dict*, meaning the order of sheets are preserved.

.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("book.xls")


File format transcoding
-------------------------------------------

.. testcode::
   :hide:

   >>> import datetime
   >>> data = [
   ...    ["name", "weight", "birth"],
   ...    ["Adam", 3.4, datetime.date(2015, 2, 3)],
   ...    ["Smith", 4.2, datetime.date(2014, 11, 12)]
   ... ]
   >>> p.save_as(array=data, dest_file_name="birth.xls")

Suppose we want to save previous used example 'birth.xls' as a csv file ::

   >>> import pyexcel
   >>> p.save_as(file_name="birth.xls", dest_file_name="birth.csv")

Again it is really simple. Let's verify what we have gotten:

   >>> sheet = p.get_sheet(file_name="birth.csv")
   >>> sheet
   birth.csv:
   +-------+--------+----------+
   | name  | weight | birth    |
   +-------+--------+----------+
   | Adam  | 3.4    | 03/02/15 |
   +-------+--------+----------+
   | Smith | 4.2    | 12/11/14 |
   +-------+--------+----------+

.. NOTE::

   Please note that csv(comma separate value) file is pure text file. Formula, charts, images and formatting in xls file will disappear no matter which transcoding tool you use. Hence, pyexcel is a quick alternative for this transcoding job.



How to open an xls file and save it as xlsx
----------------------------------------------------------------------

.. WARNING::

   Formula, charts, images and formatting in xls file will disappear as pyexcel does not support Formula, charts, images and formatting.


Let use previous example and save it as ods instead

   >>> import pyexcel
   >>> p.save_as(file_name="birth.xls",
   ...           dest_file_name="birth.xlsx") # change the file extension

Again let's verify what we have gotten:

   >>> sheet = p.get_sheet(file_name="birth.xlsx")
   >>> sheet
   pyexcel_sheet1:
   +-------+--------+----------+
   | name  | weight | birth    |
   +-------+--------+----------+
   | Adam  | 3.4    | 03/02/15 |
   +-------+--------+----------+
   | Smith | 4.2    | 12/11/14 |
   +-------+--------+----------+

.. testcode::
   :hide:

   >>> os.unlink('birth.xls')
   >>> os.unlink('birth.csv')
   >>> os.unlink('birth.xlsx')


How to open a xls multiple sheet excel book and save it as csv
----------------------------------------------------------------

Well, you write similar codes as before but you will need to use :meth:`~pyexcel.save_book_as` function.
