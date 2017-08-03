Signature functions
================================================================================

Four data access functions
--------------------------------------------------------------------------------

It is believed that once a Python developer could easily operate on list,
dictionary and various mixture of both. This library provides four module level
functions to help you obtain excel data in those formats. Please refer to
"A list of module level functions", the first three functions operates on any
one sheet from an excel book and the fourth one returns all data in all sheets
in an excel book.

.. table:: A list of module level functions

   =============================== ======================================= ================================ 
   Functions                       Name                                    Python name                      
   =============================== ======================================= ================================ 
   :meth:`~pyexcel.get_array`      two dimensional array                   a list of lists                 
   :meth:`~pyexcel.get_dict`       a dictionary of one dimensional arrays  an ordered dictionary of lists           
   :meth:`~pyexcel.get_records`    a list of dictionaries                  a list of dictionaries           
   :meth:`~pyexcel.get_book_dict`  a dictionary of two dimensional arrays  a dictionary of lists of lists      
   =============================== ======================================= ================================


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

Now, we wanted to draw a bar chart using coffee name vs coffeine. 
   
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


The following two variants of the data access function use generator and should work well with big data files

.. table:: A list of variant functions

   =============================== ======================================= ================================ 
   Functions                       Name                                    Python name                      
   =============================== ======================================= ================================ 
   :meth:`~pyexcel.iget_array`     a memory efficient two dimensional      a generator of a list of lists
                                   array
   :meth:`~pyexcel.iget_records`   a memory efficient list                 a generator of
                                   list of dictionaries                    a list of dictionaries
   =============================== ======================================= ================================

However, you will need to call :meth:`~pyexcel.free_resource` to make sure file
handles are closed.


The python data structures are list, dict, records and book dict. `records`
refers to a list of dictionaries. `book dict` referes to a dictionary of
key-value pair where value is a two dimensional array.


Get back into pyexcel
++++++++++++++++++++++++++++++++

list
********************************************************************************

.. code-block :: python

    >>> import pyexcel as p
    >>> two_dimensional_list = [
    ...    [1, 2, 3, 4],
    ...    [5, 6, 7, 8],
    ...    [9, 10, 11, 12],
    ... ]
    >>> sheet = p.get_sheet(array=two_dimensional_list)
    >>> sheet
    pyexcel_sheet1:
    +---+----+----+----+
    | 1 | 2  | 3  | 4  |
    +---+----+----+----+
    | 5 | 6  | 7  | 8  |
    +---+----+----+----+
    | 9 | 10 | 11 | 12 |
    +---+----+----+----+

dict
***********

.. code-block :: python

    >>> a_dictionary_of_key_value_pair = {
    ...    "IE": 0.2,
    ...    "Firefox": 0.3
    ... }
    >>> sheet = p.get_sheet(adict=a_dictionary_of_key_value_pair)
    >>> sheet
    pyexcel_sheet1:
    +---------+-----+
    | Firefox | IE  |
    +---------+-----+
    | 0.3     | 0.2 |
    +---------+-----+

.. code-block :: python

    >>> a_dictionary_of_one_dimensional_arrays = {
    ...     "Column 1": [1, 2, 3, 4],
    ...     "Column 2": [5, 6, 7, 8],
    ...     "Column 3": [9, 10, 11, 12],
    ... }
    >>> sheet = p.get_sheet(adict=a_dictionary_of_one_dimensional_arrays)
    >>> sheet
    pyexcel_sheet1:
    +----------+----------+----------+
    | Column 1 | Column 2 | Column 3 |
    +----------+----------+----------+
    | 1        | 5        | 9        |
    +----------+----------+----------+
    | 2        | 6        | 10       |
    +----------+----------+----------+
    | 3        | 7        | 11       |
    +----------+----------+----------+
    | 4        | 8        | 12       |
    +----------+----------+----------+

records
*************

.. code-block :: python

    >>> a_list_of_dictionaries = [
    ...     {
    ...         "Name": 'Adam',
    ...         "Age": 28
    ...     },
    ...     {
    ...         "Name": 'Beatrice',
    ...         "Age": 29
    ...     },
    ...     {
    ...         "Name": 'Ceri',
    ...         "Age": 30
    ...     },
    ...     {
    ...         "Name": 'Dean',
    ...         "Age": 26
    ...     }
    ... ]
    >>> sheet = p.get_sheet(records=a_list_of_dictionaries)
    >>> sheet
    pyexcel_sheet1:
    +-----+----------+
    | Age | Name     |
    +-----+----------+
    | 28  | Adam     |
    +-----+----------+
    | 29  | Beatrice |
    +-----+----------+
    | 30  | Ceri     |
    +-----+----------+
    | 26  | Dean     |
    +-----+----------+

book dict
**************

.. code-block :: python

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
    >>> book = p.get_book(bookdict=a_dictionary_of_two_dimensional_arrays)
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

Two pyexcel functions
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

In cases where the excel data needs custom manipulations, a pyexcel user got a
few choices: one is to use :class:`~pyexcel.Sheet` and :class:`~pyexcel.Book`,
the other is to look for more sophisticated ones:

* Pandas, for numerical analysis
* Do-it-yourself

=============================== ================================ 
Functions                       Returns                      
=============================== ================================ 
:meth:`~pyexcel.get_sheet`      :class:`~pyexcel.Sheet`
:meth:`~pyexcel.get_book`       :class:`~pyexcel.Book`
=============================== ================================ 

For all six functions, you can pass on the same command parameters while the
return value is what the function says.


Export data from Python
--------------------------------------------------------------------------------

This library provides one application programming interface to transform them
into one of the data structures:

   * two dimensional array
   * a (ordered) dictionary of one dimensional arrays
   * a list of dictionaries
   * a dictionary of two dimensional arrays
   * a :class:`~pyexcel.Sheet`
   * a :class:`~pyexcel.Book`

and write to one of the following data sources:

   * physical file
   * memory file
   * SQLAlchemy table
   * Django Model
   * Python data structures: dictionary, records and array


Here are the two functions:

=============================== =================================
Functions                       Description
=============================== ================================= 
:meth:`~pyexcel.save_as`        Works well with single sheet file
:meth:`~pyexcel.isave_as`       Works well with big data files    
:meth:`~pyexcel.save_book_as`   Works with multiple sheet file
                                and big data files
:meth:`~pyexcel.isave_book_as`  Works with multiple sheet file
                                and big data files
=============================== =================================

If you would only use these two functions to do format transcoding, you may enjoy a
speed boost using :meth:`~pyexcel.isave_as` and :meth:`~pyexcel.isave_book_as`,
because they use `yield` keyword and minimize memory footprint. However, you will
need to call :meth:`~pyexcel.free_resource` to make sure file handles are closed.
And :meth:`~pyexcel.save_as` and :meth:`~pyexcel.save_book_as` reads all data into
memory and **will make all rows the same width**.

See also:

* :ref:`save_an_array_to_an_excel_sheet`
* :ref:`save_an_book_dict_to_an_excel_book`
* :ref:`save_an_array_to_a_csv_with_custom_delimiter`

How to save an python array as an excel file
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

How to save a dictionary of two dimensional array as an excel file
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

How to an excel sheet to a database using SQLAlchemy
----------------------------------------------------

.. NOTE::

   You can find the complete code of this example in examples folder on github

Before going ahead, let's import the needed components and initialize sql
engine and table base::

   >>> from sqlalchemy import create_engine
   >>> from sqlalchemy.ext.declarative import declarative_base
   >>> from sqlalchemy import Column , Integer, String, Float, Date
   >>> from sqlalchemy.orm import sessionmaker
   >>> engine = create_engine("sqlite:///birth.db")
   >>> Base = declarative_base()
   >>> Session = sessionmaker(bind=engine)

Let's suppose we have the following database model:

   >>> class BirthRegister(Base):
   ...     __tablename__='birth'
   ...     id=Column(Integer, primary_key=True)
   ...     name=Column(String)
   ...     weight=Column(Float)
   ...     birth=Column(Date)

Let's create the table::
  
   >>> Base.metadata.create_all(engine)

Now here is a sample excel file to be saved to the table:


.. pyexcel-table::
   
   ---pyexcel:data table---
   name,weight,birth     
   Adam,3.4,2015-02-03
   Smith,4.2,2014-11-12

.. testcode::
   :hide:

   >>> import datetime
   >>> data = [
   ...    ["name", "weight", "birth"],
   ...    ["Adam", 3.4, datetime.date(2015, 2, 3)],
   ...    ["Smith", 4.2, datetime.date(2014, 11, 12)]
   ... ]
   >>> p.save_as(array=data, dest_file_name="birth.xls")

Here is the code to import it:

   >>> session = Session() # obtain a sql session
   >>> p.save_as(file_name="birth.xls", name_columns_by_row=0, dest_session=session, dest_table=BirthRegister)

Done it. It is that simple. Let's verify what has been imported to make sure.

   >>> sheet = p.get_sheet(session=session, table=BirthRegister)
   >>> sheet
   birth:
   +------------+----+-------+--------+
   | birth      | id | name  | weight |
   +------------+----+-------+--------+
   | 2015-02-03 | 1  | Adam  | 3.4    |
   +------------+----+-------+--------+
   | 2014-11-12 | 2  | Smith | 4.2    |
   +------------+----+-------+--------+

.. testcode::
   :hide:

   >>> session.close()
   >>> os.unlink('birth.db')

.. _save_a_xls_as_a_csv:

How to open an xls file and save it as csv
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


.. _save_a_xls_as_a_xlsx:

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

   >>> session.close()
   >>> os.unlink('birth.xls')
   >>> os.unlink('birth.csv')
   >>> os.unlink('birth.xlsx')



How to open a xls multiple sheet excel book and save it as csv
----------------------------------------------------------------

Well, you write similar codes as before but you will need to use :meth:`~pyexcel.save_book_as` function.

  
Data transportation/transcoding
--------------------------------------------------------------------------------

Based the capability of this library, it is capable of transporting your data in
between any of these data sources:

   * physical file
   * memory file
   * SQLAlchemy table
   * Django Model
   * Python data structures: dictionary, records and array

See also:

* :ref:`import_excel_sheet_into_a_database_table`
* :ref:`save_a_xls_as_a_xlsx`
* :ref:`save_a_xls_as_a_csv`
