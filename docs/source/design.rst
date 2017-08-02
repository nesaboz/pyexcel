Introduction
================================================================================

**pyexcel** provides single application programming interface(API) to read, write and
manipulate data in different excel file formats, in different storage media(
disk, memory, database) and in different python data structures. Its loosely
coupled architecture and well-defined plugin interface make the library easily
extensible without break the promise of the single API.

The excel file formats are csv, xls, xlsx and ods. For example, the following
code will read any of the supported formats in the directory::

    import pyexcel as p
	import glob
	for excel_file in glob.glob("*"):
	    book = p.get_book(file_name=excel_file)
        # then do some manipulation on book

.. table:: a list of support file formats

    ============ =======================================================
    file format  definition
    ============ =======================================================
    csv          comma separated values
    tsv          tab separated values
    csvz         a zip file that contains one or many csv files
    tsvz         a zip file that contains one or many tsv files
    xls          a spreadsheet file format created by
                 MS-Excel 97-2003 [#f1]_
    xlsx         MS-Excel Extensions to the Office Open XML
                 SpreadsheetML File Format. [#f2]_
    xlsm         an MS-Excel Macro-Enabled Workbook file
    ods          open document spreadsheet
	fods         flat open document spreadsheet
    json         java script object notation
    html         html table of the data structure
    simple       simple presentation
    rst          rStructured Text presentation of the data
    mediawiki    media wiki table
    ============ =======================================================

		 
.. [#f1] quoted from `whatis.com <http://whatis.techtarget.com/fileformat/XLS-Worksheet-file-Microsoft-Excel>`_. Technical details can be found at `MSDN XLS <https://msdn.microsoft.com/en-us/library/office/gg615597(v=office.14).aspx>`_
.. [#f2] xlsx is used by MS-Excel 2007, more information can be found at `MSDN XLSX <https://msdn.microsoft.com/en-us/library/dd922181(v=office.12).aspx>`_


The python data structures are list, dict, records and book dict. `records`
refers to a list of dictionaries. `book dict` referes to a dictionary of
key-value pair where value is a two dimensional array.

.. _a-list-of-data-structures:
.. table:: A list of supported data structures

   ======================================= ================================ =========================
   Pesudo name                             Python name                      Related model
   ======================================= ================================ =========================
   two dimensional array                   a list of lists                  :class:`pyexcel.Sheet`
   a dictionary of key value pair          a dictionary                     :class:`pyexcel.Sheet`
   a dictionary of one dimensional arrays  a dictionary of lists            :class:`pyexcel.Sheet`
   a list of dictionaries                  a list of dictionaries           :class:`pyexcel.Sheet`
   a dictionary of two dimensional arrays  a dictionary of lists of lists   :class:`pyexcel.Book`
   ======================================= ================================ =========================

Examples of supported data structure
--------------------------------------------------------------------------------

list
********************************************************************************

::
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

::
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

::
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

::
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

::
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
