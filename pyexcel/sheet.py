"""
    pyexcel.sheet
    ~~~~~~~~~~~~~~~~~~~~~

    Building on top of matrix, adding named columns and rows support

    :copyright: (c) 2014-2017 by Onni Software Ltd.
    :license: New BSD License, see LICENSE for more details
"""
import pyexcel._compact as compact
import pyexcel.constants as constants
from pyexcel.internal.sheets.matrix import Matrix
from pyexcel.internal.sheets.row import Row as NamedRow
from pyexcel.internal.sheets.column import Column as NamedColumn


class Sheet(Matrix):
    """Two dimensional data container for filtering, formatting and iteration

    :class:`~pyexcel.Sheet` is a container for a two dimensional array, where
    individual cell can be any Python types. Other than numbers, value of these
    types: string, date, time and boolean can be mixed in the array. This
    differs from Numpy's matrix where each cell are of the same number type.

    In order to prepare two dimensional data for your computation, formatting
    functions help convert array cells to required types. Formatting can be
    applied not only to the whole sheet but also to selected rows or columns.
    Custom conversion function can be passed to these formatting functions. For
    example, to remove extra spaces surrounding the content of a cell, a custom
    function is required.


    **Random access**

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

    .. code-block:: python

        >>> import pyexcel
        >>> content = "1,2,3\\n3,4,5"
        >>> sheet = pyexcel.get_sheet(file_type="csv", file_content=content)
        >>> sheet.tsv
        '1\t2\t3\r\n3\t4\t5\r\n'
        >>> print(sheet.simple)
        csv:
        -  -  -
        1  2  3
        3  4  5
        -  -  -

    What's more, you could as well set value to an attribute, for example:

    .. code-block:: python

        >>> import pyexcel
        >>> content = "1,2,3\n3,4,5"
        >>> sheet = pyexcel.Sheet()
        >>> sheet.csv = content
        >>> sheet.array
        [[1, 2, 3], [3, 4, 5]]

    You can get the direct access to underneath stream object. In some situation,
    it is desired.


    .. code-block:: python

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

    """
    def __init__(self, sheet=None,
                 name=constants.DEFAULT_NAME,
                 name_columns_by_row=-1,
                 name_rows_by_column=-1,
                 colnames=None,
                 rownames=None,
                 transpose_before=False,
                 transpose_after=False):
        """Constructor

        :param sheet: two dimensional array
        :param name: this becomes the sheet name.
        :param name_columns_by_row: use a row to name all columns
        :param name_rows_by_column: use a column to name all rows
        :param colnames: use an external list of strings to name the columns
        :param rownames: use an external list of strings to name the rows
        """
        self.__column_names = []
        self.__row_names = []
        self.__row_index = 0
        self.init(
            sheet=sheet,
            name=name,
            name_columns_by_row=name_columns_by_row,
            name_rows_by_column=name_rows_by_column,
            colnames=colnames,
            rownames=rownames,
            transpose_before=transpose_before,
            transpose_after=transpose_after
        )

    def init(self, sheet=None,
             name=constants.DEFAULT_NAME,
             name_columns_by_row=-1,
             name_rows_by_column=-1,
             colnames=None,
             rownames=None,
             transpose_before=False,
             transpose_after=False):
        """custom initialization functions

        examples::

            >>> import pyexcel as pe
            >>> data = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
            >>> sheet = pe.Sheet(data)
            >>> sheet.row[1]
            [4, 5, 6]
            >>> sheet.row[0:3]
            [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
            >>> sheet.row += [11, 12, 13]
            >>> sheet.row[3]
            [11, 12, 13]
            >>> sheet.row[0:4] = [0, 0, 0] # set all to zero
            >>> sheet.row[3]
            [0, 0, 0]
            >>> sheet.row[0] = ['a', 'b', 'c'] # set one row
            >>> sheet.row[0]
            ['a', 'b', 'c']
            >>> del sheet.row[0] # delete first row
            >>> sheet.row[0] # now, second row becomes the first
            [0, 0, 0]
            >>> del sheet.row[0:]
            >>> sheet.row[0]  # nothing left
            Traceback (most recent call last):
                ...
            IndexError
        """
        # this get rid of phatom data by not specifying sheet
        if sheet is None:
            sheet = []
        Matrix.__init__(self, sheet)
        self.name = name
        self.__column_names = []
        self.__row_names = []
        if transpose_before:
            self.transpose()
        self.row = NamedRow(self)
        self.column = NamedColumn(self)
        if name_columns_by_row != -1:
            if colnames:
                raise NotImplementedError(
                    constants.MESSAGE_NOT_IMPLEMENTED_02)
            self.name_columns_by_row(name_columns_by_row)
        else:
            if colnames:
                self.__column_names = colnames
        if name_rows_by_column != -1:
            if rownames:
                raise NotImplementedError(
                    constants.MESSAGE_NOT_IMPLEMENTED_02)
            self.name_rows_by_column(name_rows_by_column)
        else:
            if rownames:
                self.__row_names = rownames
        if transpose_after:
            self.transpose()

    def transpose(self):
        self.__column_names, self.__row_names = (
            self.__row_names, self.__column_names
        )
        Matrix.transpose(self)

    def name_columns_by_row(self, row_index):
        """Use the elements of a specified row to represent individual columns

        The specified row will be deleted from the data
        :param row_index: the index of the row that has the column names
        """
        self.__row_index = row_index
        self.__column_names = make_names_unique(self.row_at(row_index))
        del self.row[row_index]

    def name_rows_by_column(self, column_index):
        """Use the elements of a specified column to represent individual rows

        The specified column will be deleted from the data
        :param column_index: the index of the column that has the row names
        """
        self.__row_names = make_names_unique(self.column_at(column_index))
        del self.column[column_index]

    def top(self, lines=5):
        """
        Preview top most 5 rows
        """
        sheet = Sheet(self.row[:lines])
        if len(self.colnames) > 0:
            sheet.colnames = self.__column_names
        return sheet

    def top_left(self, rows=5, columns=5):
        """
        Preview top corner: 5x5
        """
        region = Sheet(self.region((0, 0), (rows, columns)))
        if len(self.__row_names) > 0:
            rownames = self.__row_names[:rows]
            region.rownames = rownames
        if len(self.__column_names) > 0:
            columnnames = self.__column_names[:columns]
            region.colnames = columnnames

        return region

    @property
    def colnames(self):
        """Return column names if any"""
        return self.__column_names

    @colnames.setter
    def colnames(self, value):
        """Set column names"""
        self.__column_names = make_names_unique(value)

    @property
    def rownames(self):
        """Return row names if any"""
        return self.__row_names

    @rownames.setter
    def rownames(self, value):
        """Set row names"""
        self.__row_names = make_names_unique(value)

    def named_column_at(self, name):
        """Get a column by its name"""
        index = name
        if compact.is_string(type(index)):
            index = self.colnames.index(name)
        column_array = self.column_at(index)
        return column_array

    def set_named_column_at(self, name, column_array):
        """
        Take the first row as column names

        Given name to identify the column index, set the column to
        the given array except the column name.
        """
        index = name
        if compact.is_string(type(index)):
            index = self.colnames.index(name)
        self.set_column_at(index, column_array)

    def delete_columns(self, column_indices):
        """Delete one or more columns

        :param list column_indices: a list of column indices
        """
        Matrix.delete_columns(self, column_indices)
        if len(self.__column_names) > 0:
            new_series = [self.__column_names[i]
                          for i in range(0, len(self.__column_names))
                          if i not in column_indices]
            self.__column_names = new_series

    def delete_rows(self, row_indices):
        """Delete one or more rows

        :param list row_indices: a list of row indices
        """
        Matrix.delete_rows(self, row_indices)
        if len(self.__row_names) > 0:
            new_series = [self.__row_names[i]
                          for i in range(0, len(self.__row_names))
                          if i not in row_indices]
            self.__row_names = new_series

    def delete_named_column_at(self, name):
        """Works only after you named columns by a row

        Given name to identify the column index, set the column to
        the given array except the column name.
        :param str name: a column name
        """
        if isinstance(name, int):
            if len(self.rownames) > 0:
                self.rownames.pop(name)
            self.delete_columns([name])
        else:
            index = self.colnames.index(name)
            self.colnames.pop(index)
            Matrix.delete_columns(self, [index])

    def named_row_at(self, name):
        """Get a row by its name """
        index = name
        index = self.rownames.index(name)
        row_array = self.row_at(index)
        return row_array

    def set_named_row_at(self, name, row_array):
        """
        Take the first column as row names

        Given name to identify the row index, set the row to
        the given array except the row name.
        """
        index = name
        if compact.is_string(type(index)):
            index = self.rownames.index(name)
        self.set_row_at(index, row_array)

    def delete_named_row_at(self, name):
        """Take the first column as row names

        Given name to identify the row index, set the row to
        the given array except the row name.
        """
        if isinstance(name, int):
            if len(self.rownames) > 0:
                self.rownames.pop(name)
            self.delete_rows([name])
        else:
            index = self.rownames.index(name)
            self.rownames.pop(index)
            Matrix.delete_rows(self, [index])

    def extend_rows(self, rows):
        """Take ordereddict to extend named rows

        :param ordereddist/list rows: a list of rows.
        """
        incoming_data = []
        if isinstance(rows, compact.OrderedDict):
            keys = rows.keys()
            for k in keys:
                self.rownames.append(k)
                incoming_data.append(rows[k])
            Matrix.extend_rows(self, incoming_data)
        elif len(self.rownames) > 0:
            raise TypeError(
                constants.MESSAGE_DATA_ERROR_ORDEREDDICT_IS_EXPECTED)
        else:
            Matrix.extend_rows(self, rows)

    def extend_columns_with_rows(self, rows):
        """Put rows on the right most side of the data"""
        if len(self.colnames) > 0:
            headers = rows.pop(self.__row_index)
            self.__column_names += headers
        Matrix.extend_columns_with_rows(self, rows)

    def extend_columns(self, columns):
        """Take ordereddict to extend named columns

        :param ordereddist/list columns: a list of columns
        """
        incoming_data = []
        if isinstance(columns, compact.OrderedDict):
            keys = columns.keys()
            for k in keys:
                self.colnames.append(k)
                incoming_data.append(columns[k])
            Matrix.extend_columns(self, incoming_data)
        elif len(self.colnames) > 0:
            raise TypeError(
                constants.MESSAGE_DATA_ERROR_ORDEREDDICT_IS_EXPECTED)
        else:
            Matrix.extend_columns(self, columns)

    def to_array(self):
        """Returns an array after filtering"""
        ret = []
        ret += list(self.rows())
        if len(self.rownames) > 0:
            ret = [[value[0]] + value[1] for value in
                   zip(self.rownames, ret)]
            if not compact.PY2:
                ret = list(ret)
        if len(self.colnames) > 0:
            if len(self.rownames) > 0:
                ret.insert(0, [constants.DEFAULT_NA] + self.colnames)
            else:
                ret.insert(0, self.colnames)
        return ret

    def to_records(self, custom_headers=None):
        """
        Make an array of dictionaries

        It takes the first row as keys and the rest of
        the rows as values. Then zips keys and row values
        per each row. This is particularly helpful for
        database operations.
        """
        ret = []
        if len(self.colnames) > 0:
            if custom_headers:
                headers = custom_headers
            else:
                headers = self.colnames
            for row in self.rows():
                the_dict = compact.OrderedDict(zip(headers, row))
                ret.append(the_dict)
        elif len(self.rownames) > 0:
            if custom_headers:
                headers = custom_headers
            else:
                headers = self.rownames
            for column in self.columns():
                the_dict = compact.OrderedDict(zip(headers, column))
                ret.append(the_dict)
        else:
            raise ValueError(
                constants.MESSAGE_DATA_ERROR_NO_SERIES)
        return ret

    def to_dict(self, row=False):
        """Returns a dictionary"""
        the_dict = compact.OrderedDict()
        if len(self.colnames) > 0 and row is False:
            for column in self.named_columns():
                the_dict.update(column)
        elif len(self.rownames) > 0:
            for row in self.named_rows():
                the_dict.update(row)
        else:
            raise NotImplementedError("Not implemented")
        return the_dict

    def named_rows(self):
        """iterate rows using row names"""
        for row_name in self.__row_names:
            yield {row_name: self.row[row_name]}

    def named_columns(self):
        """iterate rows using column names"""
        for column_name in self.__column_names:
            yield {column_name: self.column[column_name]}

    @property
    def content(self):
        """
        Plain representation without headers
        """
        content = self.get_texttable(write_title=False)
        return _RepresentedString(content)

    # python magic methods

    def __getitem__(self, aset):
        if isinstance(aset, tuple):
            if isinstance(aset[0], str):
                row = self.rownames.index(aset[0])
            else:
                row = aset[0]

            if isinstance(aset[1], str):
                column = self.colnames.index(aset[1])
            else:
                column = aset[1]
            return self.cell_value(row, column)
        else:
            return Matrix.__getitem__(self, aset)

    def __setitem__(self, aset, c):
        if isinstance(aset, tuple):
            if isinstance(aset[0], str):
                row = self.rownames.index(aset[0])
            else:
                row = aset[0]

            if isinstance(aset[1], str):
                column = self.colnames.index(aset[1])
            else:
                column = aset[1]
            self.cell_value(row, column, c)
        else:
            Matrix.__setitem__(self, aset, c)

    def __len__(self):
        return self.number_of_rows()


class _RepresentedString(object):
    """present in text"""
    def __init__(self, text):
        self.text = text

    def __repr__(self):
        return self.text

    def __str__(self):
        return self.text


def make_names_unique(alist):
    """Append the number of occurences to duplicated names"""
    duplicates = {}
    new_names = []
    for item in alist:
        if not compact.is_string(type(item)):
            item = str(item)
        if item in duplicates:
            duplicates[item] = duplicates[item] + 1
            new_names.append("%s-%d" % (item, duplicates[item]))
        else:
            duplicates[item] = 0
            new_names.append(item)
    return new_names
