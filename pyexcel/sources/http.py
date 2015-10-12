"""
    pyexcel.sources.http
    ~~~~~~~~~~~~~~~~~~~

    Representation of http sources

    :copyright: (c) 2015 by Onni Software Ltd.
    :license: New BSD License
"""
from .base import ReadOnlySource, one_sheet_tuple
from ..constants import KEYWORD_URL
from pyexcel_io import load_data
from .._compact import request, PY2
import js


FILE_TYPE_MIME_TABLE = {
    "text/csv": "csv",
    "text/tab-separated-values": "tsv",
    "application/vnd.oasis.opendocument.spreadsheet": "ods",
    "application/vnd.ms-excel": "xls",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
    "application/vnd.ms-excel.sheet.macroenabled.12": "xlsm"
}


def get_file_type_from_url(url):
    extension = url.split('.')
    return extension[-1]


class HttpBookSource(ReadOnlySource):
    """
    Multiple sheet data source via http protocol
    """
    fields = [KEYWORD_URL]

    def __init__(self, url=None, **keywords):
        self.url = url
        self.keywords = keywords

    def get_data(self):
        global SHEETS
        sheets = None
        jquery = js.globals['$']
        xhr = jquery.ajax({"url": self.url, "async":False})
        mime_type = xhr.getResponseHeader('content-type')
        file_type = FILE_TYPE_MIME_TABLE.get(mime_type, None)
        if file_type is None:
            file_type = get_file_type_from_url(self.url)
        sheets = load_data(xhr.responseText,
                           file_type=file_type,
                           **self.keywords)
        return sheets, KEYWORD_URL, None


class HttpSheetSource(HttpBookSource):
    """
    Single sheet data source via http protocol
    """

    fields = [KEYWORD_URL]

    def __init__(self, url=None, **keywords):
        self.url = url
        self.keywords = keywords

    def get_data(self):
        sheets, unused1, unused2 = HttpBookSource.get_data(self)
        print type(sheets)
        print sheets
        return one_sheet_tuple(sheets.items())

