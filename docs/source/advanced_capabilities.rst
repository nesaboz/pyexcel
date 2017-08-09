Advanced signature functions
================================================================================

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

However, you will need to call :meth:`~pyexcel.free_resources` to make sure file
handles are closed.


The python data structures are list, dict, records and book dict. `records`
refers to a list of dictionaries. `book dict` referes to a dictionary of
key-value pair where value is a two dimensional array.



=============================== =================================
Functions                       Description
=============================== ================================= 
:meth:`~pyexcel.isave_as`       Works well with big data files    
:meth:`~pyexcel.isave_book_as`  Works with multiple sheet file
                                and big data files
=============================== =================================

If you would only use these two functions to do format transcoding, you may enjoy a
speed boost using :meth:`~pyexcel.isave_as` and :meth:`~pyexcel.isave_book_as`,
because they use `yield` keyword and minimize memory footprint. However, you will
need to call :meth:`~pyexcel.free_resource` to make sure file handles are closed.
And :meth:`~pyexcel.save_as` and :meth:`~pyexcel.save_book_as` reads all data into
memory and **will make all rows the same width**.
