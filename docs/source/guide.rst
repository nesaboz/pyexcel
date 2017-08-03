Developer's guide
=================

Architecture
--------------

**pyexcel** uses loosely couple plugins to fullfil the promise to access
various file formats. **lml** is the plugin management library that
provide the specialized support for the loose coupling.

The components of **pyexcel** is designed as building blocks. For your
project, you can cherry-pick the file format support without affecting
the core functionality of pyexcel. Each plugin will bring in additional
dependences. For example, if you choose pyexcel-xls, xlrd and xlwt will
be brought in as 2nd level depndencies.

Looking at the following architectural diagram, pyexcel hosts plugin
interfaces for data source, data renderer and data parser. pyexel-pygal,
pyexcel-matplotlib, and pyexce-handsontable extend pyexcel using data
renderer interface. pyexcel-io package takes away the responsibilities
to interface with excel libraries, for example: xlrd, openpyxl, ezodf.

As in :ref:`a-map-of-plugins-and-file-formats`, there are overlapping
capabilities in reading and writing xlsx, ods files. Because each
third parties express different personalities although they may
read and write data in the same file format, you as the pyexcel is
left to pick which suit your task best.

Dotted arrow means the package or module is loaded later.

.. image:: _static/images/architecture.svg

Development steps for code changes

#. git clone https://github.com/pyexcel/pyexcel.git
#. cd pyexcel

Upgrade your setup tools and pip. They are needed for development and testing only:

#. pip install --upgrade setuptools pip

Then install relevant development requirements:

#. pip install -r rnd_requirements.txt # if such a file exists
#. pip install -r requirements.txt
#. pip install -r tests/requirements.txt

Once you have finished your changes, please provide test case(s), relevant documentation
and update CHANGELOG.rst.

.. note::

    As to rnd_requirements.txt, usually, it is created when a dependent
    library is not released. Once the dependecy is installed
    (will be released), the future
    version of the dependency in the requirements.txt will be valid.


How to test your contribution
------------------------------

Although `nose` and `doctest` are both used in code testing, it is adviable that unit tests are put in tests. `doctest` is incorporated only to make sure the code examples in documentation remain valid across different development releases.

On Linux/Unix systems, please launch your tests like this::

    $ make

On Windows systems, please issue this command::

    > test.bat

How to update test environment and update documentation
---------------------------------------------------------

Additional steps are required:

#. pip install moban
#. git clone https://github.com/moremoban/setupmobans.git # generic setup
#. git clone https://github.com/pyexcel/pyexcel-commons.git commons
#. make your changes in `.moban.d` directory, then issue command `moban`

What is pyexcel-commons
---------------------------------

Many information that are shared across pyexcel projects, such as: this developer guide, license info, etc. are stored in `pyexcel-commons` project.

What is .moban.d
---------------------------------

`.moban.d` stores the specific meta data for the library.

Acceptance criteria
-------------------

#. Has Test cases written
#. Has all code lines tested
#. Passes all Travis CI builds
#. Has fair amount of documentation if your change is complex
#. Please update CHANGELOG.rst
#. Please add yourself to CONTRIBUTORS.rst
#. Agree on NEW BSD License for your contribution


