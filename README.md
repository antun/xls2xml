xls2xml
=======

Script for converting tabular data in Excel spreadsheets into simple XML

Usage:

    ./xls2xml.py test/people.xls

    ... and that will output to STDOUT:

    <node ID="1" First_Name="Homer" Last_Name="Simpson" Phone="1-312-118-0853" />
    <node ID="2" First_Name="Wilma" Last_Name="Flintstone" Phone="1-961-423-3877" />
    <node ID="3" First_Name="Danger" Last_Name="Mouse" Phone="1-244-683-1796" />
    <node ID="4" First_Name="Astro" Last_Name="Boy" Phone="1-439-944-5821" />

Notes:

 - There are options to output cells (columns) as nodes instead of
   attributes, to specify the name of the nodes themselves, to add a
   root node. Just run xls2xml.py -h

 - This uses the Python xlrd library, for interacting with Excel files.
   (http://pypi.python.org/pypi/xlrd). That library is bundled here, in
   the "library" directory, to make it easy for people to install.

 - You can put the whole directory in a directory in your path (e.g.
   ~/bin). Make sure that you leave the library directory in place.

 - There are two spreadsheets for testing in the "test" directory.

 - You should set your PYTHONIOENCODING environment variable to utf-8,
   to avoid problems with piping unicode characters from stdout e.g.
        export PYTHONIOENCODING=utf-8
   Otherwise you may run into errors such as:
        UnicodeEncodeError: 'ascii' codec can't encode characters in
        position

