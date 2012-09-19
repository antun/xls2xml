#!/usr/bin/python

import os, sys, inspect
# Not sure if this bit is really needed:

# realpath() with make your script run, even if you symlink it :)
cmd_folder = os.path.realpath(os.path.abspath(os.path.split(inspect.getfile( inspect.currentframe() ))[0]))
if cmd_folder not in sys.path:
   sys.path.insert(0, cmd_folder)

# Info:
# cmd_folder = os.path.dirname(os.path.abspath(__file__)) # DO NOT USE __file__ !!!
# __file__ fails if script is called in different ways on Windows
# __file__ fails if someone does os.chdir() before
# sys.argv[0] also fails because it doesn't not always contains the path

from library import xlrd

def toXML(filename):
    xlsFile = xlrd.open_workbook(filename)
    firstSheet = xlsFile.sheet_by_index(0)
    attributes = firstSheet.row_values(0)
    for rownum in range(1, firstSheet.nrows):
        firstSheet.row_values(rownum) 
        cells = [ convertFloatsToIntStrings(i) for i in firstSheet.row_values(rownum) ]
        s = "<node "
        for index in range(len(attributes)):
            s += sanitizeAttributeName(attributes[index].strip()) \
               + "=\"" + cells[index].decode("utf-8").strip() + "\" "
        s += "/>"
        print s

def convertFloatsToIntStrings(i):
    if type(i) is float:
        return str(int(i))
    elif type(i) is long:
        return str(int(i))
    else:
        return i.encode("utf-8", "replace")

def sanitizeAttributeName(dirtyA):
    a = dirtyA.replace(" ", "_")
    a = a.replace("'", "_")
    a = a.replace('"', "_")
    a = a.replace("<", "_")
    a = a.replace(">", "_")
    return a

if __name__ == '__main__':
    args = sys.argv[1:]
    if len(args) < 1:
        print "Usage: xls2xml.py filename"
    else:
        filename = args[0]
        if (len(args) == 2):
            print "<" + args[1] + ">"
        toXML(filename)
        if (len(args) == 2):
            print "</" + args[1] + ">"

