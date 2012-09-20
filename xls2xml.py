#!/usr/bin/python

import os, sys, inspect, argparse
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

def toXML(filename, node_name):
    xlsFile = xlrd.open_workbook(filename)
    firstSheet = xlsFile.sheet_by_index(0)
    attributes = firstSheet.row_values(0)
    for rownum in range(1, firstSheet.nrows):
        firstSheet.row_values(rownum) 
        cells = [ convertFloatsToIntStrings(i) for i in firstSheet.row_values(rownum) ]
        s = "<" + node_name + " "
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
    parser = argparse.ArgumentParser(description='Converts tabular data in Excel spreadsheets into XML.')
    parser.add_argument('filename', help='Excel (xls or xlsx) file.')
    parser.add_argument("-n", "--node", help="node for each table row", default="node")
    parser.add_argument("-r", "--root-node", help="wrap nodes in specified root node")
    args = parser.parse_args()
    if (args.root_node != None):
        print "<" + args.root_node + ">"
    toXML(args.filename, args.node)
    if (args.root_node != None):
        print "</" + args.root_node + ">"
