#!/usr/bin/env python
#
# $Id$
#

import os
import sys
import getopt
import xlrd
import unicodecsv


##########################################
# Submodules start here
##########################################

def fromxls2csv(xls_filename, sheet_no, csv_filename, delim):

    # Converts desired page (sheet) of an Excel file to a CSV file.
    # Uses unicodecsv, so it will handle Unicode characters.
    # This should handle both .xls and .xlsx equally well.

    try:
        workbook = xlrd.open_workbook(xls_filename)
    except:
        print("Error opening file %s !" % xls_filename)
        sys.exit(1)

    try:
        sheet = workbook.sheet_by_index(sheet_no)
    except:
        print("%s has no page %d " % (xls_filename, sheet_no))
        sys.exit(1)

    nrows = sheet.nrows
    if nrows < 1:
        print("File %s has no lines to process !" % xls_filename)
        sys.exit(1)

    try:
        myfile = open(csv_filename, "wb")
    except:
        print("Error opening file %s !" % csv_filename)
        sys.exit(1)

    fileout = unicodecsv.writer(myfile, encoding='utf-8', delimiter=delim)
    for row in range(0, nrows):
        fileout.writerow(sheet.row_values(row))

# close output file
    myfile.close()


#######################################
# Main starts here
#######################################
def main():

    # Init variables
    xlsfile = ''
    outfile = ''
    delim = ';'
    sheet_no = 1

# Read command line args
    try:
        myopts, args = getopt.getopt(sys.argv[1:], "i:o:d:p:")
    except getopt.GetoptError as err:
        print(str(err))
        print("Usage: %s -i input -o output [-p sheet -d delim]" % sys.argv[0])
        sys.exit(3)

# opt  == option
# argu == argument passed to the opt

    for opt, argu in myopts:
        if opt == '-i':
            xlsfile = argu
        elif opt == '-o':
            outfile = argu
        elif opt == '-d':
            delim = argu
        elif opt == '-p':
            sheet_no = abs(int(float(argu)))
        else:
            print(
                "Usage: %s -i input -o output [-p sheet -d delim]" % sys.argv[0])
            sys.exit(3)

    if xlsfile == '':
        print("Input Excel file missing!")
        print("Usage: %s -i input -o output [-p sheet -d delim]" % sys.argv[0])
        sys.exit(3)
    elif outfile == '':
        print("Output file missing!")
        print("Usage: %s -i input -o output [-p sheet -d delim]" % sys.argv[0])
        sys.exit(3)

# Make conversion
    fromxls2csv(xlsfile, sheet_no, outfile, delim)


if __name__ == "__main__":
    main()
