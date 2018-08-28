#Program: jabberEmployeeList.py
#Author: Jake Davis
#Date: 8/28/2018
#Description: Generates an .xml file that can be used to import contacts into Jabber
#INPUT: excel file listing all employees of a company
#OUTPUT: .xml file formatted so that Jabber can import and update its contact list
#USAGE: jabberEmployeeList.py <filename>

#! python3

#load modules
from openpyxl import *
import sys

DEBUG = True

if len(sys.argv) != 2:
    print("Error: please follow usage Rules")
    print("USAGE: randomemployee.py <filename>")
    sys.exit()

FILENAME = sys.argv[1]

#read in the excel file or throw exception
try:
    wb = load_workbook(FILENAME)
except:
    print("ERROR: No workboook found:" + FILENAME)

#get the active worksheet and how many rows there are
sheet = wb.active
rows = sheet.max_row



#for each line in the input sheet, create and XML entry
