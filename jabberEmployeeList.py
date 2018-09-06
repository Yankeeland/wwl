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
import sys, os, xml.etree.ElementTree as et

DEBUG = True

if len(sys.argv) != 2:
    print("Error: please follow usage Rules")
    print("USAGE: jabberEmployeeList.py <filename>")
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

#create the XML Element
root = et.Element("buddylist")

for x in range(1,rows):
    row = tuple(sheet.rows)[x]
    group = et.SubElement(root, "group")
    et.SubElement(group, "gname").text = "WWL"
    user = et.SubElement(group, "user")
    et.SubElement(user, "uname").text = row[1].value
    et.SubElement(user, "fname").text = row[0].value
#    print(row[0].value, row[1].value)


#write the XML file

tree = et.ElementTree(root)
tree.write('jabberEmployeeList.xml', encoding='utf-8', xml_declaration=True)

