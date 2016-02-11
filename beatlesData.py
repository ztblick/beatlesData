#!/usr/bin/python

import sys
import os
import re
import xlwt
import xlrd

#--------------selection criteria------------------#

def accept(desc):
	
	#delete if no string is passed
	if not desc:
		return False

	#if a year is present, it must be 1968-1971.
	m = re.search('19\d{2}', desc)
	if m and m.group(0)!="1968" and m.group(0)!="1969"\
		and m.group(0)!="1970" and m.group(0)!="1971":
		return False
	
	#remove orange labels
	m = re.search('orange', desc, re.IGNORECASE)
	if m:
		return False

	#remove purple labels
	m = re.search('purple', desc, re.IGNORECASE)
	if m:
		return False

	#delete other albums
	albums = "rubber|revolver|abbey|submarine|mystery|jude"
	m = re.search(albums, desc, re.IGNORECASE)
	if m:
		return False

	#must have white-album in title
	m = re.search('white\Walbum', desc, re.IGNORECASE)
	if not m:
		return False

	#delete specific low numbers, #00xxxxx
	m = re.search('\d{5}', desc)
	if m:
		return False

	#only accept US and UK
	countries = "China|Chinese|Japan|Japanese|Germany|Venezuela|\
	German|France|French|India|Italy|Italian|Brazil|Portugese|\
	Korea|Korean|Russia|Russian|Spain|Mexico|Spanish|Indonesia|\
	Netherlands|Dutch|Deutsche|Turkey|Turkish|Taiwan|Thailand|\
	Malaysia|Sweden|Swedish|Poland|Polish|Greek|Greece|Danish|\
	Canada|Canadian|Denmark|Norway|Norwegian|Belgium"
	m = re.search(countries, desc, re.IGNORECASE)
	if m:
		return False

	#delete random crap
	crap = "necktie"
	m = re.search(crap, desc, re.IGNORECASE)
	if m:
		return False

	return True

#-------------------main--------------------------#

#first, access input data from local copy called "input.xlsx"
workbook = xlrd.open_workbook('Zach_1.12.16.xlsx')
sheet = workbook.sheet_by_index(0)
nrows = sheet.nrows

#copy all rows into a massive list indexed by excel rows
data = []
for row in range(nrows):
	data.append(sheet.row_values(row))

#now we create the output workbook
workbook = xlwt.Workbook()
sheet = workbook.add_sheet('clean')

#we iterate over all original rows, checking the description
#all selection criteria are handeled by the accept function
outRow = 0
for i in range(nrows):
	desc = repr(data[i][2])
	if accept(desc):
		for j in range(4):
			sheet.write(outRow, j, data[i][j])
		outRow+=1

#finally, save a clean data local copy called "output.xlsx"
workbook.save('Zach_Cleaned Data.xls')