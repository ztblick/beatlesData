#!/usr/bin/python

import sys
import os
import re
import xlwt
import xlrd

#--------------selection criteria------------------#

def accept(desc): #CHANGE THIS SHEEEEE
	if desc < 49.199:
		return True
	else:
		return False
	#if #major failure tests here#
	#	return FALSE

#-------------------main--------------------------#

#first, access input data from local copy called "input.xlsx"
workbook = xlrd.open_workbook('input.xlsx')
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
	if accept(data[i][3]): #CHANGE THIS BACK TO 2
		for j in range(4):
			sheet.write(outRow, j, data[i][j])
		outRow+=1

#finally, save a clean data local copy called "output.xlsx"
workbook.save('output.xls')