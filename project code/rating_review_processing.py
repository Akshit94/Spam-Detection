import time
from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy
import os
import xlwt
import xlsxwriter

rb = open_workbook("SmartPhones_Rating_and_Reviews Lenovo Vibe K5.xls")
ws = rb.sheet_by_index(0)
wb = copy(rb)
wr = wb.get_sheet(0)

average = 0
for row in range(2,4302):
	average = average + int(ws.cell(row,0).value[0])

average = average / (4302)
wr.write(0,2,"Average Rating")
for row in range(2,4302):
	wr.write(row,2,average)

wr.write(0,3,"Review Count")
wr.write(0,4,"Rating Deviation")
wr.write(0,5,"Rating Deviation Weight")
for row in range(2,4302):
	value = int(ws.cell(row,0).value[0]) - average
	value = abs(value)
	if value == 0:
		wr.write(row,3,1)
	elif value >=2:
		wr.write(row,3,0.25)
	else:
		wr.write(row,3,0.5)
wb.save("SmartPhones_Rating_and_Reviews Lenovo Vibe K5.xls")
for row in range(2,4302):
	value = int(ws.cell(row,0).value[0]) - average
	value = abs(value)
	wr.write(row,4,value)
	if value == 0:
		wr.write(row,5,1)
	elif value >=2:
		wr.write(row,5,0.25)
	else:
		wr.write(row,5,0.5)
wb.save("SmartPhones_Rating_and_Reviews Lenovo Vibe K5.xls")
wr.write(0,6,"Caps Score")
wr.write(0,7,"Inverse Rating")
for row in range(2,4302):
	caps_count = 0
	total_words = 0
	review = ws.cell(row,1).value.split(" ")
	for word in review:
		total_words = total_words + 1
		if word.isupper():
			caps_count = caps_count + 1;
	wr.write(row,6,caps_count/float(total_words))
	wr.write(row,7,1-(caps_count/float(total_words)))
wb.save("SmartPhones_Rating_and_Reviews Lenovo Vibe K5.xls")

rb = open_workbook("SmartPhones_Rating_and_Reviews Lenovo Vibe K5.xls")
ws = rb.sheet_by_index(0)
wb = copy(rb)
wr = wb.get_sheet(0)

wr.write(0,8,"Final Score")
wr.write(0,9,"th=0.50")
wr.write(0,10,"th=0.55")
wr.write(0,11,"th=0.60")
wr.write(0,12,"th=0.65")
wr.write(0,13,"th=0.70")
wr.write(0,14,"th=0.75")
wr.write(0,15,"th=0.80")
for row in xrange(2,4302):
	rating_weight = ws.cell(row,5).value
	review_count_weight = ws.cell(row,3).value
	caps_weight = ws.cell(row,7).value
	average_score = (review_count_weight + rating_weight + caps_weight)/float(3)
	wr.write(row,8,average_score)
	if average_score > 0.50:
		wr.write(row, 9 , "Helpful")
	else:
		wr.write(row, 9 , "Not Helpful")
	if average_score > 0.55:
		wr.write(row, 10 , "Helpful")
	else:
		wr.write(row, 10 , "Not Helpful")
	if average_score > 0.60:
		wr.write(row, 11 , "Helpful")
	else:
		wr.write(row, 11 , "Not Helpful")
	if average_score > 0.65:
		wr.write(row, 12 , "Helpful")
	else:
		wr.write(row, 12 , "Not Helpful")
	if average_score > 0.70:
		wr.write(row, 13 , "Helpful")
	else:
		wr.write(row, 13 , "Not Helpful")
	if average_score >= 0.75:
		wr.write(row, 14 , "Helpful")
	else:
		wr.write(row, 14 , "Not Helpful")
	if average_score > 0.80:
		wr.write(row, 15 , "Helpful")
	else:
		wr.write(row, 15 , "Not Helpful")
wb.save("SmartPhones_Rating_and_Reviews Lenovo Vibe K5.xls")
rb = open_workbook("SmartPhones_Rating_and_Reviews Lenovo Vibe K5.xls")
ws = rb.sheet_by_index(0)
wb = copy(rb)
wr = wb.get_sheet(0)

data = [ws.row_values(i) for i in xrange(4302)]
labels = data[0]
data = data[2:]
data.sort(key=lambda x: x[7])

for idx_r, row in enumerate(data):
	for idx_c, value in enumerate(row):
		wr.write(idx_r+1, idx_c, value)

wb.save("SmartPhones_Rating_and_Reviews Lenovo Vibe K5.xls")


