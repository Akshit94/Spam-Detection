from BeautifulSoup import *
from selenium import webdriver
import time
from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy
from selenium.webdriver import ActionChains
import os
import xlsxwriter
from selenium.webdriver.common.keys import Keys

workbook = xlsxwriter.Workbook("SmartPhones_Rating_and_Reviews Lenovo Vibe K5.xls")
worksheet=workbook.add_worksheet()
workbook.close()



rb = open_workbook("SmartPhones_Rating_and_Reviews Lenovo Vibe K5.xls")
ws = rb.sheet_by_index(0)
r = ws.nrows
r += 2

wb = copy(rb)

ws = wb.get_sheet(0)

Setting up heading for the excel sheet.

ws.write(0,0,"Rating")
ws.write(0,1,"Reviews_123")
ws.write(0,16,"User Name")

driver = webdriver.Firefox()
driver.get("http://www.amazon.in/product-reviews/B01M0ULUC8/ref=cm_cr_dp_see_all_summary?ie=UTF8&reviewerType=all_reviews&showViewpoints=1&sortBy=helpful")
# j=1
for i in range(430):

	list_of_reviews_html_text = driver.find_elements_by_css_selector(".a-section.review")
	for each in list_of_reviews_html_text:
		html_body = BeautifulSoup(each.get_attribute('innerHTML'))
		rating = html_body.div.a.i.getText()
		ws.write(r,0,rating)
		review = html_body.find('div',{"class" : "a-row review-data"}).getText()
		ws.write(r,1,review)
		author_name = html_body.find('a',{"class" : "a-size-base a-link-normal author"}).getText()
		ws.write(r,16,author_name)
		r=r+1
		wb.save("SmartPhones_Rating_and_Reviews Lenovo Vibe K5.xls")
		driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
	time.sleep(6)
	driver.find_element_by_class_name("a-last").click()
	time.sleep(6)
	wb.save("SmartPhones_Rating_and_Reviews Lenovo Vibe K5.xls")



# Adding user information ot another spreadsheet
workbook = xlsxwriter.Workbook("SmartPhones_User_Info Lenovo Vibe K5.xls")
worksheet=workbook.add_worksheet()
workbook.close()
rb1 = open_workbook("SmartPhones_User_Info Lenovo Vibe K5.xls")
ws1 = rb1.sheet_by_index(0)
r1 = ws1.nrows
r1 += 2

rb = open_workbook("SmartPhones_Rating_and_Reviews Lenovo Vibe K5.xls")
ws = rb.sheet_by_index(0)
no_of_rows = ws.nrows
wb1 = copy(rb1)
ws1 = wb1.get_sheet(0)

ws1.write(0,0,"User Name")
ws1.write(0,1,"No of reviews")
print no_of_rows
# driver = webdriver.Firefox()
driver.get("http://www.amazon.in/product-reviews/B01M0ULUC8/ref=cm_cr_dp_see_all_summary?ie=UTF8&reviewerType=all_reviews&showViewpoints=1&sortBy=helpful")
map = []
for i in range(430):

	list_of_reviews_html_text = driver.find_elements_by_css_selector(".a-section.review")
	for each in list_of_reviews_html_text:
		html_body = BeautifulSoup(each.get_attribute('innerHTML'))
		author_name = html_body.find('a',{"class" : "a-size-base a-link-normal author"}).getText()
		if author_name in map:
			pass
		else:
			ws1.write(r1,0,author_name)
			map.append(author_name)
		r1=r1+1
		driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
		wb1.save("SmartPhones_User_Info Lenovo Vibe K5.xls")
	time.sleep(6)
	driver.find_element_by_class_name("a-last").click()
	time.sleep(6)
	wb1.save("SmartPhones_User_Info Lenovo Vibe K5.xls")


for r1 in range(2,4258):
	count=1 # The no of reviews per author
	try:
		for row in range(2,no_of_rows):
			# print ws.cell(0,1).value
			name = str(ws.cell(r1,16).value)
			if str(ws.cell(row,16)) == name:
				count=count+1
	except:
		pass
	print count
	ws1.write(r1,1,count)
	r1=r1+1
	wb1.save("SmartPhones_User_Info Lenovo Vibe K5.xls")
