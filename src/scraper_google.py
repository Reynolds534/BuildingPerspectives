import os
import datetime
import urllib
import time
import xlrd
import xlwt

from selenium import webdriver
from selenium.webdriver.common.keys import Keys

from cStringIO import StringIO


#----------GLOBAL SETTINGS------------#

# USE FORWARD SLASHES IN DIRECTORIES ONLY

read_path = os.path.abspath("../files/00_toscrape/SampleBuildings.xlsx")
save_path = "C:/temp/"

min_row = 1
max_row = 2

google_pages_max=30

# Global tracking variables

google_page_cur = 1
	

#----------LAUNCH DRIVER--------------#

driver = webdriver.Firefox()

#----------DEFINITIONS----------------#


# Restarts Firefox within another definition
def startFirefox():
	return webdriver.Firefox()
	
# Get in between substring
def findBetween( s, first, last ):
    try:
        start = s.index( first ) + len( first )
        end = s.index( last, start )
        return s[start:end]
    except ValueError:
        return ""

# Gets Google pages, blurbs, URL links
def getGooglePages(bldg_id,bldg_string):
	global driver
	global google_pages_max
	
	# get google url query
	search_string = '"'+bldg_string+'"'
	search_string += ' went blog'
	google_url = 'https://www.google.com/#q='+urllib.quote_plus(search_string)
	
	print google_url
	
	driver.get(google_url)
	
	results_text=driver.find_element_by_id("resultStats").text
	num_results_str=(findBetween(results_text,"About "," results"))
	num_results=int(num_results_str.replace(',',''))
	
	print num_results
	
	# write num_results to csv at some point
	
	scrapeGooglePage(google_url)
	
	time.sleep(3)
	
	
def scrapeGooglePage(url):
	global google_page_cur
	print google_page_cur
	
	if google_page_cur < google_pages_max:
		scrapeGooglePage(url)
	
	google_page_cur += 1
		
	

def getGoogleResult(bldg_id,bldg_string):
	return



#----------ACTIONS----------------#

book = xlrd.open_workbook(read_path)
sheet = book.sheet_by_index(0)

num_rows = sheet.nrows
print "Number of rows in sheet: "+str(num_rows)

max_row_scrape=min(max_row, num_rows)
print "Max row to scrape: "+str(max_row_scrape)

row_range=range(min_row,max_row_scrape+1)
print "CSV range to scrape:"+str(row_range)

for row_index in row_range:
	bldg_id=str(sheet.cell(row_index,0).value)
	bldg_string=str(sheet.cell(row_index,1).value)
	
	print "-----"	
	print row_index
	print bldg_id
	print bldg_string
	
	getGooglePages(bldg_id,bldg_string)
	time.sleep(3)