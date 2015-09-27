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
max_row = 5

filename_excel = "GoogleResultHits_"+str(min_row)+"-"+str(max_row)+".xls"

google_pages_max=2

#----------SET UP EXCEL OUTPUT BOOK--------------#
	
writebook=xlwt.Workbook(encoding='latin-1')
writesheet=writebook.add_sheet('SearchResults')
writesheet.set_panes_frozen(True)
writesheet.set_horz_split_pos(1)
writesheet.set_vert_split_pos(1)

writesheet_headers=["bldg_id","result_count"]
writesheet_data1=[]
writesheet_data2=[]

writebook_savepath=os.path.join(save_path,filename_excel)

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
	global writesheet_data
	
	# get google url query
	search_string = '"'+bldg_string+'"'
	search_string += ' went blog'
	google_url = 'https://www.google.com/#q='+urllib.quote_plus(search_string)
	time.sleep(5)
	
	print google_url
	
	driver.get(google_url)
	
	results_text=driver.find_element_by_id("resultStats").text
	num_results_str=(findBetween(results_text,"About "," results"))
	num_results=int(num_results_str.replace(',',''))
	
	print num_results
	
	writesheet_data1.append(bldg_id)
	writesheet_data2.append(num_results)
	
	all_results = []
	for page_num in xrange(int(google_pages_max)):
		page_num = page_num+1 # since it starts at 0
		go_to_page(driver, page_num, search_string)
		titles_urls = scrape_results(driver)
		for title in titles_urls:
			all_results.append(title)
			
	outputFile=open(os.path.join(save_path,"googleresult_"+bldg_id+".txt"), 'w' )
	outputFile.write(str(all_results))
	outputFile.close()

# Go to page, lifted from github search-google.py
def go_to_page(driver, page_num, search_term):
	page_num = page_num - 1
	start_results = page_num * 100
	start_results = str(start_results)
	url = 'https://www.google.com/webhp?#num=100&start='+start_results+'&q='+search_term
	print '[*] Fetching 100 results from page '+str(page_num+1)+' at '+url
	driver.get(url)
	time.sleep(2)

def scrape_results(driver):
    # Xpath will find a subnode of h3, a[@href] specifies that we only want <a> nodes with
    # any href attribute that are subnodes of <h3> tags that have a class of 'r'
    links = driver.find_elements_by_xpath("//h3[@class='r']/a[@href]")
    results = []
    for link in links:
        title = link.text.encode('utf8')
        url = link.get_attribute('href')
        title_url = (title, url)
        results.append(title_url)
    return results



#----------MAIN----------------#

def main():

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
		
		google_page_cur = 1
		getGooglePages(bldg_id,bldg_string)
		time.sleep(3)
		
	# Write to excel
	for i, header in enumerate(writesheet_headers):
		writesheet.write(0,i,header)

	for i, bldg_id in enumerate(writesheet_data1):
		writesheet.write(i+1,0,bldg_id)
		writesheet.write(i+1,1, writesheet_data2[i])
		
	writebook.save(writebook_savepath)
	
main()