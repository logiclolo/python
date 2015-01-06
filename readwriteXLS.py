#  python note:
#  (excel write)     http://www.blog.pythonlibrary.org/2014/03/24/creating-microsoft-excel-spreadsheets-with-python-and-xlwt/
#  (excel read)      http://www.blog.pythonlibrary.org/2014/04/30/reading-excel-spreadsheets-with-python-and-xlrd/
#  (xlutils install) http://scicomp.stackexchange.com/questions/2987/what-is-the-simplest-way-to-do-a-user-local-install-of-a-python-package
# coding=UTF-8

import xml.etree.ElementTree as ET
import urllib2
from bs4 import BeautifulSoup
import sys
import xlwt
import xlrd

sys.path.append('/home/logic.lo/.local/lib/python2.7/site-packages')
from xlutils.copy import copy 

readfile = 'adj-2.xls'
#readfile = '/Users/logic/personal/GRE/AnkiStuff/adj-2.xls'

writefile = 'output2.xls'
def read():

    book = xlrd.open_workbook(readfile)

    #print sheet names
    print book.sheet_names()

    #get the first worksheet
    first_sheet = book.sheet_by_index(0)

    #read a row
    print first_sheet.row_values(0)

    #read a cell
    cell = first_sheet.cell(0,0)
    print cell
    print cell.value

    #read a row slice
    print first_sheet.row_slice(rowx=0, start_colx=0, end_colx=2)

def write():

    book = xlwt.Workbook()
    sheet1 = book.add_sheet("PySheet1")

    cols = ["A", "B", "C", "D", "E"]
    txt = 'Row %s, Col %s'

    for num in range(5):
        row = sheet1.row(num)
        for index, col in enumerate(cols):
            print index
            print col
            value = txt % (num+1, col)
            row.write(index, value)

    book.save(readfile)

def read_write(readpath, writepath):
    
	start_row = 1
	fillin_col = 4 

	# open file
	purefile = open('output.txt', 'w')

	# open/read excel file
	book = xlrd.open_workbook(readpath)
	r_sheet = book.sheet_by_index(0)
	wb = copy(book)
	w_sheet = wb.get_sheet(0)

	w_sheet.write(0, fillin_col, 'KK')
	for row in range(start_row, r_sheet.nrows):
		aPronounce = []
		# search the kk from YAHOO dictionary
		cell = r_sheet.cell(row, 0)
		print cell.value
		string = u'https://tw.dictionary.yahoo.com/dictionary?p='+cell.value
		req = urllib2.Request(url=string)
		req_str = urllib2.urlopen(req)

		# use beautiful soup to accept the return HTML
		soup = BeautifulSoup(req_str)

		for elem in soup.find_all('span', class_='proun_value'):
			aPronounce.append(elem.text)
		print len(aPronounce) 
		print aPronounce[0]
		w_sheet.write(row, fillin_col, aPronounce[0])

		string_row = r_sheet.cell(row, 0).value + '%' +  r_sheet.cell(row, 1).value.replace('\n','') + '%' + aPronounce[0] + '\t' + r_sheet.cell(row, 2).value.replace('\n','') + '\n'
		#string_row = r_sheet.cell(row, 1).value + 'sdf' + r_sheet.cell(row,2).value
		string_row = string_row.encode('utf-8')
		purefile.write(string_row)
	wb.save(writepath)

	purefile.close()



if __name__ == '__main__':
    #read()
	#write()
	read_write(readfile, writefile)
