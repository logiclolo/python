#  python note:
#  (excel write)     http://www.blog.pythonlibrary.org/2014/03/24/creating-microsoft-excel-spreadsheets-with-python-and-xlwt/
#  (excel read)      http://www.blog.pythonlibrary.org/2014/04/30/reading-excel-spreadsheets-with-python-and-xlrd/
#  (xlutils install) http://scicomp.stackexchange.com/questions/2987/what-is-the-simplest-way-to-do-a-user-local-install-of-a-python-package
import sys
import xlwt
import xlrd

sys.path.append('/home/logic.lo/.local/lib/python2.7/site-packages')
from xlutils.copy import copy 

readfile = 'test.xls'
writefile = 'output.xls'
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

    book.save("text.xls")

def read_write(readpath, writepath):
    
    book = xlrd.open_workbook(readpath);
    r_sheet = book.sheet_by_index(0)
    wb = copy(book)
    w_sheet = wb.get_sheet(0)

    w_sheet.write(0, 0, 'HIHIHI')
    wb.save(writepath)



if __name__ == '__main__':
    read()
    #write()
    #read_write(readfile, writefile)
