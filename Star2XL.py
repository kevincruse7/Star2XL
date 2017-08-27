#! python3*
#  Scrape bond data from Morningstar and dump into Excel

import datetime, openpyxl
from openpyxl.styles import Font, colors
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import WebDriverWait

class Bond:
    '''Bond Values'''
    def __init__(self, ticker='', row=0, yld=0.0, ytd=0.0, mtd=0.0, qtd=0.0, t1=0.0, t3=0.0, t5=0.0):
        self.index  = []
        self.ticker = ticker
        self.row    = row
        self.yld    = yld
        self.ytd    = ytd
        self.mtd    = mtd
        self.qtd    = qtd
        self.t1     = t1
        self.t3     = t3
        self.t5     = t5

# Get list of bonds or indexes from spreadsheet
def get_bonds(sheet=None, index=False):
    bonds = []
    for row in range(2, sheet.max_row + 1):
        ticker = sheet.cell(row=row, column=3).value
        if ticker and len(ticker) <= 5:
            if index:
                if sheet.cell(row=row, column=14).value == None:
                    bonds.append(Bond(ticker, row))
                    break
            else:
                if sheet.cell(row=row, column=14).value != None:
                    bonds.append(Bond(ticker, row))
                    break
    return bonds

# Get bond values from Morningstar
def get_values(browser=None, bonds=[]):
    for bond in bonds:
        print('Fetching data for ' + bond.ticker + '...')

        # View quote page for bond
        while True:
            try:
                browser.get('http://quotes.morningstar.com/fund/fundquote/f?t={}&culture=en_us&platform=RET&test=QuoteiFrame'.format(bond.ticker))
                element = WebDriverWait(browser, 60).until(expected_conditions.presence_of_element_located((By.CSS_SELECTOR, 'td[class="gr_table_colm21"] > span')))
                break
            except:
                print('\nError loading page. Refreshing...\n')

        # Get trailing twelve-month yield
        bond.yld = element.text.rstrip('%')
        print('TTM Yield'.ljust(29) + ':' + bond.yld.rjust(7) + '%')

        # View bond performance page 
        while True:
            try:
                browser.get('http://performance.morningstar.com/fund/performance-return.action?t={}&region=usa&culture=en_US'.format(bond.ticker))
                element = WebDriverWait(browser, 60).until(expected_conditions.presence_of_element_located((By.CSS_SELECTOR, 'a[tabname="#tabquarter"]')))
                element.click()
                element = WebDriverWait(browser, 60).until(expected_conditions.presence_of_all_elements_located((By.CSS_SELECTOR, 'div[id="tab-quar-end-content"] td[class="row_data"]')))
                break
            except:
                print('\nError loading page. Refreshing...\n')

        # Get trailing total returns for different periods
        bond.ytd = element[3].text
        print('YTD Trailing Total Return'.ljust(29)    + ':' + bond.ytd.rjust(7) + '%')
        bond.mtd = element[0].text
        print('MTD Trailing Total Return'.ljust(29)    + ':' + bond.mtd.rjust(7) + '%')
        bond.qtd = element[1].text
        print('QTD Trailing Total Return'.ljust(29)    + ':' + bond.qtd.rjust(7) + '%')
        bond.t1  = element[4].text
        print('1-Year Trailing Total Return'.ljust(29) + ':' + bond.t1.rjust(7)  + '%')
        bond.t3  = element[5].text
        print('3-Year Trailing Total Return'.ljust(29) + ':' + bond.t3.rjust(7)  + '%')
        bond.t5  = element[6].text
        print('5-Year Trailing Total Return'.ljust(29) + ':' + bond.t5.rjust(7)  + '%\n')

# Convert bond values to floats
def to_floats(bonds=[]):
    for bond in bonds:
         bond.yld = float(bond.yld)
         bond.ytd = float(bond.ytd)
         bond.mtd = float(bond.mtd)
         bond.qtd = float(bond.qtd)
         bond.t1  = float(bond.t1)
         bond.t3  = float(bond.t3)
         bond.t5  = float(bond.t5)

# Write bonds to spreadsheet
def write_bonds(sheet=None, bonds=[]):
    for bond in bonds:
        values = {7:(bond.yld / 100), 15:bond.ytd, 16:bond.mtd, 17:bond.qtd, 18:bond.t1, 19:bond.t3, 20:bond.t5}

        for column in list(values.keys()):
            sheet.cell(row=bond.row, column=column).value = values[column]
            sheet.cell(row=bond.row, column=column).font = Font(color=colors.BLACK, italic=False)

        for column in bond.index:
            sheet.cell(row=bond.row, column=column).font = Font(color=colors.BLUE, italic=True)

# Get path to spreadsheet
print('Getting spreadsheet path...')
file = open('excelpath.txt')
path = file.read()
file.close()

# Create new sheet in workbook
print('Preparing workbook...')
workbook = openpyxl.load_workbook(path)
sheet = workbook.copy_worksheet(workbook.active)
sheet.conditional_formatting = workbook.active.conditional_formatting
sheet.title = datetime.datetime.now().strftime('%b %Y')

# Get list of bonds and indexes from spreadsheet
print('Getting list of bonds...')
indexes = get_bonds(sheet, index=True)
bonds = get_bonds(sheet)

# Get bond values from Morningstar
print('Getting index bond values...\n')
browser = webdriver.Chrome()
get_values(browser, indexes)
print('Getting bond values...\n')
get_values(browser, bonds)
browser.quit()

# Fill in empty bond values with index data
print('Filling in empty bond values with index data...')
for bond in bonds:
    # Find corresponding index
    for index in indexes:
        if bond.row < index.row:
            break

    # Fill in empty bond values with index data
    if bond.yld == '':
        bond.yld = index.yld
        bond.index.append(7)
    if bond.ytd == '':
        bond.ytd = index.ytd
        bond.index.append(15)
    if bond.mtd == '':
        bond.mtd = index.mtd
        bond.index.append(16)
    if bond.qtd == '':
        bond.qtd = index.qtd
        bond.index.append(17)
    if bond.t1 == '':
        bond.t1 = index.t1
        bond.index.append(18)
    if bond.t3 == '':
        bond.t3 = index.t3
        bond.index.append(19)
    if bond.t5 == '':
        bond.t5 = index.t5
        bond.index.append(20)

# Convert index and bond values to floats
print('Converting bond values to floating point numbers...')
to_floats(indexes)
to_floats(bonds)

# Save bonds to spreadsheet
print('Saving bonds to spreadsheet...')
write_bonds(sheet, indexes)
write_bonds(sheet, bonds)
workbook.active = workbook.get_sheet_names().index(sheet.title)
path = path.split('\\')
del path[-1]
path.append('output.xlsx')
path = '\\'.join(path)
workbook.save(path)

print('Done!')