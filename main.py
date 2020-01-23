from selenium import webdriver
from selenium.webdriver.support.select import By
#from selenium import By
#from selenium.webdriver.support import WebDriverWait, expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
#from selenium.common import TimeoutException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException as no_element
from selenium.common.exceptions import StaleElementReferenceException as staleElement
#from selenium import Keys
from selenium.webdriver.common.keys import Keys
#from selenium.common import WebDriverException
from selenium.common.exceptions import WebDriverException
import os
import xlrd
import xlwt
import time


timeout = 20

browser = webdriver.Firefox(executable_path='/Users/tanner/Documents/geckodriver')
browser.get("https://www.economicmodeling.com/")

wb = xlwt.Workbook()
sheet = wb.add_sheet('Active Postings')

arr = ['Sep 2016', 'Oct 2016', 'Nov 2016', 'Dec 2016', 'Jan 2017',
'Feb 2017',
'Mar 2017',
'Apr 2017',
'May 2017',
'Jun 2017',
'Jul 2017',
'Aug 2017',
'Sep 2017',
'Oct 2017',
'Nov 2017',
'Dec 2017',
'Jan 2018',
'Feb 2018',
'Mar 2018',
'Apr 2018',
'May 2018',
'Jun 2018',
'Jul 2018',
'Aug 2018',
'Sep 2018',
'Oct 2018',
'Nov 2018',
'Dec 2018',
'Jan 2019',
'Feb 2019',
'Mar 2019',
'Apr 2019',
'May 2019',
'Jun 2019',
'Jul 2019',
'Aug 2019',
'Sep 2019',
'Oct 2019',
'Nov 2019']

xl_dict = {'Sep 2016': 2, 'Oct 2016' : 3, 'Nov 2016' : 4,
'Dec 2016':5,
'Jan 2017':6,
'Feb 2017':7,
'Mar 2017':8,
'Apr 2017':9,
'May 2017':10,
'Jun 2017':11,
'Jul 2017':12,
'Aug 2017':13,
'Sep 2017':14,
'Oct 2017':15,
'Nov 2017':16,
'Dec 2017':17,
'Jan 2018':18,
'Feb 2018':19,
'Mar 2018':20,
'Apr 2018':21,
'May 2018':22,
'Jun 2018':23,
'Jul 2018':24,
'Aug 2018':25,
'Sep 2018':26,
'Oct 2018':27,
'Nov 2018':28,
'Dec 2018':29,
'Jan 2019':30,
'Feb 2019':31,
'Mar 2019':32,
'Apr 2019':33,
'May 2019':34,
'Jun 2019':35,
'Jul 2019':36,
'Aug 2019':37,
'Sep 2019':38,
'Oct 2019':39,
'Nov 2019':40 }
sheet.write(0, 0, 'Company')
i = 1
for key in arr:
    sheet.write(0, i, key)
    i += 1

comps_book = xlrd.open_workbook('/Users/tanner/Downloads/pulls.xlsx')
comps_sheet = comps_book.sheet_by_index(2)

comps_rows = comps_sheet.nrows
j = 0
companies = []
while j < comps_rows:
    try:
        companies.append(str(comps_sheet.cell(j, 0).value))
    except UnicodeEncodeError:
        print("missed row at", j)
        companies.append('Universidade de Sao Paulo')
    j += 1


def waitForElement(selector):
    try:
        WebDriverWait(browser, timeout).until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, selector)))
    except TimeoutException:
        print("There was an unexpected error. Please re-run query")

waitForElement('.secondary-nav > li:nth-child(1) > a:nth-child(1)')
browser.find_element_by_css_selector('.secondary-nav > li:nth-child(1) > a:nth-child(1)').click()

waitForElement('#userinput')
browser.find_element_by_css_selector('#userinput').send_keys('bret.swango@colliers.com')
browser.find_element_by_css_selector('#passwordinput').send_keys('Colliers1')

waitForElement('.sn-sidebar > ul:nth-child(1) > li:nth-child(8) > a:nth-child(1)')

waitForElement('a.cc-btn:nth-child(1)')
browser.find_element_by_css_selector('a.cc-btn:nth-child(1)').click()
waitForElement('.sn-sidebar > ul:nth-child(1) > li:nth-child(8) > a:nth-child(1)')

browser.find_element_by_css_selector('.sn-sidebar > ul:nth-child(1) > li:nth-child(8) > a:nth-child(1)').click()
waitForElement('div.sn-page-boxes:nth-child(2) > a:nth-child(1) > div:nth-child(2)')
browser.find_element_by_css_selector('div.sn-page-boxes:nth-child(2) > a:nth-child(1) > div:nth-child(2)').click()




waitForElement('#gofer_pill_search_companies-container')
#browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
browser.find_element_by_css_selector('#gofer_pill_search_companies').click()
browser.find_element_by_css_selector('#gofer_pill_search_companies').send_keys('Apple Inc.')
time.sleep(2)
browser.find_element_by_css_selector('#gofer_pill_search_companies').send_keys(Keys.ENTER)
browser.find_element_by_css_selector('button.button:nth-child(9)').click()

row = 1
for company in companies:
    if row > 0:
        #waitForElement('div.gofer-option:nth-child(5) > ul:nth-child(5) > li:nth-child(1) > a:nth-child(1) > i:nth-child(2)')
        try:
            browser.execute_script("window.scrollTo(0, 500)")
            try:
                WebDriverWait(browser, timeout).until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, '#companies_clear_sidebar')))
            except TimeoutException:
                print("There was an unexpected error. Please re-run query")
            browser.find_element_by_css_selector('#companies_clear_sidebar').click()
        except no_element:
            print("no button")
        time.sleep(1)
        browser.find_element_by_css_selector('#gofer_search_companies').click()
        browser.find_element_by_css_selector('#gofer_search_companies').send_keys(company)
        time.sleep(3)
        try:
            browser.find_element_by_css_selector('div.grouped:nth-child(6) > div:nth-child(2) > div:nth-child(6) > div:nth-child(3)')
        except no_element:
            print("going to next")
            continue
        browser.find_element_by_css_selector('#gofer_search_companies').send_keys(Keys.ENTER)
        time.sleep(5)
        try:
            actual_company_name = browser.find_element_by_css_selector('div.gofer-option:nth-child(5) > ul:nth-child(5) > li:nth-child(1)').get_attribute('title')
        except no_element:
            actual_company_name = "NULL"
        waitForElement('#postings_trend > map:nth-child(3)')
        try:
            graph = browser.find_element_by_css_selector('#postings_trend > map:nth-child(3)')
        except no_element:
            continue
        areas = graph.find_elements_by_tag_name('area')
        sheet.write(row, 0, company)
        sheet.write(row, 1, actual_company_name)
        i = 1
        for area in areas:
            val = area.get_attribute('alt')
            val_arr = val.split('Active Postings: ')
            col = xl_dict[val_arr[0]]
            sheet.write(row, col, val_arr[1])

        wb.save('/Users/tanner/Documents/MedNational.xls')
    row += 1














