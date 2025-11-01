from selenium import webdriver
from selenium.webdriver.common.by import By
import xlsxwriter

def record_year(driver, workbook, year):
    HEADER = ['Headline', 'Link', 'Date', 'Summary']
    BOLD = workbook.add_format({'bold':True})
    WRAP = workbook.add_format()
    WRAP.set_shrink()

    worksheet = workbook.add_worksheet(str(year))
    worksheet.write_row(0, 0, HEADER, BOLD)
    cur_row = 1

    months = ['01', '03', '04', '05', '06', '07', '09', '10', '11', '12']

    for month in months:
        driver.get(f"https://www.theatlantic.com/magazine/toc/{year}/{month}/")
        elements = driver.find_elements(By.XPATH, "//a[contains(@class,'GridItem_hedLink')]")

        for element in elements:
            link = element.get_attribute("href")
            row_to_write = (element.text, element.get_attribute("href"), str(year) + "-" + month)
            worksheet.write_row(cur_row, 0, row_to_write, WRAP)
            cur_row += 1

def check_article(driver, link):
    driver.get(link)
    driver.back()

driver = webdriver.Firefox()
'''
workbook = xlsxwriter.Workbook("atlantic.xlsx")
for year in range(2015, 2025, 1):
    record_year(driver, workbook, year)
workbook.close()
'''
driver.get('https://www.theatlantic.com/magazine/archive/2015/01/the-tragedy-of-the-american-military/383516/')
driver.quit()

