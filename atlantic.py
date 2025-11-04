from selenium import webdriver
from selenium.webdriver.common.by import By
import xlsxwriter
import time

def record_year(driver, workbook, year):
    HEADER = ['Headline', 'Link', 'Date', 'Summary']
    BOLD = workbook.add_format({'bold':True})

    worksheet = workbook.add_worksheet(str(year))
    worksheet.write_row(0, 0, HEADER, BOLD)
    cur_row = 1

    months = ['01', '03', '04', '05', '06', '07', '09', '10', '11', '12']

    for month in months:
        driver.get(f"https://www.theatlantic.com/magazine/toc/{year}/{month}/")
        elements = driver.find_elements(By.XPATH, "//a[contains(@class,'GridItem_hedLink')]")
        articles = {}
        for element in elements:
            articles[element.get_attribute("href")] = (element.text, element.get_attribute("href"), str(year) + "-" + month)
        cur_row = check_articles(driver, worksheet, cur_row, articles)

def check_articles(driver, worksheet, cur_row, articles):
    WRAP = workbook.add_format()
    WRAP.set_shrink()
    WRAP.set_font_name('Calibri')
    WRAP.set_font_size(12)
    for link in articles:
        driver.get(link)
        if len(driver.find_elements(By.XPATH, "//*[contains(text(), 'friend')]")) == 0:
            continue
        abstract = ''
        try:
            element = driver.find_element(By.XPATH, "//p[contains(@class,'ArticleDek_feature')]")
            abstract = element.text
        except:
            pass
        worksheet.write_row(cur_row, 0, articles[link]+(abstract,), WRAP)
        cur_row += 1

    return cur_row

def sign_in(driver):
    email = "jkrems@gmail.com"
    password = open("atlantic.key", "r").read()

    driver.get('https://accounts.theatlantic.com/login/?redirect=%2F')
    username_input = driver.find_element(By.XPATH, "//input[@id='username']")
    username_input.send_keys(email)
    button = driver.find_element(By.XPATH, "//button[contains(text(), 'Continue')]")
    button.click()
    time.sleep(1)

    password_input = driver.find_element(By.XPATH, "//input[@id='password']")
    password_input.send_keys(password)
    button = driver.find_element(By.XPATH, "//button[contains(text(), 'Sign In')]")
    button.click()
    time.sleep(1)

driver = webdriver.Firefox()
sign_in(driver)

workbook = xlsxwriter.Workbook("atlantic.xlsx")
for year in range(2015, 2025, 1):
    record_year(driver, workbook, year)
workbook.close()

driver.get('https://www.theatlantic.com/magazine/archive/2015/01/the-tragedy-of-the-american-military/383516/')
driver.quit()

