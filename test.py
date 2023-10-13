import xlrd
from urllib import parse
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
import time
def paperUrl(name):
    q = name
    params = {
        'q':q
    }
    params = parse.urlencode(params)
    url = "https://sc.panda321.com/scholar?" + params
    return url
def getBib(url):
    options = Options()
    options.add_argument('-headless')
    driver = webdriver.Firefox(options=options)
    driver.get(url)
    driver.find_element_by_class_name('gs_or_cit.gs_nph').click()
    time.sleep(4)
    s = driver.find_element_by_class_name('gs_citi')
    if s.text == 'BibTeX':
        hr = s.get_attribute('href')
    driver.get(hr)
    bib = driver.find_element_by_xpath("//*").text
    driver.quit()
    return bib
if __name__ == "__main__":
    readfile = xlrd.open_workbook(r"C:/Users/vincent/Desktop/1.xlsx")
    print(readfile)
    names = readfile.sheet_names()
    print(names)
    obj_sheet = readfile.sheet_by_name("Sheet1")
    print(obj_sheet)
    # col = obj_sheet.ncols
    # 获取sheet行数
    row = obj_sheet.nrows
    # 获取sheet列数
    col = obj_sheet.ncols
    print("row:", row)
    print("col:", col)
    # 获取 列数据:
    # obj_sheet.col_values(0) 文章所有的名字
    url = paperUrl(obj_sheet.col_values(0)[0])
    bib = getBib(url)
    
    
    