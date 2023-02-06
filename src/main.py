# from seleniumrequests import Firefox
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
# import undetected_chromedriver.v2 as uc
from statistics import mean 
import xlsxwriter
import requests
import time 
import pandas as pd

outWorkbook = xlsxwriter.Workbook("data.xlsx")
outSheet = outWorkbook.add_worksheet()
outSheet.write("A1","UPC")
outSheet.write("B1","ASIN")
outSheet.write("C1","Price1")
outSheet.write("D1","Price2")
df = pd.read_excel(
    'oridata.xlsx',
    engine='openpyxl',
    usecols="A,B"
)
class Main:
    ur=""
    def __init__(self,urx):
        self.ur = urx
        self.options = Options()
        self.options.add_argument("--auto-open-devtools-for-tabs")
        self.options.add_argument("--start-maximized")  
        self.options.add_argument('--headless')
        self.options.add_argument('--disable-gpu')
        # self.driver = uc.Chrome(options=self.options,version_main=98)
        self.driver = webdriver.Chrome(ChromeDriverManager().install(), options=self.options) 
        self.actions = ActionChains(self.driver)
        self.wait = WebDriverWait(self.driver, 30)
    def getPrice(self,asin):
        self.driver.get('https://www.amazon.com/dp/'+asin)
        # sPrice = self.driver.find_elements(By.XPATH,"/html/body/div[1]/div[2]/div[5]/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr/td[2]/span[1]/span[3]") 
        fPrice = self.driver.find_elements(By.XPATH,"/html/body/div[1]/div[2]/div[5]/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr/td[2]/span[1]/span[1]") 
        print("data fPrice: ",fPrice)
        for i,x in enumerate(fPrice):
            harga1 = x.text
            if harga1 == "":
                FirstPrice = self.driver.find_elements(By.XPATH,"/html/body/div[1]/div[2]/div[5]/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr/td[2]/span[1]") 
                for f,u in enumerate(FirstPrice):
                    return u.text
            elif "<selenium.webdriver.remote.webelement.WebElement" in fPrice:
                SeccondElement = self.driver.find_elements(By.XPATH,"/html/body/div[1]/div[2]/div[5]/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr/td[2]/span[1]/span[1]") 
                for f,u in enumerate(FirstPrice):
                    return u.text
            return harga1
    def getSprice(self, ASIN):
        self.driver.get('https://www.amazon.com/dp/'+ASIN)
        sPrice = self.driver.find_elements(By.XPATH,"/html/body/div[1]/div[2]/div[5]/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr/td[2]/span[1]/span[3]") 
        # fPrice = self.driver.find_elements(By.XPATH,"/html/body/div[1]/div[2]/div[5]/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr/td[2]/span[1]/span[1]") 
        print("data sPrice: ",sPrice)
        for i,x in enumerate(sPrice):
            harga2 = x.text
            if harga2 is None:
                fsPrice = self.driver.find_elements(By.XPATH,"/html/body/div[1]/div[2]/div[5]/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div[3]/div[2]/div/table/tbody/tr/td[2]/span[1]")
                for j,y in enumerate(fsPrice):
                    return y.text
            return harga2
mn = Main("B08TVDWM9W")
for i, row in df.iterrows():
    # print(row)
    rx = str(row)
    ra = str(rx.split("ASIN")[1:])
    rz = str(ra.split("\\")[0])
    ASIN = rz.split("['")[1].replace(" ", "")
    # print(result)
    rx = str(row)
    ra = str(rx.split("UPC")[1:])
    rz = str(ra.split("\\")[0])
    UPC = rz.split("['")[1].replace(" ", "")
    print(ASIN)
    lowprc=mn.getPrice(ASIN)
    hprc = mn.getSprice(ASIN)
    print(lowprc,hprc)
    outSheet.write(i+1,0,UPC)
    outSheet.write(i+1,1,ASIN)
    outSheet.write(i+1,2,lowprc)
    outSheet.write(i+1,3,hprc)
    time.sleep(10)
outWorkbook.close()
