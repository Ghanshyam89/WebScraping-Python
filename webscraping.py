from os import sep
from numpy import amax
import openpyxl
from pathlib import Path
from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
import requests as rq

ptitles = []
imgs = []
product_prices = []
products_detailss = []

# URL = 'https://www.amazon.de/dp/000102163X'
headers={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36"}

amazon_file = Path('Amazon_Scraping.xlsx')

file_obj = openpyxl.load_workbook(amazon_file)

sheet = file_obj.active

# print(sheet["C1"].value)

urls = []

for row in sheet.iter_rows(max_row=25):
    # for cell in row:
    #     print(cell.value, end=" ")
    # print()
    # print(row[2].value, row[3].value, sep=" ")
    country = row[3].value
    asin = row[2].value

    if type(asin) == float:
        asin = int(asin)
    try:
        # Creating & Appending the url
        url = "https://www.amazon."+str(country)+"/dp/"+str(asin)
        
        # url = 'https://www.amazon.de/dp/000102163X'

        print(url)
        urls.append(url)

        page = rq.get(url,headers=headers)

        soup = BeautifulSoup(page.content,'html.parser')

        ptitle = soup.find(id="productTitle").get_text()
        img = soup.find(id="imgBlkFront")
        product_price = soup.find(class_="a-size-base a-color-price a-color-price").get_text()
        products_details = soup.find_all(class_="a-list-item")
        
        print("Product Details: \n")

        print(ptitle)
        print(img.get('src'))
        print(product_price)
        for i  in range(7, 15):
            print(products_details[i].get_text())
        
        ptitles.append(ptitle.strip())
        imgs.append(img.strip())
        product_prices.append(product_price.strip())
        products_details.append(products_details.strip())

    except AttributeError:
        print(url, "not available")
    except Exception:
        print(url, "not available")
        

df = pd.DataFrame.from_dict({'Product name': ptitles, 'Product Image Url':imgs, 'Product Price':product_prices,'Product Details': str(products_detailss)}, orient = 'index') 
df.to_excel('amazonproductdetails.xlsx', index=False, encoding='utf-8')