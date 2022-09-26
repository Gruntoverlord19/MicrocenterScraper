import importlib.util
import sys
import subprocess

def selenium_check():
    name = 'selenium'
    if name in sys.modules:
        print(f"{name!r} already in sys.modules")
    elif (spec := importlib.util.find_spec(name)) is not None:
    # If you choose to perform the actual import ...
        module = importlib.util.module_from_spec(spec)
        sys.modules[name] = module
        spec.loader.exec_module(module)
        #print(f"{name!r} has been imported")
    else:
        print(f"can't find the {name!r} module")
        print(f'Will now install {name}')
        while subprocess.check_call([sys.executable, "-m", "pip", "install", name]):
            print(f'Installing {name}')
        print(f'{name} is now installed.')

def webdriver_check():
    name = 'webdriver_manager'
    if name in sys.modules:
        print(f"{name!r} already in sys.modules")
    elif (spec := importlib.util.find_spec(name)) is not None:
        module = importlib.util.module_from_spec(spec)
        sys.modules[name] = module
        spec.loader.exec_module(module)
        #print(f"{name!r} has been imported")
    else:
        print(f"can't find the {name!r} module")
        print(f'Will now install {name}')
        while subprocess.check_call([sys.executable, "-m", "pip", "install", name]):
            print(f'Installing {name}...')
        print(f'{name} is now installed.')
        
selenium_check()
webdriver_check()
   
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time
import csv
import os

ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__), '.'))

def load_page():
    print("Welcome to my MicroCenter Scraper!\n")
    print("This Script will scrape the Sharonville Microcenter website and generate a CSV file of up to 96 results.\n")
    item_search = input("Please enter your search term (no dashes):\n")
    print("\n")
    print("Now scraping...")
    print("\n")
    options = Options()
    options.headless = True
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    global driver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=options)
    driver.get(f'https://www.microcenter.com/search/search_results.aspx?N=0&NTX=mode+MatchPartial&NTT={item_search}&NTK=all&sortby=match&rpp=96&storeid=071')
    #scroll to the bottom then back up to load the whole page
    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.END)
    time.sleep(.5)
    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.HOME)

def get_results():
    file = open((os.path.join(ROOT_DIR, "Scraped Data.csv")), "w", newline="")
    header = ["Product","Price ($)","Inventory","SKU","Link"]
    writer = csv.writer(file)
    writer.writerow(header)
    
    card = driver.find_elements(By.XPATH, "//li[@class='product_wrapper']")
    amount = len(card)
    for c in card:
        product = driver.find_elements(By.CLASS_NAME, 'product_wrapper > div.result_right > div > div.detail_wrapper > div.pDescription.compressedNormal1 > div.normal > h2 > a')
        if product == []:
            product = driver.find_elements(By.CLASS_NAME, 'product_wrapper > div.result_right > div > div.detail_wrapper > div.pDescription.compressedNormal2 > div.normal > h2 > a')
            if product == []:
                product = driver.find_elements(By.CLASS_NAME, 'product_wrapper > div.result_right > div > div.detail_wrapper > div.pDescription.compressedNormal3 > div.normal > h2 > a')
                if product == []:
                    product = driver.find_elements(By.CLASS_NAME, 'product_wrapper > div.result_right > div > div.detail_wrapper > div.pDescription.compressedNormal4 > div.normal > h2 > a')
                    
        link = driver.find_elements(By.CLASS_NAME, 'product_wrapper > div.result_right > div > div.detail_wrapper > div.pDescription.compressedNormal1 > div.normal > h2 > a')
        if link == []:
            link = driver.find_elements(By.CLASS_NAME, 'product_wrapper > div.result_right > div > div.detail_wrapper > div.pDescription.compressedNormal2 > div.normal > h2 > a')
            if link == []:
                link = driver.find_elements(By.CLASS_NAME, 'product_wrapper > div.result_right > div > div.detail_wrapper > div.pDescription.compressedNormal3 > div.normal > h2 > a')
                if link == []:
                    link = driver.find_elements(By.CLASS_NAME, 'product_wrapper > div.result_right > div > div.detail_wrapper > div.pDescription.compressedNormal4 > div.normal > h2 > a')        
                    
        price = driver.find_elements(By.CLASS_NAME, 'product_wrapper > div.result_right > div > div.price_wrapper > div.price > span')
        sku = driver.find_elements(By.CLASS_NAME, 'product_wrapper> div.result_right > div > div.detail_wrapper > p')
        inventory = driver.find_elements(By.CLASS_NAME, 'product_wrapper > div.result_right > div > div.detail_wrapper > div.stock > span.inventoryCnt')
            
    for p,l,pr,i,s in zip(product,link,price,inventory,sku):
        name = p.get_property("textContent").strip()
        #print(name)
        link = l.get_attribute("href").strip()
        #print(link)
        price = pr.get_property("textContent").strip()
        price = price[1:]
        #print(price)
        sku = s.get_property("textContent").strip()
        sku = sku.translate(str.maketrans('', '', 'SKU:'))
        #print(sku)
        inventory = i.get_property("textContent").strip()
        #print(inventory + "\n")

        data = [name,price,inventory,sku,link]
        writer.writerow(data)
    print(f"{amount} Results have been saved.\nExiting in 3 seconds...\n")
    time.sleep(3)
    driver.quit()
    
load_page()
get_results()
