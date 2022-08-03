# from IPython import get_ipython
# get_ipython().magic('reset -sf')
import requests
from selenium import webdriver
import pandas as pd
from bs4 import BeautifulSoup
from tkinter import * 
from tkinter import messagebox
import tkinter as tk
from tkinter import filedialog
# from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

def msg_box(string):
    root = Tk()
    root.withdraw()
    # stay_on_top(root)
    
    # root.geometry("300x200")
      
    w = Label(root) 
    w.pack()
    
    messagebox.showinfo("Info", string)
    root.destroy() 
PATH = "C:\Program Files (x86)\chromedriver.exe"
data = pd.DataFrame(pd.read_excel("KeepaExport-2022-06-05-ProductViewer.xlsx"))
titles = data['Title'].tolist()
images = data['Image'].tolist()
OriginalTitles = []
Titles = []
Prices = []
Links = []
OriginalImages = []
ImgLinks = []
distances = []
discounts = []
noResponse = False

try:
    # driver = webdriver.Chrome(ChromeDriverManager().install())
    driver = webdriver.Chrome(PATH)
    driver.maximize_window()
except:
    msg_box('Please Check Your Internet Connection And Try Again')
    
    
    
for i, title in enumerate(titles):
    url = 'https://www.google.com/search?q='+title+'&tbm=shop'
    url = url.replace(' ','+')
    # driver.get("http://webcache.googleusercontent.com/search?q=cache:"+url)
    driver.get(url)
    # WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR,"iframe[name^='a-'][src^='https://www.google.com/recaptcha/api2/anchor?']")))
    # WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[@id='recaptcha-anchor']"))).click()
    # break
    page_source = driver.page_source;
    soup = BeautifulSoup(page_source, 'lxml')
    mydivs = soup.find_all("div", {"class": "sh-dgr__content"})
    matches = 0
    while matches != 20:
        for div in mydivs:
            t = div.find("h4").string
            price = div.find("span",{"class": "OFFNJ"}).string
            if div.find("span",{"class": "Ib8pOd"}):
                discount = div.find("span",{"class": "Ib8pOd"}).string
            else:
                discount = "No Discount"
            link = div.find("a",{"class": "translate-content"})['href']
            link = link.replace("/url?url=", "")
            if link.startswith('/shopping'):
                link = "https://www.google.com"+link
            print(link)
            img_link = div.find("div",{"class": "ArOc1c"}).find('img')['src']
            r = requests.post(
                "https://api.deepai.org/api/image-similarity",
                data={
                    'image1': images[i],
                    'image2': img_link,
                },
                headers={'api-key': 'f4c98be1-04f0-4ccc-8e36-ebda5ca635b5'}
            )
            apiDetails = r.json()
            try:
                distance = apiDetails['output']['distance']
                if distance < 20:
                    percentage = str(100 - distance)+"%"
                    matches = matches + 1
                    OriginalTitles.append(title)
                    Titles.append(t)
                    Prices.append(price)
                    Links.append(link)
                    OriginalImages.append(images[i])
                    ImgLinks.append(img_link)
                    distances.append(percentage)
                    discounts.append(discount)
                if matches == 20:
                    break
            except:
                print("API isn't responding.")
                noResponse = True
                break
        btn = driver.find_elements(By.ID, "pnnext")
        if len(btn) > 0:
            driver.find_element(By.ID, "pnnext").click()
        else:
            break
        if matches == 20:
            break
        page_source = driver.page_source;
        soup = BeautifulSoup(page_source, 'lxml')
        mydivs = soup.find_all("div", {"class": "sh-dgr__content"})
    if noResponse:
        break
    if i==1:
        break
d = {'Original Title': OriginalTitles, 'New Title': Titles, 'Original Price': Prices, 'Discount': discounts, 'Product Link': Links, 'Original Image': OriginalImages, 'New Image': ImgLinks, 'Matching': distances}
df = pd.DataFrame(data=d)
df.to_excel("output.xlsx")