from bs4 import BeautifulSoup
import requests
import concurrent.futures
import time
from workbook import sheet,newSheet,newWorkbook,dest_filename

baseUrl = "https://www.spx.com.tr" 
urls = []

#Gets urls from sheet
#Add them to an array
def getUrls(): 
    for i in range(1,101):
        urls.append(baseUrl + sheet["A"+str(i)].value)

#Scrapes urls
#Writes data to new sheet   
def scrap(url):
    global urls
    rowNumber = str(urls.index(url)+2) #Row number of url

    page = requests.get(url)
    time.sleep(0.15) #Sleeps 0.15 seconds
    soup = BeautifulSoup(page.text, 'html.parser')

    price = soup.find("div","product__price").text.strip() #Gets price of the product
    code = soup.find("div","product__code").text.strip()   #Gets code of the product
    brand = soup.find("a","product__brand").text.strip()   #Gets brand of the product
    name = soup.find("h1","product__title").text.strip()   #Gets name of the product

    sizes = soup.find_all("a","product__size-variant")     #Gets sizes of the product
    available = 0 #Available sizes count

    for size in sizes:
        if "disabled" not in size["class"]: #If size does not have a disabled tag
            available += 1                  #Increase available sizes count
    
    sizesLength = len(sizes)

    if sizesLength != 0:
        percent = round(available/sizesLength * 100)
    else:
        percent = 100

    #Write values to associated cells
    newSheet["A"+rowNumber] = url 
    newSheet["B"+rowNumber] = brand 
    newSheet["C"+rowNumber] = name
    newSheet["D"+rowNumber] = code
    newSheet["E"+rowNumber] = price
    newSheet["F"+rowNumber] = "%" + str(percent)

getUrls()

with concurrent.futures.ThreadPoolExecutor() as executor:
        executor.map(scrap, urls)

newWorkbook.save(filename = dest_filename) #Saving new xlsx file