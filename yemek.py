##bu python kodu, selenium ve chromedriver ile çalışmakta, siteyi normal  kullanıcı gibi ziyaret edip, gerekli verileri parse ediyor

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import os
from bs4 import BeautifulSoup
import time, datetime
import json
import requests
import sys
import ftplib
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=1920x1080")



chrome_driver = os.getcwd() +"\\chromedriver.exe"
browser = webdriver.Chrome(chrome_options=chrome_options, executable_path=chrome_driver) #replace with .Firefox(), or with the browser of your choice
url = "http://sks.istanbulc.edu.tr/tr/yemeklistesi"
browser.get(url) #navigate to the page
time.sleep(10)
kah_buton = browser.find_element_by_xpath('//*[@id="4E00590053005F006D004C00500035005500720059003100"]/div/div/div[2]/ul/li[1]')
ogle_buton = browser.find_element_by_xpath('//*[@id="4E00590053005F006D004C00500035005500720059003100"]/div/div/div[2]/ul/li[2]')
#aksam_buton = browser.find_element_by_xpath('//*[@id="4E00590053005F006D004C00500035005500720059003100"]/div/div/div[2]/ul/li[3]')
vegan_buton = browser.find_element_by_xpath('//*[@id="4E00590053005F006D004C00500035005500720059003100"]/div/div/div[2]/ul/li[6]')
kumanya_buton = browser.find_element_by_xpath('//*[@id="4E00590053005F006D004C00500035005500720059003100"]/div/div/div[2]/ul/li[4]')

son = {}
son["yemek_liste"] = []

def kah_json_olustur():
    time.sleep(5)
    kah = browser.find_element_by_id("tab-kahvalti")
    bs = BeautifulSoup(kah.get_attribute('innerHTML'), "lxml")
    bs2 = bs.find_all('table')

    js = []
    for h in bs2:
        b = h.find_all('tr')
        try:
            yemek1 = b[1].text.split('\n')[2]
        except:
            yemek1="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek2 = b[1].text.split('\n')[3]
        except:
            yemek2="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek3 = b[1].text.split('\n')[4]
        except:
            yemek3="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek4 = b[1].text.split('\n')[5]
        except:
            yemek4="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            calori = b[2].text.replace("\n", "")
        except:
            calori="---"
            print("Oops!", sys.exc_info()[0], "occured.")

        if yemek1 == "":
            yemek1 = "---"
        if yemek2 == "":
            yemek2 = "---"
        if yemek3 == "":
            yemek3 = "---"
        if yemek4 == "":
            yemek4 = "---"

        dt = datetime.datetime.strptime(b[0].text.replace("\n", ""), '%d.%m.%Y')
        dt = dt.strftime('%Y-%m-%d %H:%M:%S')
        ta = {"tarih": dt,"ogun":"Kahvaltı","yemek1":yemek1,"yemek2":yemek2,"yemek3":yemek3,"yemek4":yemek4,"calori":calori }
        son["yemek_liste"].append(ta)
def ogle_json_olustur():
    time.sleep(5)
    kah = browser.find_element_by_id("tab-ogle")
    bs = BeautifulSoup(kah.get_attribute('innerHTML'), "lxml")
    bs2 = bs.find_all('table')
    js = []
    for h in bs2:
        b = h.find_all('tr')
        try:
            yemek1 = b[1].text.split('\n')[1]
        except:
            yemek1="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek2 = b[1].text.split('\n')[2]
        except:
            yemek2="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek3 = b[1].text.split('\n')[3]
        except:
            yemek3="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek4 = b[1].text.split('\n')[4]
        except:
            yemek4="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            calori = b[2].text.replace("\n", "")
        except:
            calori="---"
            print("Oops!", sys.exc_info()[0], "occured.")

        dt = datetime.datetime.strptime(b[0].text.replace("\n", ""), '%d.%m.%Y')
        dt = dt.strftime('%Y-%m-%d %H:%M:%S')

        if yemek1 == "":
            yemek1 = "---"
        if yemek2 == "":
            yemek2 = "---"
        if yemek3 == "":
            yemek3 = "---"
        if yemek4 == "":
            yemek4 = "---"

        ta = {"tarih": dt,"ogun":"Öğle Yemeği","yemek1":yemek1,"yemek2":yemek2,"yemek3":yemek3,"yemek4":yemek4,"calori":calori}
        son["yemek_liste"].append(ta)
def aksam_json_olustur():
    time.sleep(5)
    kah = browser.find_element_by_id("tab-ogle")
    bs = BeautifulSoup(kah.get_attribute('innerHTML'), "lxml")
    bs2 = bs.find_all('table')
    js = []
    for h in bs2:
        b = h.find_all('tr')
        try:
            yemek1 = b[1].text.split('\n')[1]
        except:
            yemek1="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek2 = b[1].text.split('\n')[2]
        except:
            yemek2="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek3 = b[1].text.split('\n')[3]
        except:
            yemek3="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek4 = b[1].text.split('\n')[4]
        except:
            yemek4="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            calori = b[2].text.replace("\n", "")
        except:
            calori="---"
            print("Oops!", sys.exc_info()[0], "occured.")

        if yemek1 == "":
            yemek1 = "---"
        if yemek2 == "":
            yemek2 = "---"
        if yemek3 == "":
            yemek3 = "---"
        if yemek4 == "":
            yemek4 = "---"

        dt = datetime.datetime.strptime(b[0].text.replace("\n", ""), '%d.%m.%Y')
        dt = dt.strftime('%Y-%m-%d %H:%M:%S')
        ta = {"tarih": dt,"ogun":"Akşam Yemeği","yemek1":yemek1,"yemek2":yemek2,"yemek3":yemek3,"yemek4":yemek4,"calori": calori}
        son["yemek_liste"].append(ta)
def vegan_json_olustur():
    time.sleep(5)
    kah = browser.find_element_by_id("tab-vegan")
    bs = BeautifulSoup(kah.get_attribute('innerHTML'), "lxml")
    bs2 = bs.find_all('table')
    for h in bs2:
        b = h.find_all('tr')
        try:
            yemek1 = b[1].text.split('\n')[1]
        except:
            yemek1="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek2 = b[1].text.split('\n')[2]
        except:
            yemek2="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek3 = b[1].text.split('\n')[3]
        except:
            yemek3="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek4 = b[1].text.split('\n')[4]
        except:
            yemek4="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            calori = b[2].text.replace("\n", "")
        except:
            calori="---"
            print("Oops!", sys.exc_info()[0], "occured.")

        if yemek1 == "":
            yemek1 = "---"
        if yemek2 == "":
            yemek2 = "---"
        if yemek3 == "":
            yemek3 = "---"
        if yemek4 == "":
            yemek4 = "---"

        dt = datetime.datetime.strptime(b[0].text.replace("\n", ""), '%d.%m.%Y')
        dt = dt.strftime('%Y-%m-%d %H:%M:%S')
        ta = {"tarih": dt,"ogun":"Vegan","yemek1":yemek1,"yemek2":yemek2,"yemek3":yemek3,"yemek4":yemek4,"calori": calori}
        son["yemek_liste"].append(ta)
def kumanya_json_olustur():
    time.sleep(5)
    kah = browser.find_element_by_id("tab-kumanya")
    bs = BeautifulSoup(kah.get_attribute('innerHTML'), "lxml")
    bs2 = bs.find_all('table')
    for h in bs2:
        b = h.find_all('tr')
        try:
            yemek1 = b[1].text.split('\n')[1]
        except:
            yemek1="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek2 = b[1].text.split('\n')[2]
        except:
            yemek2="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek3 = b[1].text.split('\n')[3]
        except:
            yemek3="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            yemek4 = b[1].text.split('\n')[4]
        except:
            yemek4="---"
            print("Oops!", sys.exc_info()[0], "occured.")
        try:
            calori = b[2].text.replace("\n", "")
        except:
            calori="---"
            print("Oops!", sys.exc_info()[0], "occured.")

        if yemek1 == "":
            yemek1 = "---"
        if yemek2 == "":
            yemek2 = "---"
        if yemek3 == "":
            yemek3 = "---"
        if yemek4 == "":
            yemek4 = "---"

        dt = datetime.datetime.strptime(b[0].text.replace("\n", ""), '%d.%m.%Y')
        dt = dt.strftime('%Y-%m-%d %H:%M:%S')
        ta = {"tarih": dt,"ogun":"Öğle Yemeği","yemek1":yemek1,"yemek2":yemek2,"yemek3":yemek3,"yemek4":yemek4,"calori": calori}
        son["yemek_liste"].append(ta)
def dosya_olsutur():
    with open('yemek.json', 'w') as outfile:
        json.dump(son, outfile)




def mysql_isleri():
    requests.get("*****")


def ftp_yukle():
    print("----------------")
    print(" ")
    print("ftp deneniyor...")
    import ftplib
    ftp = ftplib.FTP()
    host = "****"
    port = 21
    ftp.connect(host, port)
    print(ftp.getwelcome())
    File2Send = "yemek.json"
    Output_Directory = "//****//"
    try:
        print("Giriş Yapılıyor...")
        ftp.login("****", "****")

        time.sleep(6)
        mysql_isleri()
        print("Başarılı")
    except Exception as e:
        print(e)
    try:
        file = open('yemek.json', 'rb')  # file to send
        ftp.storbinary('STOR yemek.json', file)  # send the file

    except Exception as e:
        print(e)
    ftp.quit()
    print(" ")
    print("----------------")
def sonuc():
    dosya_olsutur()
    ftp_yukle()


try:
    kah_buton.click()
    kah_json_olustur()
except:
    print("Kahvaltı oluşturalamadı", sys.exc_info()[0])
try:
    ogle_buton.click()
    ogle_json_olustur()

except:
    print("Öğle oluşturalamadı", sys.exc_info()[0])
try:
    ogle_buton.click()
    aksam_json_olustur()

except:
    print("Akşam oluşturalamadı", sys.exc_info()[0])

try:
    vegan_buton.click()
    vegan_json_olustur()

except:
    print("Vegan oluşturalamadı", sys.exc_info()[0])

try:
    kumanya_buton.click()
    kumanya_json_olustur()

except:
    print("Kumanya oluşturalamadı", sys.exc_info()[0])





browser.close()
print(json.dumps(son))
print("-------------")
sonuc()

time.sleep(5)
mysql_isleri()

print("-----Güncelleme Bitti----")




