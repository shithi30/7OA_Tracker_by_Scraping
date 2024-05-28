#!/usr/bin/env python
# coding: utf-8

## ref
# - C:\Users\Shithi.Maitra\Unilever Codes\Scraping Scripts\Packshots Daraz
# - C:\Users\Shithi.Maitra\Unilever Codes\Scraping Scripts\Packshots Daraz Grammage

# import
from selenium import webdriver
from bs4 import BeautifulSoup
from random import shuffle
import time
import requests
import warnings
from pathlib import Path
from os.path import basename
from PIL import Image
import numpy as np
import zipfile
import win32com.client

# preferences

start_time = time.time()
warnings.filterwarnings("ignore")
print("Packshots with no grammage info are displayed.")

options = webdriver.ChromeOptions()
options.add_argument("ignore-certificate-errors")

# open window
driver = webdriver.Chrome(options = options)
driver.maximize_window()

# folder(s)

folder = "Packshots Daraz"
filenames = []
output_dir = Path.cwd() / folder
output_dir.mkdir(parents=True, exist_ok=True)
folder_zip = zipfile.ZipFile(folder + ".zip", 'w')

folder_gm = "Packshots Daraz Grammage"
filenames_gm = []
output_dir_gm = Path.cwd() / folder_gm
output_dir_gm.mkdir(parents=True, exist_ok=True)
folder_gm_zip = zipfile.ZipFile(folder_gm + ".zip", 'w')

# sources
filenames_src = []
filenames_gm_src = []

# link
pg = 0
while(1): 
    pg = pg + 1
    link = "https://www.daraz.com.bd/unilever-bangladesh/?from=wangpu&lang=en&langFlag=en&page=" + str(pg) + "&pageTypeId=2&q=All-Products"
    driver.get(link)

    # scroll smooth
    y = 500
    for timer in range(0, 10):
        time.sleep(1)
        driver.execute_script("window.scrollTo(0, " + str(y) + ")")
        y = y + 500

    # soup
    soup_init = BeautifulSoup(driver.page_source, "html.parser")
    soup = soup_init.find_all("a", attrs={"id": "id-a-link"})
    shuffle(soup)

    # scrape
    img_count = len(soup)
    if img_count == 0: break 
    print("\nFetching packshots from page: " + str(pg))
    for i in range(0, img_count):

        # SKU
        try: val = soup[i].find("div", attrs={"id": "id-title"}).get_text()
        except: val = ""
        sku = "" + val
    
        # packshot
        img_link = soup[i].find("img", attrs={"id": "id-img"})["src"]    
        img_data = Image.open(requests.get(img_link, stream=True, verify=False).raw).resize((600, 600)).convert("RGB")
        filenames_src.append(img_link)

        # save
        filename = sku.replace(' ', "") + ".jpg"
        filenames.append(filename)
        filepath = str(output_dir) + "\\" + filename
        img_data.save(filepath, "JPEG")
        folder_zip.write(filepath, basename(filepath))
        
        # grammage
        gm = 1
        for x in range(590, 595):
            for y in range(450, 455):
                if img_data.getpixel((x, y)) == (255, 255, 255): 
                    gm = 0
        pix = img_data.getpixel((595, 595))
        if gm == 0 and np.argmax(list(pix)) != 1 and pix != (255, 255, 255): gm = 1
                
        # save, report
        if (gm == 0): 
            filenames_gm_src.append(img_link)
            filenames_gm.append(filename)
            img_data.save(str(output_dir_gm) + "\\" + filename, "JPEG")
            folder_gm_zip.write(filepath, basename(filepath))
            display(img_data.resize(tuple(int(s*0.30) for s in img_data.size)))
        print(str(i+1) + ". " + filename)

# close window
driver.close()
        
# zip
folder_zip.close()
folder_gm_zip.close()

# seperate
filenames_src = [f for f in filenames_src if f not in filenames_gm_src]

# email
ol = win32com.client.Dispatch("outlook.application")
olmailitem = 0x0
newmail = ol.CreateItem(olmailitem)

# subject, recipients
newmail.Subject = "Grammage Packshots - Daraz"
newmail.TO = "avra.barua@unilever.com; safa-e.nafee@unilever.com; rafid-al.mahmood@unilever.com"
newmail.BCC = "shithi30@outlook.com"

# inline
def inline(fl_src, fl_nme, idx): return f'<img src="' + fl_src[idx] + '" style="display: block; width: 97px; height: 100px; border: 1px solid"><figcaption><b>Fig</b> ' + fl_nme[idx][0:7] + "..." + '</figcaption>'

# body
newmail.HTMLbody = f'''
Dear concern,<br><br>
Packshot images (total: ''' + str(len(filenames) + len(filenames_gm)) + ''') from <a href="https://www.daraz.com.bd/unilever-bangladesh/">Daraz's UBL page</a> are scraped, per a former requirement. Some examples found in the process are displayed below:<br>
<br>
<table style="margin-left: auto; margin-right: auto">
    <tr><td>''' + inline(filenames_src, filenames, 0) + '''</td>
        <td>''' + inline(filenames_src, filenames, 1) + '''</td>
        <td>''' + inline(filenames_src, filenames, 2) + '''</td>
        <td>''' + inline(filenames_src, filenames, 3) + '''</td>
        <td>''' + inline(filenames_src, filenames, 4) + '''</td></tr>
</table>
<br>
Packshots missing grammage info (count: ''' + str(len(filenames_gm)) + ''') have been seperated using <b>image processing</b> techniques. Shown below are some instances for your ref.<br>
<br>
<table style="margin-left: auto; margin-right: auto">
    <tr><td>''' + inline(filenames_gm_src, filenames_gm, 0) + '''</td>
        <td>''' + inline(filenames_gm_src, filenames_gm, 1) + '''</td>
        <td>''' + inline(filenames_gm_src, filenames_gm, 2) + '''</td>
        <td>''' + inline(filenames_gm_src, filenames_gm, 3) + '''</td>
        <td>''' + inline(filenames_gm_src, filenames_gm, 4) + '''</td></tr>
</table>
<br>
The images are in <i>.jpg</i>, with compressed dimensions for easier portability. Please find the attached <i>.zip</i>s. This email was auto generated by <i>win32com</i>.<br><br>
Thanks,<br>
Shithi Maitra<br>
Asst. Manager, Cust. Service Excellence<br>
Unilever BD Ltd.<br>
'''

# attach
files = [str(Path.cwd()) + "\\" + folder + ".zip", str(Path.cwd()) + "\\" + folder_gm + ".zip"] 
for file in files: newmail.Attachments.Add(file)

# send
newmail.Send()

# stats
print("Total packshots found: " + str(len(filenames) + len(filenames_gm)))
elapsed_time = str(round((time.time() - start_time) / 60.00, 2))
print("Elapsed time to run script (mins): " + elapsed_time)
