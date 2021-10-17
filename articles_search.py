### https://github.com/MazarsLabs/hse-rpa

import os
import pandas as pd
from selenium.webdriver.chrome.options import Options
import selenium.webdriver as webdriver
import time
from wcm import get_credentials
from email.message import EmailMessage
import smtplib
from urllib.request import urlretrieve
import uuid
import wget
from conf import query, num_page, receiver

query_link = f"https://www.semanticscholar.org/search?q={query}&sort=relevance&pdf=true&page="
# working paths
working_dir = os.path.dirname(os.path.realpath(__file__))
folder_for_pdf = os.path.join(working_dir, "articles")
webdriver_path = os.path.join(working_dir, "chromedriver")   # proper version https://chromedriver.chromium.org/

# chek if articles directory is exist and create if not
if not os.path.isdir(folder_for_pdf):
    os.mkdir(folder_for_pdf)
login, password = get_credentials("hse")   # could be setup manually

# webdriver
chrome_options = Options()
prefs = {"plugins.always_open_pdf_externally": True, "plugins.plugins_list": [{"enabled": False, "name": "Chrome PDF Viewer"}], "download.default_directory": folder_for_pdf, "download.prompt_for_download": False}

chrome_options.add_experimental_option('prefs', prefs)
os.environ["webdriver.chrome.driver"] = webdriver_path   # 'webdriver' executable needs to be in PATH. Please see https://sites.google.com/a/chromium.org/chromedriver/home

links_list = [query_link + str(page+1) for page in range(num_page)]   # create links to follow

driver = webdriver.Chrome(executable_path=webdriver_path, chrome_options=chrome_options)

final_info = []   # empty dictionary for articles info
for search_link in links_list:
    # get all links to articles from the page
    driver.get(search_link)
    time.sleep(5)
    articles = driver.find_elements_by_class_name("cl-paper-row")
    
    links = driver.find_elements_by_xpath("//a[@data-selenium-selector='title-link']")
    articles_links = []

    for l in links:
        articles_links.append(l.get_attribute("href"))

    for link in articles_links:
        # get info of each article 
        tmp_info = {}

        driver.get(link)
        try:
            toggler = driver.find_element_by_xpath("//a[@data-selenium-selector='text-truncator-toggle']")
            if(toggler.get_attribute("aria-label") == 'Expand truncated text'):
                toggler.click()
        except:
            pass

        try:
            toggler = driver.find_element_by_class_name("more-authors-label")
            toggler.click()
        except:
            pass

        
        title = driver.find_element_by_xpath("//h1[@data-selenium-selector='paper-detail-title']").text
        abstract = ""
        try:
            abstract = driver.find_element_by_xpath("//span[@data-selenium-selector='text-truncator-text']").text
        except:
            pass
        authors = []
        for el in driver.find_elements_by_class_name("author-list__author-name"):
            authors.append(el.find_element_by_tag_name("span").text)
        publishDate = driver.find_element_by_xpath("//span[@data-selenium-selector='paper-year']/span/span").text
        citations = ''
        try:
            citations = driver.find_element_by_class_name("dropdown-filters__result-count__header").text[:-10]
        except:
            pass
        tmp_info.update({
                        'title': title,
                        'abstract': abstract,
                        'date' : publishDate,   # TODO: might convert to datetime
                        'authors': ', '.join(authors),
                        'citations': citations
                        })

        # trying to download the article's doc
        try:
            initial_dir = os.listdir(folder_for_pdf)
            #container = driver.find_element_by_class_name("flex-paper-actions__item-container")
            driver.find_element_by_xpath("//a[@data-heap-direct-pdf-link='true']").click()
            time.sleep(5)
            #filePath = _filePath.get_attribute("href")
            #filename = str(uuid.uuid4()) + ".pdf"
            current_dir = os.listdir(folder_for_pdf)

            filename = list(set(current_dir) - set(initial_dir))[0]

            full_path = os.path.join(folder_for_pdf, filename)
            #wget.download(filePath, full_path)


        except Exception as e:
            full_path = None
            filePath = None

        tmp_info.update(
            {
                'path_to_file':full_path
            })
        if(full_path != None):
            final_info.append(tmp_info.copy())
        time.sleep(2)

driver.quit()


# write all info to excel
df = pd.DataFrame(final_info)
excel_path = os.path.join(working_dir, "data.xlsx")
df.to_excel(excel_path, index=False)

# create email
mail = EmailMessage()
mail['From'] = login
mail['To'] = receiver
mail['Subject'] = "Topics analysis"
mail.set_content("Hi!\n\nFind attached excel file with articles info.\n\nRegard,")

# add attachment
with open(excel_path, 'rb') as f:
    file_data = f.read()
    file_name = f'articles_info.xlsx'
mail.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

# send email    
server = smtplib.SMTP('smtp.office365.com')  
server.starttls()  
server.login(login, password)    
server.send_message(mail)      
server.quit()      


    # ┏━┓┏━┓┏━━━┓┏━━━━┓┏━━━┓┏━━━┓┏━━━┓  
    # ┃ ┗┛ ┃┃┏━┓┃┗━━┓ ┃┃┏━┓┃┃┏━┓┃┃┏━┓┃  
    # ┃┏┓┏┓┃┃┃ ┃┃  ┏┛┏┛┃┃ ┃┃┃┗━┛┃┃┗━━┓  
    # ┃┃┃┃┃┃┃┗━┛┃ ┏┛┏┛ ┃┗━┛┃┃┏┓┏┛┗━━┓┃  
    # ┃┃┃┃┃┃┃┏━┓┃┏┛ ┗━┓┃┏━┓┃┃┃┃┗┓┃┗━┛┃  
    # ┗┛┗┛┗┛┗┛ ┗┛┗━━━━┛┗┛ ┗┛┗┛┗━┛┗━━━┛  
