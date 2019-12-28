import openpyxl
import os
import requests
import selenium
import pdftotext
import datetime
import json
from shutil import copyfile
from html.parser import HTMLParser

# checks if the string is a number
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

def is_date(date):
    # print("is Date: ", date)
    date_format = '%m/%d/%Y'
    try:
        # print("is Date #2: ", date)
        date_obj = datetime.datetime.strptime(date, date_format)
        # print("is Date #3: ", date)
        return True
    except ValueError:
        # print("is Date #4: ", date)
        return False
    

# function that downloads the embedded pdf
def download_pdf(lnk):

    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options 
    from time import sleep

    options = webdriver.ChromeOptions()

    download_folder = "/Users/Susana/finance_report_generator"

    profile = {
               "download.default_directory": download_folder,
               "download.prompt_for_download": False,
               "download.directory_upgrade": True, 
               "plugins.always_open_pdf_externally": True}

    options.add_experimental_option("prefs", profile)

    print("Downloading file from link: {}".format(lnk))

    driver = webdriver.Chrome(executable_path='/Users/Susana/finance_report_generator/chromedriver', options = options)
    driver.get(lnk)

    filename = lnk.split("/")[4].split(".jsp")[0]
    print("File: {}".format(filename))

    print("Status: Download Complete.")
    print("Folder: {}".format(download_folder))
    # driver.switch_to.window(driver.window_handles[0])
    sleep(2)

    driver.close()

def main():
    redirect_link = ''
    class MyHTMLParser(HTMLParser):
        def __init__(self):
            HTMLParser.__init__(self)
            self.script = False
            self.numscripts = 0
            self.link = ''
        def handle_starttag(self, tag, attrs):
            if tag == 'script':
                self.script = True
        def handle_endtag(self, tag):
            if tag == 'script':
                self.script = False
        def handle_data(self, data):
            if not data.strip():
                return
            if self.script:
                self.numscripts = self.numscripts +1
                if self.numscripts == 5:
                    # print("Encountered some script data  :", data)
                    endquote_found = False
                    link_found = False
                    link = ''
                    for element in data: 
                        if element == '\'' and not endquote_found:
                            if link_found:
                                endquote_found = True
                                link_found = False
                                continue
                            link_found = True
                            continue
                        if link_found:
                            link = link + element
                    self.link = link
                    print(redirect_link)

    class MonthHTMLParser(HTMLParser):
        def __init__(self):
            HTMLParser.__init__(self)
            self.numlinks = 0
            self.link = ''
        def handle_starttag(self, tag, attrs):
            if tag == 'a':
                self.numlinks = self.numlinks + 1
                for a,b in attrs:
                    if a == 'href' and self.numlinks == 3:
                        endquote_found = False
                        link_found = False
                        link = ''
                        for element in b: 
                            if element == '\'' and not endquote_found:
                                if link_found:
                                    endquote_found = True
                                    link_found = False
                                    continue
                                link_found = True
                                continue
                            if link_found:
                                link = link + element
                        self.link = link
        def handle_endtag(self, tag):
            if tag == 'a' and self.numlinks == 3:
                pass

    # Error Checks
    path = os.getcwd()

    # Check for recent_reports folders
    if not os.path.exists(path + 'recent_reports'):
        pass

    # check for recent umcu and SOAS files
    if not os.path.isfile(path + '\\recent_reports\\umcu.xlsx'):
    #return print("File doesn't exits")
        pass
    if not os.path.isfile(path + '\\recent_reports\\soas.xlsx'):
    #return print("File doesn't exits")
        pass

    umcu = path + ''
    #copyfile(src, dst)

    # Download the bank statement from UMCU
    # Make session
    with requests.Session() as s:
        # Start at sign-in page
        signin_url = 'https://my.umcu.org/User/AccessSignin/Start'
        username = ''
        data = {'FormInstanceToken': '2DE044188F60428C92E16DA72DAAC94A', 'UsernameField': username, 'SubmitNext': 'Sign In'}
        response = s.post(signin_url, data=data, verify=True)
        print('Status code: ', response.status_code)
        print('Response URL: ' + response.url + '\n')
        # Step 2: Password page
        pass_url = 'https://my.umcu.org/User/AccessSignin/Password'
        password = ''
        data = {'FormInstanceToken': 'D1866717F71547E9B9E6817C360AD16A' ,'PasswordField': password, 'SubmitNext': 'Next'}
        response2 = s.post(pass_url, data=data, verify=True)
        print('Status code: ', response2.status_code)
        print('Response URL: ' + response2.url + '\n')
        # Step 3: Challenge Question
        challenge_url = 'https://my.umcu.org/User/AccessSignin/Challenge'
        answer = ''
        data = {'FormInstanceToken': '3DFD6FAF269340ED9D696FD23217E7F6', 'Answer': answer, 'Fizzbot': 'D8EBC11D78178119', 'Remember': 'False', 'SubmitNext': 'Next'}
        response3 = s.post(challenge_url, data=data, verify=True)
        print('Status code: ', response3.status_code)
        print('Response URL: ' + response3.url + '\n')
        # Step 4: eStatements Page
        estatement_url ='https://my.umcu.org/User/CustomSignonToEStatements/Start'
        response4 = s.get(estatement_url)
        print('Status code: ', response4.status_code)
        print('Response URL: ' + response4.url + '\n')
        # Step 5: Parse eStatement page for redirect link
        txt = response4.text
        parser = MyHTMLParser()
        parser.feed(txt)
        print("This is the estatement link: ", parser.link, '\n')
        #Step 6: eStatements redirect link
        response5 = s.get(parser.link)
        print('Status code: ', response5.status_code)
        print('Response URL: ' + response5.url + '\n')
        #Step 7: Parse estatements for current month
        txt = response5.text
        parser = MonthHTMLParser()
        parser.feed(txt)
        link = 'https://www.mystatement.org' + parser.link
        print("This is the November estatement link: ", link, "\n")
        #Step 8: PDF of November estatement

        # before = os.listdir('/home/jason/Downloads')

        # Download the file using Selenium here
        download_pdf(link)

        # after = os.listdir('/home/jason/Downloads')
        # change = set(after) - set(before)
        # if len(change) == 1:
        #     file_name = change.pop()
        # else:
        #     print("More than one file or no file downloaded")

    # Exit session and Extract data from UMCU pdf file

    # Load your PDF
    with open("mystatement.pdf", "rb") as f:
        pdf = pdftotext.PDF(f)

    # How many pages?
    print(len(pdf))

    # Iterate over all the pages
    for page in pdf:
        #print(page)
        pass

    # list of dictionaries of entries
    entries = []

    text = "\n\n".join(pdf)
    splits = text.split()
    index = 0
    total = 0
    bbc_num = 0
    start_collecting = False
    for split in splits:
        if split == 'BUSINESS' and splits[index+1] == 'BASIC' and splits[index+2] == 'CHECKING':
            if bbc_num > 0:
                start_collecting = True
            bbc_num = bbc_num +1
        
        if start_collecting and is_date(split):
            print("Date: ", split)
            print("Amount: ",splits[index+1])
            print("Balance: ", splits[index+2])
            temp = {}
            temp["data"] = split
            temp["amount"] = splits[index+1]
            temp["balance"] = splits[index+2]
            if temp["balance"] == "Ending":
                start_collecting = False
            temp_index = index+3
            temp_str = ""
            while splits[temp_index] != "Page" and splits[temp_index] != "WITHDRAWALS" and not is_date(splits[temp_index]):
                temp_str = temp_str + splits[temp_index] + ' '
                temp_index = temp_index+1
            if splits[temp_index] == "Page":
                start_collecting = False
            temp["description"] = temp_str
            print("Description: ", temp_str)
            entries.append(temp)
        index = index + 1

    print(entries)
    


    
        
        

    # Output data into Excel sheet


if __name__ == '__main__':
    main()