import os           # library used to interface with the os
import shutil       # library used to move the files into a different directory
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
import time
import datetime
import calendar
import connect_onedrive as OD
import requests



school_names = [
    # after using the names, splice out the number which is 9 character with the space, and then append the .pdf to the end of the string
    "A+ ACADEMY (057829)",
    "GEORGE I SANCHEZ CHARTER (101804)",
    "ADVANTAGE ACADEMY (057806)",
    "ARLINGTON CLASSICS ACADEMY (220802)",
    "CITYSCAPE SCHOOLS (057841)",
    "CUMBERLAND ACADEMY (212801)",
    "GOLDEN RULE CHARTER SCHOOL (057835)",
    "IDEA PUBLIC SCHOOLS (108807)",
    "IMAGINE INTERNATIONAL ACADEMY OF NORTH TEXAS (043801)",
    "LEADERSHIP PREP SCHOOL (061804)",
    "LEGACY PREPARATORY (057846)",
    "LONE STAR LANGUAGE ACADEMY (043802)",      # This is supposed to be the Imagine Lone Star Language Academy, but for some reason it is listed as the lone star language academy
    "MANARA ACADEMY (057844)",
    "MEYERPARK CHARTER (101855)",
    "THE PRO-VISION ACADEMY (101868)",
    "PIONEER TECHNOLOGY & ARTS ACADEMY (057850)",
    "SAN ANTONIO PREPARATORY SCHOOLS (015840)",
    "ST MARY'S ACADEMY CHARTER SCHOOL (013801)",
    "TRINITY BASIN PREPARATORY (057813)",
    "TRIVIUM ACADEMY (061805)",
    "UME PREPARATORY ACADEMY (057845)",
    "VILLAGE TECH SCHOOLS (057847)",
    "WINFREE ACADEMY CHARTER SCHOOLS (057828)",
    "NOVA ACADEMY (057809)",
    "NYOS CHARTER SCHOOL (227804)"
                ]

file_names =  [
    # these will change given the actual file names, so these aren't permanent
    "Gene Zhu - A+ Academy",
    "Gene Zhu - George I Sanchez AAMA",    # Need to ask about whether this is AAMA - GIS or just AAMA
    "Gene Zhu - Advantage Academy",
    "Gene Zhu - ACA Team Folder",
    "Cityscape Schools",    #
    "Gene Zhu - Cumberland Academy",
    "Gene Zhu - Golden Rule",
    "Gene Zhu - IDEA Public Schools",
    "Gene Zhu - Imagine",
    "Gene Zhu - Leadership Prep School",
    "Gene Zhu - Legacy Prep Charter Academy",
    "Gene Zhu - Lone Star Language Academy",
    "Manara Academy",   #
    "Meyerpark Charter",    #
    "Pro-vision Academy",   #
    "Gene Zhu - PTAA",
    "Gene Zhu - San Antonio Prep",
    "St. Mary's Academy Charter SChool",    #
    "Gene Zhu - Trinity Basin Prep - TBP",
    "Gene Zhu - Trivium Academy",
    "Gene Zhu - UME Prep",
    "Gene Zhu - Village Tech Schools",
    "Winfree Academy Charter Schools"   #
    "Gene Zhu - Nova Academy",
    "Gene Zhu - NYOS Charter School"
]


# downloads the files based on the rows listed on the index array
def download_multiple_files(rows, index, school, drive, folder):
    # iterates through each row marked by index to be searched and downloaded
    for line in index:
        row = rows[-(line)].find_elements(By.XPATH, ".//td")
        link = row[5].find_element(By.XPATH, ".//a")
        link.click()

        # wait for file to download and rename it, but don't close the window
        old_name = r"C:\Users\Victoria Nguyen\Downloads\report.pdf" # os.environ['USERPROFILE'] is the alternative for %USERPROFILE%
        while not os.path.exists(old_name):
            time.sleep(2)
        
        # get the written out date the file was uploaded in Month Day, Year format
        char_length = len(school)
        month = calendar.month_name[int((row[1].text)[:2].strip("/"))]
        day = ""
        year = ""
        if (row[1].text)[1] == "/":
            day = (row[1].text)[2:4].strip("/")
            year = (row[1].text)[4:9].strip("/ ")
        else:
            day = (row[1].text)[3:5].strip("/")
            year = (row[1].text)[5:10].strip("/ ")

        # rename the file to include the school name and date the report was uploaded
        new_name = '\\' + school[:char_length - 9] + " " + month + " " + day + ", " + year + ".pdf"
        os.rename(old_name, r"C:\Users\Victoria Nguyen\Downloads" + new_name)
        # grabs a new access token just in case the previous one has expired, and uploads the file to the specified path on OneDrive
        drive = OD.refresh_access_token()

        OD.upload_file(new_name[1:], "2024 Cost Project - Victoria", drive)

        # drive.upload_item(file_path=r"C:\Users\Victoria Nguyen\Downloads" + new_name, item_path=f"https://graph.microsoft.com/v1.0/me/drive" + "\\SoF Reports" + "\\" + new_name)
        # deletes the file from the downloads folder so that there is only a copy on the cloud and not on the local drive
        os.remove(r"C:\Users\Victoria Nguyen\Downloads" + new_name)




    return






# creates an instance of the Edge driver
driver = webdriver.Edge()
#  navigates to the TEA page to find the SoF reports
driver.get("https://tealprod.tea.state.tx.us/fsp/Reports/ReportSelection.aspx")

drive = OD.login_onedrive()

# iterate through every school to pull reports from their table
for j in range(len(school_names)):
    # selects the Summary of Finances for the first drop down
    form_type = Select(driver.find_element(By.ID, "ctl00_Body_ReportTypeDropDownList"))
    form_type.select_by_value('SummaryOfFinance')
    # clicks the submit button to move onto the next two forms
    button = driver.find_element(By.ID, "ctl00_Body_SelectButton")  # use this same id for the next select button
    button.click()

    # selects the school year for the second drop down
    school_year = Select(driver.find_element(By.ID, "ctl00_Body_SchoolYearDropDownList"))
    school_year.select_by_value(datetime.datetime.now().strftime("%Y"))

    # inputs the school name and presses submit
    name_input = driver.find_element(By.ID, "ctl00_Body_DistrictIdTextBox")
    name_input.send_keys(school_names[j])
    name_input.send_keys(Keys.ENTER)

    # Grab the latest row and check if it was the last month's report
    table = driver.find_element(By.ID, "ctl00_Body_SofDistrictRunGridView")
    rows = table.find_elements(By.XPATH, ".//tr")


    i = 1
    # for the upcoming school year, you'll need to put in some kind of input so that it knows to switch between different school years
    month_today = datetime.datetime.now().strftime("%m")
    if month_today[0] == "0":
        month_today = month_today.strip("0")

    # get the negative index of the rows that have the previous month
    index = []
    for i in range(len(rows)):
        row = rows[-(i+1)].find_elements(By.XPATH, ".//td")
        month_web = row[1].text
        if month_web[1] == "/":
            month_web = month_web[:1]
        else:
            month_web = month_web[:2]
        
        # if the previous month shows up, then we add it to the index array
        if int(month_today)-1 == int(month_web):
            index.append(i+1)
        elif int(month_today)-1 > int(month_web):
            break

    download_multiple_files(rows, index, school_names[j], drive, file_names[j])

    # after downloading the reports, return back to the drop down page to repeat the process for the next school
    reset = driver.find_element(By.ID, "ctl00_Body_SofDistrictRunCancelButton")
    reset.click()
    



