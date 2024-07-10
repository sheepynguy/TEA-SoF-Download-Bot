import os           # library used to interface with the os
import shutil
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
import time
import calendar


# need to ask Dillon if we wants every school's report or just the ones that he wanted

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


# holds where the file is going to be relocated. If you want it to 
folder_paths =  [
    # these will change given the actual file names, so these aren't permanent
    "Gene Zhu - A+ Academy/Financials",
    "Gene Zhu - George I Sanchez AAMA",    # Need to ask about whether this is AAMA - GIS or just AAMA
    "Gene Zhu - Advantage Academy/Financials",
    "Gene Zhu - ACA Team Folder/Financials",
    "Gene Zhu - Cityscape Schools",    #
    "Gene Zhu - Cumberland Academy/Financials",
    "Gene Zhu - Golden Rule",
    "Gene Zhu - IDEA Public Schools",
    "Gene Zhu - Imagine",
    "Gene Zhu - Leadership Prep School",
    "Gene Zhu - Legacy Prep Charter Academy",
    "Gene Zhu - Lone Star Language Academy",
    "Gene Zhu - Manara Academy",   #
    "Gene Zhu - Meyerpark Charter",    #
    "Gene Zhu - Pro-vision Academy",   #
    "Gene Zhu - PTAA",
    "Gene Zhu - San Antonio Prep",
    "Gene Zhu - St. Mary's Academy Charter School",    #
    "Gene Zhu - Trinity Basin Prep - TBP",
    "Gene Zhu - Trivium Academy",
    "Gene Zhu - UME Prep",
    "Gene Zhu - Village Tech Schools",
    "Gene Zhu - Winfree Academy Charter Schools"   #
    "Gene Zhu - Nova Academy",
    "Gene Zhu - NYOS Charter School"
]

# creates an instance of the Edge driver
driver = webdriver.Edge()
#  navigates to the TEA page to find the SoF reports
driver.get("https://tealprod.tea.state.tx.us/fsp/Reports/ReportSelection.aspx")

for i in range(len(school_names)):
    # selects the Summary of Finances for the first drop down
    form_type = Select(driver.find_element(By.ID, "ctl00_Body_ReportTypeDropDownList"))
    form_type.select_by_value('SummaryOfFinance')
    # clicks the submit button to move onto the next two forms
    button = driver.find_element(By.ID, "ctl00_Body_SelectButton")  # use this same id for the next select button
    button.click()

    # selects the school year for the second drop down
    school_year = Select(driver.find_element(By.ID, "ctl00_Body_SchoolYearDropDownList"))
    school_year.select_by_value("2024")

    # inputs the school name and presses submit
    name_input = driver.find_element(By.ID, "ctl00_Body_DistrictIdTextBox")
    name_input.send_keys(school_names[i])
    name_input.send_keys(Keys.ENTER)

    # Grab the latest row and check if it was the last month's report
    table = driver.find_element(By.ID, "ctl00_Body_SofDistrictRunGridView")
    rows = table.find_elements(By.XPATH, ".//tr")
    components = rows[-1].find_elements(By.XPATH, ".//td")
    link = components[5].find_element(By.XPATH, ".//a")
    link.click()

    # wait for the file to finish downloading
    old_name = f"C:/Users/{os.getlogin()}/Downloads/report.pdf"
    while not os.path.exists(old_name):
        time.sleep(2)

    char_length = len(school_names[i])
    month = calendar.month_name[int((components[1].text)[:2].strip("/"))]
    day = ""
    year = ""
    if (components[1].text)[1] == "/":
        day = (components[1].text)[2:4].strip("/")
        year = (components[1].text)[4:9].strip("/ ")
    else:
        day = (components[1].text)[3:5].strip("/")
        year = (components[1].text)[5:10].strip("/ ")

        # rename the file to include the school name and date the report was uploaded
    new_name = school_names[i][:char_length - 9] + " SoF " + month + " " + day + ", " + year + ".pdf"
    os.rename(old_name, f"C:/Users/{os.getlogin()}/Downloads/" + new_name)
        # moves the file to a shared and synced onedrive folder
    shutil.move(f"C:/Users/{os.getlogin()}/Downloads/" + new_name, f"C:/Users/{os.getlogin()}/Dynamic Support Solutions/{folder_paths[i]}/{new_name}")
    reset = driver.find_element(By.ID, "ctl00_Body_SofDistrictRunCancelButton")
    reset.click()