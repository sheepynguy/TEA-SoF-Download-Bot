# TEA SoF Download Bot
This program allows you to download the previous month's worth of Summary of Finance reports from the TEA website. It utilizes Selenium, which helps do the human task of clicking and typing on the website. It is also able to move the files to a synced OneDrive folder locally or can connect to OneDrive using a login.

## Setup
1) Download the Selenium driver for your preferred web broswer. The following line is going to change if you don't use Edge:
    - If you use Chrome
    ```
    driver = webdriver.Chrome()
    ```
2) If you don't have python installed onto the computer, do so now
3) Open the command prompt and install the following python libraries if you haven't yet:
   ```
   py -m pip install selenium
   py -m pip install requests
   py -m pip install openpyxl
   py -m pip install xls2xlsx
4) Replace the file path names in the folder_paths array with the desired file paths


## How to Use
Run the retrieve_reports.py file, and it will prompt you to open a login.microsoft.com link. Use the access code that is given inside the terminal, and log into your DSS microsoft account. After logging in, you can close the window and let the download bot run without an supervision. Do not close the browser that the bot opened on its own or let your computer go to sleep, or else the program will not run properly.

In the event you need to change to a different school year, you will need to manually change some hard-coded values.
These are the values and line placements:

    Ln 184 school_year.select_by_value("<Desired School Year>")
    Ln 149 ws2["C"+ str(last_row)].value = "<School Year Range i.e. 2023-2024>"

If you add more schools to the school_names array, make sure that is has the same placement in the folder_paths array as well.
