# TEA-SoF-Download-Bot
This program allows you to download the previous month's worth of Summary of Finance reports from the TEA website. It utilizes Selenium, which helps do the human task of clicking and typing on the website.

## Setup
1) Download the Selenium driver for your preferred web broswer. The following line is going to change if you don't use Edge:
    - If you use Chrome
    ```
    driver = webdriver.Chrome()
    ```
2) If you don't have python installed onto the computer, do so now
3) Open the command prompt and install the following python libraries
   ```
   py -m pip install selenium
4) Replace the destination variable with the desired destination path you want the files to go on your local drive.


## How to Use
Run the retrieve_reports.py file, and it will prompt you to open a login.microsoft.com link. Use the access code that is given inside the terminal, and log into your DSS microsoft account. After logging in, you can close the window and let the download bot run without an supervision. Do not close the browser that the bot opened on its own, or else the program will not run properly.

