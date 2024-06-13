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
