import json
import sys
import msal
from msdrive import OneDrive
import requests



# make a call to the API that we want to connect to OneDrive
#app = msal.PublicClientApplication(authority=AUTHORITY_URL, client_id=CLIENT_ID)

app = None
result = None

def login_onedrive():
    AUTHORITY_URL = "https://login.microsoftonline.com/088fd6f9-1d62-44b5-87b9-64b9699ac96a/"
    CLIENT_ID = "ee6ef894-12c7-489f-b8d0-f60f6a79a799"

    global app
    app = msal.PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY_URL)

    # print out the message to login to your DSS account
    flow = app.initiate_device_flow(scopes=["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"])
    print(flow["message"])

    # get the access key and return it back
    global result
    result = app.acquire_token_by_device_flow(flow)

    return result["access_token"]



# this returns a new instance of logging into OneDrive to upload the files to the cloud
def refresh_access_token():
    global result
    result = app.acquire_token_by_refresh_token(result["refresh_token"], scopes=["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"])
    if "access_token" in result:
        return result["access_token"]
    else:
        print(result.get("error"))



# this function will take the file name and folder name parameters and upload the file to the designated folder
def upload_file(file_name, folder_name, token):
    # make a GET request to the API to look through OneDrive to find the folder
    search_url = f"https://graph.microsoft.com/v1.0/me/drive/search(q=\'{folder_name}\')"
    response = requests.get(search_url, headers={"Authorization": "Bearer " + token})

    # if the request is successful, then it will get all the results from the search and return it back. The first result is the only one we will look at
    folder = None
    if response.status_code == 200:
        results = response.json().get("value", [])
        folder = results[0]

    # use the folder's ID to upload the file to that folder
    file_path = r"C:\Users\Victoria Nguyen\Downloads\\" + file_name
    upload_url = f"https://graph.microsoft.com/v1.0/drives/{folder["parentReference"]["driveId"]}/items/{folder["id"]}:/{file_name}:/content"   # not sure if this url will work
    headers = {"Authorization": "Bearer " + token,
               "Content-Type": "application/octet-stream"}
    
    with open(file_path, "rb") as file:
        file_content = file.read()

    # make the PUT request to upload the file to the OneDrive folder
    response = requests.put(upload_url, headers=headers, data=file_content)

    if response.status_code == 201:
        print('File uploaded successfully')
    else:
        print(response.status_code, response.text)
    return