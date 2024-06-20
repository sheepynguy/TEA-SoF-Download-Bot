import json
import msal
import requests
import os



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
    search_url = f"https://graph.microsoft.com/v1.0/search/query"
    request_body = {
    "requests": [
        {
            "entityTypes": [
                "driveItem"
            ],
            "query": {
                "queryString": folder_name
            }
        }
    ]
}

    response = requests.post(search_url, headers={"Authorization": "Bearer " + token, "Content-Type": "application/json"}, json=request_body) 
    if response.status_code == 200:
        results = response.json().get("value")
        results = results[0]["hitsContainers"][0]["hits"]   # this parses through the response to get the array we want
    else:
        print(response.status_code, response.text)


    # this will get us the school's folder and then we can access the financials folder
    folder = results[0]
    search_url = f"https://graph.microsoft.com/v1.0/drives/{folder["resource"]["parentReference"]["driveId"]}/items/{folder["resource"]["id"]}/children"
    response = requests.get(search_url, headers={"Authorization": "Bearer " + token})
    if response.status_code == 200:
        results = response.json().get("value")
        for items in results:
            if items["name"] == "Financials":
                financials_folder = items
                break
    
    # upload the file into the financials folder
    upload_url = f"https://graph.microsoft.com/v1.0/drives/{financials_folder["parentReference"]["driveId"]}/items/{financials_folder["id"]}:/{file_name}:/content"

    file_path = f"C:/Users/{os.getlogin()}/Downloads/" + file_name
    with open(file_path , "rb") as file:
        file_content = file.read()


    response = requests.put(upload_url, headers={"Authorization": "Bearer " + result["access_token"], "Content-Type": "application/octect-stream"}, data=file_content)
    if response.status_code == 201:
        print("file uploaded successfully")
    return



# We can use this to work through different file paths if different schools don't have the same file path
#     search_url = f"https://graph.microsoft.com/v1.0/drives/{folder["resource"]["parentReference"]["driveId"]}/items/{folder["resource"]["id"]}/children"
# response = requests.get(search_url, headers={"Authorization": "Bearer " + token})
# if response.status_code == 200:
# 	results = response.json().get("value")

# # go through the file path to get the last folder
# for node in path:
# 	for items in results:
# 		if items["name"] == node:
# 			folder = items
# 			break
# 	if path.len > 1:
# 		search_url = f"https://graph.microsoft.com/v1.0/drives/{folder["parentReference"]["driveId"]}/items/{folder["id"]}/children"
# 		response = requests.get(search_url, headers={"Authorization": "Bearer " + token})
# 		if response.status_code  == 200:
# 			results = response.json().get("value")
# 		else:
# 			print(response.status_code)



# upload_url = f"https://graph.microsoft.com/v1.0/drives/{folder["parentReference"]["driveId"]}/items/{folder["id"]}:/{file_name}:/content"
