import os
import requests
import msal
import json
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

client_id = '1abf6ce4-719e-4baf-8e70-27581360765b'
client_secret = 'XhB8Q~m2Onp2IAE_mruYkjtZcOOpWljR5JTpibtE'
tenant_id = '23228c04-8dbc-42f6-94cd-c658e5f50005'
authority = f'https://login.microsoftonline.com/{tenant_id}'
scope = ['https://graph.microsoft.com/.default']
sharepoint_site_url = 'esurgiinc.sharepoint.com'
folder_path = '/Products/Biostabilizer/Engineering/Software App Development/Database/Database Backups'

firebase_url = 'https://esurgiinc-default-rtdb.firebaseio.com/.json'

date_str = datetime.now().strftime("%Y-%m-%d")
file_name = f'firebaseData_{date_str}.json'

user_home_directory = os.path.expanduser("~")
file_path = os.path.join(user_home_directory, 'Automate', file_name)

def get_access_token():
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=scope)
    if 'access_token' in result:
        return result['access_token']
    else:
        raise Exception(f"Failed to obtain access token: {result.get('error_description')}")

def get_site_id_and_drive_id(access_token):
    headers = {
        'Authorization': f'Bearer {access_token}',
    }

    site_response = requests.get(f'https://graph.microsoft.com/v1.0/sites/{sharepoint_site_url}:/sites/SharedFiles', headers=headers)
    if site_response.status_code == 200:
        site_data = site_response.json()
        site_id = site_data['id']
    else:
        raise Exception(f"Failed to get site ID: {site_response.status_code} - {site_response.text}")

    drive_response = requests.get(f'https://graph.microsoft.com/v1.0/sites/{site_id}/drive', headers=headers)
    if drive_response.status_code == 200:
        drive_data = drive_response.json()
        drive_id = drive_data['id']
    else:
        raise Exception(f"Failed to get drive ID: {drive_response.status_code} - {drive_response.text}")

    return site_id, drive_id

def download_firebase_data():
    response = requests.get(firebase_url)
    if response.status_code == 200:
        data = response.json()
        os.makedirs(os.path.dirname(file_path), exist_ok=True)  
        with open(file_path, 'w') as file:
            json.dump(data, file, indent=4)
        print(f"Firebase data downloaded and saved to {file_path}")
    else:
        raise Exception(f"Failed to download Firebase data: {response.status_code} - {response.text}")
def upload_to_sharepoint(access_token, site_id, drive_id):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json',
    }
    
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:{folder_path}/{file_name}:/content"
    
    with open(file_path, 'rb') as file_data:
        response = requests.put(upload_url, headers=headers, data=file_data)
        
        if response.status_code in (200, 201):
            print("File uploaded successfully")
        else:
            raise Exception(f"Failed to upload file: {response.status_code} - {response.text}")
        
def show_loading_screen():
    loading_window = tk.Tk()
    loading_window.title("Processing")
    loading_window.geometry("300x100")
    
    window_width = 400
    window_height = 150
    screen_width = loading_window.winfo_screenwidth()
    screen_height = loading_window.winfo_screenheight()
    position_top = int(screen_height/2 - window_height/2)
    position_right = int(screen_width/2 - window_width/2)
    loading_window.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')
    
    ttk.Label(loading_window, text="The Firebase database is uploading to SharePoint").pack(pady=20)
    
    loading_window.update_idletasks()
    loading_window.update()

    return loading_window



def notify_user():
    os.system('osascript -e \'display notification "The file has been successfully uploaded." with title "Deployment Completed"\'')
    root = tk.Tk()
    root.withdraw()  
    root.destroy()


def main():
    try:
        access_token = get_access_token()
        download_firebase_data()
        site_id, drive_id = get_site_id_and_drive_id(access_token)
        upload_to_sharepoint(access_token, site_id, drive_id)
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    loading_window = show_loading_screen()
    main()
    loading_window.destroy()
    notify_user()
