# from __future__ import print_function
import pickle
from shutil import rmtree, make_archive
import os
import urllib
import sys
from io import BytesIO
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors
from googleapiclient.http import MediaIoBaseDownload

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/drive']
RESUME = False

def load_service():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('drive', 'v3', credentials=creds)
    return service

def export_file(service, file_id, file_name, folder_path, mimeType, ext):
    global RESUME
    file_name = file_name.replace("/",u"\u2215")
    file_name = file_name + ext
    file_name = os.path.join(folder_path, file_name)
    if RESUME == True:
        if os.path.exists(file_name):
            print(f"Skipping export: {file_name}")
            return
    request = service.files().export(fileId=file_id, mimeType=mimeType)
    fh = BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        # print("Download %d%%." % int(status.progress() * 100))
    f = open(file_name, 'wb')
    f.write(fh.getvalue())


def download_file(service, file_id, file_name, folder_path):
    global RESUME
    file_name = os.path.join(folder_path, file_name)
    if RESUME == True:
        if os.path.exists(file_name):
            print(f"Skipping download: {file_name}")
            return
    request = service.files().get_media(fileId=file_id)
    fh = BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        # print("Download %d%%." % int(status.progress() * 100))
    f = open(file_name, 'wb')
    f.write(fh.getvalue())
    # print("download complete")

def recurse_folder(service, folder_id, folder_path):
    global RESUME
    # """
    # Call the Drive v3 API
    # query params reference: https://developers.google.com/drive/api/v3/ref-search-terms#file_properties https://developers.google.com/drive/api/v2/ref-search-terms
    page_token = None
    while True:
        items = None
        # https://developers.google.com/drive/api/v3/reference/files/list
        try:
            results = service.files().list(
                                            q=f"parents in '{folder_id}'",
                                            pageSize=10, 
                                            fields="nextPageToken, files(id, name, mimeType)",
                                            pageToken=page_token
                                        ).execute()
            items = results.get('files', [])
        except:
            print(f"ERROR PULLING FOLDER: {folder_path}")
            continue
        if not items:
            print(f'No files found for folder: {folder_id}.')
        else:
            print('Files:')
            for item in items:
                print(item)
                if item["mimeType"] == "application/vnd.google-apps.folder":
                    new_folder = item["name"]
                    new_folder = new_folder.replace("/",u"\u2215")
                    new_path = folder_path + "/" + new_folder
                    if not os.path.exists(new_path):
                        os.makedirs(new_path)
                    recurse_folder(service, item["id"], new_path)
                elif "google-apps" in item["mimeType"]:
                    # print(f"MimeType: {item['mimeType']}")
                    if item['mimeType'] == "application/vnd.google-apps.document":
                        # mimeType = "application/rtf"
                        mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        ext = ".doc"
                    elif item['mimeType'] == "application/vnd.google-apps.spreadsheet":
                        # mimeType = "text/csv"
                        # ext = ".csv"
                        mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        ext = ".xls"
                    elif item['mimeType'] == "application/vnd.google-apps.drawing":
                        mimeType = "image/png"
                        ext = ".png"
                    elif item['mimeType'] == "application/vnd.google-apps.form":
                        break
                    elif item['mimeType'] == "application/vnd.google-apps.shortcut":
                        break
                    else:
                        mimeType = "application/pdf"
                        ext = ".pdf"
                    export_file(service, item["id"], item["name"], folder_path, mimeType, ext)
                else:
                    download_file(service, item["id"], item["name"], folder_path)
                
        page_token = results.get('nextPageToken', None)
        if page_token is None:
            break

def main():
    global RESUME
    try:
        folder_id = sys.argv[1]
    except IndexError:    
        print('\033[91m' + 'missing arguement folder_id!' + '\033[0m')
        exit()
    try:
        RESUME = bool(sys.argv[2])
    except IndexError:
        RESUME = False
    print(f"Resume Mode: {RESUME}")
    #delete anything existing and start fresh each time
    folder_name = "download_6"
    if RESUME == False:
        if os.path.exists(folder_name):
            rmtree(folder_name)
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
    root_folder = os.getcwd() + "/" + folder_name
    service = load_service()
    recurse_folder(service, folder_id, root_folder)
    print("download complete -> making zip file")
    make_archive("download", "zip", root_folder)
    if RESUME == False:
        if os.path.exists(folder_name):
            rmtree(folder_name)
    print(f"Your zipfile can be found at: {'download.zip'}")

if __name__ == '__main__':
    main()