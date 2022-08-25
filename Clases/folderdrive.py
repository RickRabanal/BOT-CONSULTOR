from __future__ import print_function
import pickle
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from apiclient.http import MediaFileUpload, MediaIoBaseDownload
import io
from apiclient import errors
from apiclient import http
import logging
from apiclient import discovery


# To list folders
def listfolders(service, filid, des):
    results = service.files().list(
        pageSize=1000, q="\'" + filid + "\'" + " in parents",
        fields="nextPageToken, files(id, name, mimeType)").execute()
    # logging.debug(folder)
    folder = results.get('files', [])
    for item in folder:
        if str(item['mimeType']) == str('application/vnd.google-apps.folder'):
            if not os.path.isdir(des+"/"+item['name']):
                os.mkdir(path=des+"/"+item['name'])
            print(item['name'])
            listfolders(service, item['id'], des+"/"+item['name'])  # LOOP un-till the files are found
        else:
            downloadfiles(service, item['id'], item['name'], des)
            print(item['name'])
    return folder


# To Download Files
def downloadfiles(service, dowid, name,dfilespath):
    request = service.files().get_media(fileId=dowid)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))
    with io.open(dfilespath + "/" + name, 'wb') as f:
        fh.seek(0)
        f.write(fh.read())


def Descarga_drive(Folder_id):
##    global creds,results,Folder_id,service,items
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/drive']

    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)  # credentials.json download from drive API
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('drive', 'v3', credentials=creds)

    Folder_id = '1m7jICaG30ym7-1B6YYzKNgFyLooXNsQv'  # Enter The Downloadable folder ID From Shared Link

    results = service.files().list().execute()

    items = results.get('files')[:3]  ##Obtengo los 3 primeros archivos    

    if not items:
        print('No files found.')
    else:
        print('Files:')
        for item in items:

            if item['mimeType'] == 'application/vnd.google-apps.folder' and item['id']=='1eVHJgJrVAu7KsrtPUroMn1VW09g8Bykt': #Filtro solo la carpeta y el link del folder

                #if not os.path.isdir("MisBases"):   #Si no existe la carpeta 'MisBases'
                #   os.mkdir("MisBases")            #Se crea una carpeta con ese nombre

                bfolderpath = os.getcwd()+"/MisBases/"
                listfolders(service, item['id'], bfolderpath) #Descarga solo los folders
                


##if __name__ == '__main__':
##    main()

##listfolders(service, item['id'], bfolderpath)
##Download 100%.
##Base Retenciones Negocios.xlsx
##[{'id': '1ZQYKP52ngjYugTOsZ7MnK0nB-9j6LGlg', 'name': 'Base Retenciones Negocios.xlsx', 'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}]

#link1: https://drive.google.com/drive/folders/1m7jICaG30ym7-1B6YYzKNgFyLooXNsQv?usp=sharing
#link2: https://drive.google.com/drive/folders/1eVHJgJrVAu7KsrtPUroMn1VW09g8Bykt?usp=sharing
