from functools import update_wrapper
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.http import MediaFileUpload

import os


class MyDrive():
    def __init__(self):
        # If modifying these scopes, delete the file token.json.
        SCOPES = ['https://www.googleapis.com/auth/drive']

        """Shows basic usage of the Drive v3 API.
        Prints the names and ids of the first 10 files the user has access to.
        """
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    './credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('./token.json', 'w') as token:
                token.write(creds.to_json())

        self.service = build('drive', 'v3', credentials=creds)

    def list_files(self, page_size=10):
        # Call the Drive v3 API
        results = self.service.files().list(
            pageSize=10, fields="nextPageToken, files(id, name)").execute()
        items = results.get('files', [])

        if not items:
            print('No files found.')
        else:
            print('Files:')
            for item in items:
                print(u'{0} ({1})'.format(item['name'], item['id']))

    def upload_file(self, filename, path):
        folder_id = "17pyMVaJVtQ9I0WeX8InsSFurlZAC8tgR"
        media = MediaFileUpload(f"{path}{filename}")

        response = self.service.files().list(
            q=f"name= '{filename}' and parents='{folder_id}'",
            spaces='drive',
            fields='nextPageToken, files(id, name)',
            pageToken=None).execute()
        
        if len(response['files']) == 0:
            file_metadata = {
                'name': filename,
                'parents': [folder_id]
            }
            file = self.service.files().create(
                body=file_metadata, media_body=media, fields='id').execute()
            print(f"A new file was created {file.get('id')}")

        else:
            for file in response.get('files', []):
                # process change

                update_file = self.service.files().update(
                    fileId=file.get('id'),
                    media_body=media,
                ).execute()
                print(f"File Updated")


def main():
    # my_drive.list_files()
    path = "./Files/"
    files = os.listdir(path)
    my_drive = MyDrive()

    for item in files:
        my_drive.upload_file(item, path)


if __name__ == '__main__':
    main()
