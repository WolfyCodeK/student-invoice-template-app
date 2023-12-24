from __future__ import print_function

import base64
from email.message import EmailMessage
from json import JSONDecodeError
from tkinter import messagebox
import os

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from cryptography.fernet import Fernet


class GmailAPI():
    # If modifying these scopes, delete the file token.json.
    SCOPES = [
    "https://mail.google.com/",
    "https://www.googleapis.com/auth/gmail.metadata"
    ]

    PARENT_FOLDER = "lib/"
        
    def __init__(self):
        # Create resources directory if it does not exist
        if not os.path.exists(self.PARENT_FOLDER):
            os.makedirs(self.PARENT_FOLDER)
        
    def missingCredentials():
        messagebox.showerror(title="Missing files", message='ERROR: Missing credentials, cannot access google API.')  
        
    def decryptCredentials(self):
        if (os.path.exists(f'{self.PARENT_FOLDER}key.key')):
            #this just opens your 'key.key' and assings the key stored there as 'key'
            with open(f'{self.PARENT_FOLDER}key.key','rb') as file:
                key = file.read()
                
            #this opens your json and reads its data into a new variable called 'data'
            with open(f'{self.PARENT_FOLDER}credentials.json','rb') as f:
                data = f.read()

            #this encrypts the data read from your json and stores it in 'encrypted'
            fernet = Fernet(key)

            #this writes your new, encrypted data into a new JSON file
            with open(f'{self.PARENT_FOLDER}credentials.json','wb') as f:
                f.write(fernet.decrypt(data))     
        else:
            # very bad code
            self.encryptCredentials()
            self.decryptCredentials()
            
    def encryptCredentials(self):
        #this generates a key and opens a file 'key.key' and writes the key there
        key = Fernet.generate_key()
        with open(f'{self.PARENT_FOLDER}key.key','wb') as file:
            file.write(key)

        #this just opens your 'key.key' and assings the key stored there as 'key'
        with open(f'{self.PARENT_FOLDER}key.key','rb') as file:
            key = file.read()

        #this opens your json and reads its data into a new variable called 'data'
        with open(f'{self.PARENT_FOLDER}credentials.json','rb') as f:
            data = f.read()

        #this encrypts the data read from your json and stores it in 'encrypted'
        fernet = Fernet(key)
        encrypted = fernet.encrypt(data)

        #this writes your new, encrypted data into a new JSON file
        with open(f'{self.PARENT_FOLDER}credentials.json','wb') as f:
            f.write(encrypted)
            
        return key

    def gmail_create_draft(self, subject: str, body: str):
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is created automatically when the authorization flow completes for the first time.
        if os.path.exists(f'{self.PARENT_FOLDER}token.json'):
            creds = Credentials.from_authorized_user_file(f'{self.PARENT_FOLDER}token.json', self.SCOPES)
            
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                
                self.decryptCredentials()
                
                try:
                    flow = InstalledAppFlow.from_client_secrets_file(f'{self.PARENT_FOLDER}credentials.json', self.SCOPES)
                except JSONDecodeError:
                    self.missingCredentials()
                    exit
                creds = flow.run_local_server(port=0)
                
                self.encryptCredentials()
                
            # Save the credentials for the next run
            with open(f'{self.PARENT_FOLDER}token.json', 'w') as token:
                token.write(creds.to_json())

        try:
            # create gmail api client
            service = build('gmail', 'v1', credentials=creds)

            message = EmailMessage()

            message.set_content(body)

            message['Subject'] = subject

            # encoded message
            encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

            create_message = {
                'message': {
                    'raw': encoded_message
                }
            }
            
            draft = service.users().drafts().create(
                userId="me",
                body=create_message
            ).execute()

        except HttpError as error:
            print(F'An error occurred: {error}')
            draft = None

        return draft

