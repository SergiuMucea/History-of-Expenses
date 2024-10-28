"""This module makes the connection with gmail and downloads the statements based on an assigned gmail label.
When an e-mail enters the mailbox, the Third Bank e-mails will be assigned label 'Extras Third Bank debit' and will be archived.
This module processes the emails with the assigned label, downloads the attachments and lastly removes the label
"""
import sys
import os.path
import base64
from typing import List
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Access scopes for gmail. If scopes modified within the gmail API, manually delete the file token.json.
SCOPES = ["https://mail.google.com/"]

def google_login():
    # defining gmail_error as global for error handling in the tkinter SecondaryGUI window
    global gmail_error
    gmail_error = False
            
    creds = None

    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    token_path = "path/to/the/token.json"

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
        
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(
            "path/to/the/credentials.json", 
            SCOPES,
        )
        creds = flow.run_local_server(port=0)
            
        # Save the credentials for the next run    
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
            
    # Initializing the gmail service
    service = build('gmail', 'v1', credentials=creds)
    return service

# Gets the messages based on the selected criteria: gmail label
def search_emails(query_filter):
    message_list = service.users().messages().list(
        userId='me', 
        q=query_filter,
    ).execute()
    
    message_items = message_list.get('messages')
    return message_items   

# Gets the information needed for getting to the attachments 
def get_message_detail(message_id, msg_format='metadata', metadata_headers: List=None):
	message_detail = service.users().messages().get(
		userId='me',
		id=message_id,
		format=msg_format,
		metadataHeaders=metadata_headers,
	).execute()
 
	return message_detail

# Gets the information within the attachment
def get_attachment_data(message_id, attachment_id):
    response = service.users().messages().attachments().get(
		userId='me',
		messageId=message_id,
		id=attachment_id,
	).execute()

    file_data = base64.urlsafe_b64decode(response.get('data').encode('UTF-8'))
    return file_data

# Creates the name of the downloaded attachment based on the subject of the email
# Format: "year - Extras perioada day/month day/month"
def get_statement_name(header_info):
    headers = header_info
    
    for item in headers:
        if item['name'] == 'Subject':
            subject = item['value'][-30:]
            year = item['value'][-4:]
            attachment_name = f"{year} - Extras {subject.replace('/', '  ').replace('-' + year, '')}.pdf"
            return attachment_name
                 
# Removing the 'Extras Third Bank debit' label to avoid duplicate processing
def remove_label(msg_id):
    label_extras_thrd_debit = ['Label_id']
    label_action = {'removeLabelIds': label_extras_thrd_debit}
    service.users().messages().modify(
        userId='me',
        id=msg_id,
        body=label_action
    ).execute()

# To avoid the very rare case (maybe inexistent) of the program crashing if there are duplicates, checking for duplicates first
def check_duplicate_statement(location, attachment_name) -> bool:
    archive_location = location + "/Archive"
    attachment_path = os.path.join(archive_location, attachment_name)
    if os.path.exists(attachment_path):
        duplicate_statement = True
        return duplicate_statement


# The main function which will be imported in the main ThirdBank module
def get_attachments():
    try:
        # Defining gmail_error, no_email and duplicate_statement as global for error handling in the tkinter SecondaryGUI window
        global gmail_error
        gmail_error = False
        global duplicate_statement
        duplicate_statement = False
        global no_email
        
        # Initializing the gmail service
        global service
        service = google_login()
        
        try:   
            no_email = False 
            
            save_location = "path/to/the/Project/InputFiles/Statements"
                                  
            email_count = 0
            attachment_count = 0
            
            # Fetching the e-mails and iterating over them
            message_items = search_emails(query_filter='label:extras-third-bank-debit')
            for msg in message_items:
                message_detail = get_message_detail(msg['id'], msg_format='full', metadata_headers=['parts'])
                message_detail_payload = message_detail.get('payload')
                email_count += 1
                
                attachment_name = get_statement_name(header_info=message_detail_payload['headers'])
                
                # Checking for duplicate statements and in the very rare case this happens, the program stops
                duplicate_statement = check_duplicate_statement(save_location, attachment_name)
                if duplicate_statement:
                    duplicate_statement = ("One or more statements have already been processed once."
                    "\nCheck the statements and try again.")
                    print(duplicate_statement)
                    sys.exit()
                
                # Itering over the e-mail parts
                if 'parts' in message_detail_payload:
                    for msg_payload in message_detail_payload['parts']:
                        body = msg_payload['body']

                        # Checking and dowloading attachment
                        if 'attachmentId' in body:
                            attachment_id = body['attachmentId']
                            attachment_content = get_attachment_data(msg['id'], attachment_id)
                                                
                            with open(os.path.join(save_location, attachment_name), 'wb') as _f:
                                _f.write(attachment_content)
                                attachment_count += 1
                                status = f"File '{attachment_name}' is saved at: {save_location}"
                                print(status)
                                
                            remove_label(msg['id'])                    
                
                time.sleep(0.5)
                
            final_status = f"Downloaded {attachment_count} attachments out of {email_count} emails"
            print(final_status)
            
            
        except TypeError:
            no_email = "There are no statements to be processed.\nThanks for coming by!"
            print(no_email)
            sys.exit()
            
    except Exception as error:
        gmail_error = f"An error occurred while logging in.\nCheck your network or gmail connection and try again!\n\n{error}"
        print(gmail_error)
        sys.exit()
    
