from utils import *

SCOPES = ['https://www.googleapis.com/auth/gmail.modify', 
          'https://www.googleapis.com/auth/generative-language.retriever',
          'https://www.googleapis.com/auth/spreadsheets']


creds = refresh_or_generate_credentials()

genai.configure(credentials=creds)
gmail_service = build('gmail', 'v1', credentials=creds)
sheets_service = build("sheets", "v4", credentials=creds)

messages = get_unread_emails(gmail_service, mark_read=True)

json_email_contents = []

if not messages: 
    exit()
for message in messages:
    model_output = email_processing(message)
    json_email_contents.append(json_output(model_output))

update_values(sheets_service, '1ZJ-bSdn_E2-tzfwds9iJtXIYFKhGPnJ_Oi7iTOa-SdY', 'Sheet1', json_email_contents)
