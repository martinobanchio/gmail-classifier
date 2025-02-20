from utils import *

SCOPES = ['https://www.googleapis.com/auth/gmail.modify', 
          'https://www.googleapis.com/auth/generative-language.retriever',
          'https://www.googleapis.com/auth/spreadsheets']


creds = None
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

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
