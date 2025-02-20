from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import google.generativeai as genai
import json
import pickle
import os.path
import base64
from datetime import datetime, timedelta

### Credential Management

def get_credentials_from_secrets():
    """Retrieves Google credentials from GitHub Secrets."""
    encoded_creds = os.environ.get("GOOGLE_CREDENTIALS")
    if not encoded_creds:
        raise ValueError("GOOGLE_CREDENTIALS secret not found.")

    try:
        decoded_creds = base64.b64decode(encoded_creds)
        creds = pickle.loads(decoded_creds)
        return creds
    except Exception as e:
        print(f"Error decoding or unpickling credentials: {e}")
        return None

def refresh_or_generate_credentials():
    """Refreshes or generates Google credentials."""
    creds = get_credentials_from_secrets()

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                return creds
            except Exception as refresh_error:
                print(f"Error refreshing credentials: {refresh_error}")
                return None
        else:
            try:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
                return creds
            except Exception as flow_error:
                print(f"Error generating new credentials: {flow_error}")
                return None
    return creds

### Gemini API Utils

def email_processing(input):
  model = genai.GenerativeModel("gemini-2.0-flash")
  prompt = """Extract from the attached email some information. I want to know:
  - whether the email is an announcement about a new show, a fair, or neither
  - what gallery sent the email
  - where the show or fair will be. if no show or fair, just leave blank
  - if it is a show or a fair, the time
  - if it is a show or a fair, the title of the exhibition.

  Use this JSON schema:

  Show = {'sender_email': str, 'is_show_or_fair': str; 'city': str, 'gallery': str, 'show_title': str, 'opening_time': str}

  Please return only the JSON with the proper formatting and nothing else. It has to be ready to be fed to the json.loads command in python.
  NEVER return an empty JSON. The JSON should always have the fields above, even if they are empty. The first and second fields should never be empty.
  
  The email is the following: """

  result = model.generate_content(prompt + input)
  return result


def json_output(input):
  return json.loads(input.text.replace("\n", "").replace("```", "").replace("json", ""))


### Gmail API Utils

def mark_as_read(service, message_id):
    """Marks a message as read."""
    try:
        service.users().messages().modify(
            userId='me', id=message_id, body={'removeLabelIds': ['UNREAD']}
        ).execute()
        print(f'Message {message_id} marked as read.')

    except Exception as e:
        print(f'An error occurred: {e}')

def last_week_query():
    """Retrieves emails received in the last week."""
    today = datetime.now()
    start_of_last_week = today - timedelta(days=6)  # Calculate the start of last week
    start_date = start_of_last_week.strftime('%Y/%m/%d')
    query = f'after:{start_date}'  # Use the calculated dates in the query
    return query


def get_unread_emails(service, mark_read=False):
    """Retrieves and prints unread email subjects."""
    try:
        results = service.users().messages().list(userId='me', q='is:unread').execute()
        messages = results.get('messages', [])

        if not messages:
            print('No unread messages found.')
            return

        output = []

        for message in messages:
            
            if mark_read:
              mark_as_read(service, message['id'])
            
            message_text = ''

            msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()
            headers = msg['payload']['headers']
            for header in headers:
                if header['name'] == 'From':
                    message_text += f'From: {header["value"]} \n'
                elif header['name'] == 'Subject':
                    subject = header['value']
                    message_text += f'Subject: {subject} \n'

            #Getting the body of the email.
            parts = msg['payload'].get('parts')
            if parts:
                for part in parts:
                    if part['mimeType'] == 'text/plain':
                        data = part['body']['data']
                        byte_code = base64.urlsafe_b64decode(data).decode("utf-8")
                        message_text += f'Body: {byte_code} \n'
                    elif part['mimeType'] == 'text/html':
                        data = part['body']['data']
                        byte_code = base64.urlsafe_b64decode(data).decode("utf-8")
                        message_text += f'HTML Body: {byte_code} \n'
            output.append(message_text)
        return output
    
    except Exception as e:
        print(f'An error occurred: {e}')

### Sheets API Utils

def json_to_list_of_strings(json_object):
    """
    Converts a JSON object to a list of strings.

    Args:
        json_object: A JSON object (dict or list).

    Returns:
        A list of strings representing the JSON object.
    """
    if isinstance(json_object, dict):
        return [f"{value}" for key, value in json_object.items()]
    elif isinstance(json_object, list):
        return [str(item) for item in json_object]
    else:
        raise TypeError("Input must be a JSON object (dict or list)")

def update_values(service, spreadsheet_id, range_name, _values, value_input_option='USER_ENTERED'):
  try:

    values = [json_to_list_of_strings(value) for value in _values]
    
    body = {"values": values}
    result = (
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption=value_input_option,
            body=body,
        )
        .execute()
    )
    print(f"{(result.get('updates').get('updatedCells'))} cells appended.")
    return result
  except Exception as e:
    print(f"An error occurred: {e}")
    return e