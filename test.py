import json
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = [
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file',
]


# The ID of a sample document.
# DOCUMENT_ID = '195j9eDD3ccgjQRttHhJPymLJUCOUjs-jmwTrekvdjFE'


def main():
    """Shows basic usage of the Docs API.
    Prints the title of a sample document.
    """
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
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    body = {
        'title': 'please',
        'body': {
            "content": [
                {
                    "startIndex": 1,
                    "endIndex": 6,
                    "paragraph": {
                        "elements": [
                            {
                                "startIndex": 1,
                                "endIndex": 6,
                                "textRun": {
                                    "content": "Test\n",
                                    "textStyle": {
                                        "fontSize": {
                                            "magnitude": 30,
                                            "unit": "PT"
                                        }
                                    }
                                }
                            }
                        ],
                        "paragraphStyle": {
                            "namedStyleType": "NORMAL_TEXT",
                            "direction": "LEFT_TO_RIGHT"
                        }
                    }
                },
                {
                    "startIndex": 6,
                    "endIndex": 10,
                    "paragraph": {
                        "elements": [
                            {
                                "startIndex": 6,
                                "endIndex": 10,
                                "textRun": {
                                    "content": "\ud14c\uc2a4\ud2b8\n",
                                    "textStyle": {
                                        "fontSize": {
                                            "magnitude": 30,
                                            "unit": "PT"
                                        }
                                    }
                                }
                            }
                        ],
                        "paragraphStyle": {
                            "namedStyleType": "NORMAL_TEXT",
                            "direction": "LEFT_TO_RIGHT"
                        }
                    }
                },
            ]
        }
    }
    service = build('docs', 'v1', credentials=creds)

    # Retrieve the documents contents from the Docs service.
    # document = service.documents().get(documentId='13m645CsJYdDI18sHcTTEQtf0n_nnEx4xvADLmtIJacI').execute()
    document = service.documents().create(body=body).execute()
    text_list = ['test1sadasddsa\n', 'test2adasdasdas\n', 'teasdsadasdst3\n', ]

    requests = [
        {
            'insertText': {
                'location': {
                    'index': 1,
                },
                'text': text_list[0]
            }
        },
        {
            'insertText': {
                'location': {
                    'index': 1+len(text_list[0]),
                },
                'text': text_list[1]
            }
        },
        {
            'insertText': {
                'location': {
                    'index': 1 + len(text_list[0]) + len(text_list[1]),
                },
                'text': text_list[2]
            }
        },
    ]
    result = service.documents().batchUpdate(
        documentId=document.get('documentId'), body={'requests': requests}).execute()
    # print('The title of the document is: {}'.format(document.get('title')))
    document = service.documents().get(documentId=result.get('documentId')).execute()

    file_path = 'create.json'
    try:
        fh = open(file_path, 'wb')
        fh.close()
        os.remove(file_path)
    except FileNotFoundError:
        pass
    finally:
        with open(file_path, 'w') as f:
            json.dump(document, f)


if __name__ == '__main__':
    main()
