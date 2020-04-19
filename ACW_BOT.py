# Googleドキュメントの自動文字起こしをMicrosoft Teamsのコメントにリアルタイム投稿するBOT
from __future__ import print_function
import pymsteams
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from time import sleep


# If modifying these scopes, delete the file token.pickle.
# SCOPES = ['https://www.googleapis.com/auth/documents.readonly']
SCOPES = ['https://www.googleapis.com/auth/documents']

# 投稿する文章のあるドキュメントのID
DOCUMENT_ID = 'GoogleドキュメントのID'

def getCredential():
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
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return creds


def upCommentToTeams(content):
    # 各チームのwebhook URLを設定
    myTeamsMessage = pymsteams.connectorcard("Microsoft TeamsのwebhookのURL")
    myTeamsMessage.text(content)
    myTeamsMessage.send()


def main():
    creds = getCredential()
    service = build('docs', 'v1', credentials=creds)

    while True:
        # Retrieve the documents contents from the Docs service.
        document = service.documents().get(documentId=DOCUMENT_ID).execute()

        # 本文の取得。表示。
        contents = format(document.get('body').get('content')[1].get('paragraph').get('elements')[0].get('textRun').get('content'))
        print("投稿しました。")
        print('The content of the document is:', contents)

        # 投稿
        upCommentToTeams(contents)

        # 本文削除のレンジを決める(最初から最後まで)
        endChar = len(contents)
        requests = [{'deleteContentRange': {'range': {'startIndex': 1,'endIndex': endChar,}}},]

        # ドキュメント内本文削除
        service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': requests}).execute()

        # 投稿の時間間隔指定
        sleep(15)


if __name__ == '__main__':
    main()
