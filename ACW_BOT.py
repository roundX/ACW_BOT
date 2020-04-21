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


def makeClearCopy(contents):
    contents = contents.replace("どうですか","どうですか？\n")
    contents = contents.replace("ごめん","ごめん、")
    contents = contents.replace("じゃん","じゃん？\n")
    contents = contents.replace("ですが","ですが、")
    contents = contents.replace("からね","からね。\n")
    contents = contents.replace("ました","ました。\n")
    contents = contents.replace("じゃあ","じゃあ、")
    contents = contents.replace("例えば","例えば、")
    contents = contents.replace("だね","だね。\n")
    contents = contents.replace("んで","んで、")
    contents = contents.replace("です","です。\n")
    contents = contents.replace("すね","すね。\n")
    contents = contents.replace("たい","たい。\n")
    contents = contents.replace("よね","よね。\n")
    contents = contents.replace("ます","ます。\n")
    contents = contents.replace("かな","かな？\n")
    contents = contents.replace("さい","さい。\n")
    contents = contents.replace("けど","けど、")
    contents = contents.replace("うね","うね。\n")
    contents = contents.replace("だよ","だよ。\n")
    contents = contents.replace("すか","すか？\n")
    contents = contents.replace("ので","ので、")
    contents = contents.replace("いね","いね。\n")
    contents = contents.replace("うか","うか。\n")
    contents = contents.replace("はい","はい。\n")
    contents = contents.replace("つって","つって、")
    contents = contents.replace("です。ね","ですね。\n")
    contents = contents.replace("だよ。ね。","だよね。\n")
    contents = contents.replace("いうと","いうと、")
    contents = contents.replace("ません","ません。")
    contents = contents.replace("ません","ません。\n")
    contents = contents.replace("です。かね","ですかね。\n")
    contents = contents.replace("あれか","あれか、")
    contents = contents.replace("いいや","いいや。\n")
    contents = contents.replace("です。か","ですか？\n")
    contents = contents.replace("ます。ね。","ますね。\n")
    contents = contents.replace("みたい。やね","みたいやね。\n")
    contents = contents.replace("ます。か","ますか？\n")
    contents = contents.replace("やります。んで、","やりますんで、")
    contents = contents.replace("たい。んで、","たいんで、")
    contents = contents.replace("んで、すが、","んですが、")

    return contents


def main():
    creds = getCredential()
    service = build('docs', 'v1', credentials=creds)

        sleep(10)

    while True:
        # Retrieve the documents contents from the Docs service.
        document = service.documents().get(documentId=DOCUMENT_ID).execute()

        # 本文の取得。
        contents = format(document.get('body').get('content')[1].get('paragraph').get('elements')[0].get('textRun').get('content'))
        if len(contents) == 2: #空の場合60秒待機
            sleep(60)
            continue
        # 清書
        fairCopy = makeClearCopy(contents)

        # 投稿
        upCommentToTeams(fairCopy)

        # 本文削除のレンジを決める(最初から最後まで)
        endChar = len(contents)
        requests = [{'deleteContentRange': {'range': {'startIndex': 1,'endIndex': endChar,}}},]
        # ドキュメント内本文削除
        service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': requests}).execute()

        # 本文に追加する文字設定
        requests = [{'insertText': {'location': {'index': 1,},'text': " "}}]
        # 本文に文字追加
        service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': requests}).execute()

        # 投稿の時間間隔指定
        sleep(60)


if __name__ == '__main__':
    main()
