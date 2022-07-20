from flask import Flask, request, abort
import pandas as pd
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials

from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage, FollowEvent,
)
import os
from dotenv import load_dotenv
config = load_dotenv(".env")

worksheets = {}

def auth():
    worksheets = {}
    profile = line_bot_api.get_profile("userID")
    worksheet = worksheets[profile("userID")]
    
    SP_CREDENTIAL_FILE = 'cocolog-7fdb0e085d80.json'
    SP_SCOPE = [
        'https://spreadsheets.google.com/feeds',
        'https//www.googleapis.com/auth/drive'
    ]

    SP_SHEET_KEY = '1GhrXpa8FsS62woXBykX7uNckTOtalEq2PLjnR5mdb5c'
    SP_SHEET = 'worksheet'

    credentials = ServiceAccountCredentials.from_json_keyfile_name(SP_CREDENTIAL_FILE, SP_SCOPE)
    gc = gspread.authorize(credentials)

    worksheet = gc.open_by_key(SP_SHEET_KEY).worksheet(SP_SHEET)
    return worksheet

# 出勤
def punch_in():
    worksheet = auth()
    df = pd.DataFrame(worksheet.get_all_records())

    timestamp = datetime.now() + timedelta(hours=9)
    date = timestamp.strftime('%Y/%m/%d')
    punch_in = timestamp.strftime('%H:%M')

    df = df.append({'日付': date, '出勤時間': punch_in, '退勤時間': '00:00'}, ignore_index=True)
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())

#　退勤
def punch_out():
    worksheet = auth()
    df = pd.DataFrame(worksheet.get_all_records())

    timestamp = datetime.now() + timedelta(hours=9)
    punch_out = timestamp.strftime('%H:%M')

    df.iloc[-1, 2] = punch_out
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())


class TalkWithBotUsers:
    def __init__(self, title):
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('cocolog-7fdb0e085d80.json', scope)
        global gc
        gc = gspread.authorize(credentials)
        gc = gc.open_by_key("1GhrXpa8FsS62woXBykX7uNckTOtalEq2PLjnR5mdb5c")

        try :
            #新たにワークシートを作成し、Worksheetオブジェクトをworksheetに格納します。
            worksheet = gc.add_worksheet(title=title)
        except :
            #すでにワークシートが存在しているときは、そのワークシートのWorksheetオブジェクトを格納します。
            worksheet = gc.worksheet(title)

        self.worksheet = worksheet #worksheetをメンバに格納
        self.worksheet.update_cell('A1', '日付')
        self.worksheet.update_cell('B1', '出勤時間')
        self.worksheet.update_cell('C1', '退勤時間')

app = Flask(__name__)

YOUR_CHANNEL_ACCESS_TOKEN = os.environ['YOUR_CHANNEL_ACCESS_TOKEN']
YOUR_CHANNEL_SERCRET = os.environ['YOUR_CHANNEL_SERCRET']

line_bot_api = LineBotApi('YOUR_CHANNEL_ACCESS_TOKEN')
handler = WebhookHandler('YOUR_CHANNEL_SERCRET')

@app.route("/callback", methods = ['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']
 
    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body:" + body)

    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        print("Invalid signature. Please check your channel access token/channel secret.")
        abort(400)

    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    if event.message.text == '出勤':
        punch_in()
        line_bot_api.reply_message(
            event.replyToken,
            TextSendMessage(text='おはようございます！本日もよろしくお願いします！'))
    elif event.message.text == '退勤':
        punch_out()
        line_bot_api.reply_message(
            event.replyToken,
            TextSendMessage(text='お疲れさまでした！'))
    else:
        line_bot_api.reply_message(
            event.replyToken,
            TextSendMessage(text='出退勤専用BOTです.「出勤」「退勤」でタイムカードを更新できます！'))

@handler.add(FollowEvent)
def handle_follow():
    profile = line_bot_api.get_profile("userId")
    worksheets[profile("userID")] = TalkWithBotUsers(profile.display_name)
    
if __name__ == "__main__ ":
    port = os.getenv("PORT")
    app.run(host="0.0.0.0", port=port)