from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage
from openpyxl import Workbook, load_workbook
import re
import os

app = Flask(__name__)

LINE_CHANNEL_ACCESS_TOKEN = 'WaARLthsMdnNK5UN4ZiWkaxh7PQZYPdnw2AlYuNmW/5u4KUQ+/b+4PhhPsm5VlpVnj9/ty7R/jiM6p4q0cd4gqO7GZl7Wiwpj+FLnAp22JsPGHZ38eO3wU86D4octYb9nBfXLf5P+7fq99IkPY/ifAdB04t89/1O/w1cDnyilFU=
LINE_CHANNEL_SECRET = '0d2611955011c95e5951ca22c5678995'

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

EXCEL_FILE = '模具紀錄.xlsx'

def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.append(["出廠日期", "入廠日期", "名稱", "原因"])
        wb.save(EXCEL_FILE)

def parse_message(text):
    data = {"出廠日期": "", "入廠日期": "", "名稱": "", "原因": ""}
    for key in data:
        match = re.search(f"{key}\\s*(.*)", text)
        if match:
            data[key] = match.group(1).strip()
    return data

def save_to_excel(data):
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append([data["出廠日期"], data["入廠日期"], data["名稱"], data["原因"]])
    wb.save(EXCEL_FILE)

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text
    if "模具出廠通知" in text or "模具入廠通知" in text:
        data = parse_message(text)
        save_to_excel(data)
        line_bot_api.reply_message(event.reply_token, TextMessage(text="✅ 已記錄"))

if __name__ == "__main__":
    init_excel()
    app.run(host="0.0.0.0", port=10000)
