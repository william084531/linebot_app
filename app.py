from flask import Flask, request, abort

import requests

import sys

import json  

import datetime

import gspread # 需在requirements.txt中加入此model如此heroku等雲端網站才會install model 

from oauth2client.service_account import ServiceAccountCredentials as SAC # 需在requirements.txt中加入此model如此heroku等雲端網站才會install model 

from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
from linebot.models import (MessageEvent, TextMessage, TextSendMessage,)


app = Flask(__name__)

# Channel Access Token
line_bot_api = LineBotApi('wSxXXjij0dj/0/z33YlxLWfgI/OsUk/b+X7xXDiD650DQ6+KChG4tIPOxiquUWhgyBenw+gTPPzTRMhIpwxe1t6fKPXG1tfew1HI54R0pNi2K8iRT/mDQmFjwfJCiLZGpIdX33IAKPm7YrbZYNCzAQdB04t89/1O/w1cDnyilFU=')#YOUR_CHANNEL_ACCESS_TOKEN
# Channel Secret
handler = WebhookHandler('d231cc7ff5d1e8a85df5af565461a1b8') # YOUR_CHANNEL_SECRET

#def excel_dowload():
#    target_url = 'http://192.168.43.246/test/index.txt'# 必須要固定ip不可用浮動ip
#    print('Start parsing excel ...')
#    rs = requests.session()
#    res = rs.get(target_url, verify=False)
#    res.encoding = 'utf-8'
#    data = str(res.content)
#    return data

class DateEncoder(json.JSONEncoder):  
    def default(self, obj):  
        if isinstance(obj, datetime.datetime):  
            return obj.strftime('%Y-%m-%d %H:%M:%S')  
        elif isinstance(obj, date):  
            return obj.strftime("%Y-%m-%d")  
        else:  
            return json.JSONEncoder.default(self, obj) 
# 監聽所有來自 /callback 的 Post Request
@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']
    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)
    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'
# 處理訊息
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    msg = event.message.text
#    print(msg)
    msg = msg.encode('utf-8')
#    if event.message.text == "test":
#      data = excel_dowload()
#      line_bot_api.reply_message(event.reply_token,TextSendMessage(text=data))
    if event.message.text == "try":
      a = "hi"
      line_bot_api.reply_message(event.reply_token,TextSendMessage(text=a))
    if event.message.text != "" and "$" in event.message.text:
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text='記錄完成'))
        pass
        #GDriveJSON就輸入下載下來Json檔名稱
        #GSpreadSheet是google試算表名稱
        GDriveJSON = 'linebot-253309-c09bb55fd137.json'
        GSpreadSheet = 'linebot'
        ss = 1
        while ss > 0:
            try:
                scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
                key = SAC.from_json_keyfile_name(GDriveJSON, scope)
                gc = gspread.authorize(key)
                worksheet = gc.open(GSpreadSheet).sheet1
            except Exception as ex:
                print('無法連線Google試算表', ex)
                sys.exit(1)
            textt=""
            textt+=event.message.text
            split = str(event.message.text).index('$')
            if textt!="" and "$" in event.message.text:
                worksheet.append_row((json.dumps(datetime.datetime.now(), cls=DateEncoder)[1:11], textt[:split],textt[split+1:]))
                print('新增一列資料到試算表' ,GSpreadSheet)
                return textt
    if "總結" in event.message.text:
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        GDriveJSON = 'linebot-253309-c09bb55fd137.json'
        GSpreadSheet = 'linebot'
        key = SAC.from_json_keyfile_name(GDriveJSON, scope)
        gc = gspread.authorize(key)
        worksheet = gc.open(GSpreadSheet).sheet1
        cal = worksheet.get_all_records()
        index = []
        value = []
        for i in range(len(cal)):
          if json.dumps(datetime.datetime.now(), cls=DateEncoder)[1:11] in cal[i]['時間']:
            index.append(i)
        for i in index:
          value.append(int(cal[i]['價格']))
        sums = sum(value)
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text='今天總共花費: '+str(sums)))
    if '#' in event.message.text and '$' not in event.message.text:
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        GDriveJSON = 'linebot-253309-c09bb55fd137.json'
        GSpreadSheet = 'linebot'
        key = SAC.from_json_keyfile_name(GDriveJSON, scope)
        gc = gspread.authorize(key)
        worksheet = gc.open(GSpreadSheet).sheet1
        cal = worksheet.get_all_records()
        index_time = []
        index_subject = []
        value = []
        for i in range(len(cal)):
          if json.dumps(datetime.datetime.now(), cls=DateEncoder)[1:11] in cal[i]['時間']:
            index_time.append(i)
        for i in index_time:
          if event.message.text[1::] in cal[i]['項目']:
            index_subject.append(i)
        for i in index_subject:
          value.append(int(cal[i]['價格']))
        sums = sum(value)
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text='今天在'+event.message.text[1::]+'上總共花費: '+str(sums)))
    if '說明' in event.message.text:
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text='#號後面加上欲查詢的項目內容可知道該項目今天花了多少錢，記帳方式為(項目名稱$價錢  ex : 交通費$123 or 住宿費$234)，總結會把所有今天內的項目價格加總，還在努力更新中請給我更多鼓勵，謝謝您'))
    if event.message.text == '你好':
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text='你好我是機器人excel_man 目前擁有記帳和計算當日花費的能力，請多多善用我喔!! ^_^ '))

import os
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
