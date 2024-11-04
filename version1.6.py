"""
Author: Berwyn
Date: 2024/11/1
Description: [Description]
Version: [Description]
"""

import requests
import json
from apscheduler.schedulers.blocking import BlockingScheduler

# 设置全局变量 - 这些变量的值需要修改为实际的应用参数
app_id = "cli_a6f036804cb1d00e"
app_secret = "KOzmEngu4xLabWFWOPsHKdLkuB8U0geJ"
spreadsheet_token = 'J4gEsgbSEhixbGt4iJtcfViVnac'
sheet_id = "8b4b0c"
range = f'{sheet_id}!DO14:DO16'
url = f'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{range}'
link = "https://li.feishu.cn/sheets/J4gEsgbSEhixbGt4iJtcfViVnac"
receive_id = "oc_4faec7f80eb3d18eda97935c261b8262"


# 获取 access token
def get_access_token(app_id, app_secret):
    try:
        url_request = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
        headers = {
            "Content-Type": "application/json; charset=utf-8"
        }
        payload = {
            "app_id": app_id,
            "app_secret": app_secret
        }
        response = requests.post(url_request, headers=headers, json=payload)
        response_data = response.json()
        if response.status_code == 200 and "tenant_access_token" in response_data:
            return response_data["tenant_access_token"]
        else:
            print(f"Error obtaining token: {response_data}")
            return None
    except requests.RequestException as e:
        print(f"Network error: {e}")
        return None


# 读取Google Sheet值
def get_message(access_token):
    try:
        headers_message = {
            'Authorization': f"Bearer {access_token}",
            'Content-Type': 'application/json; charset=utf-8'
        }

        params = {
            'valueRenderOption': 'FormattedValue'
        }

        response = requests.get(url, headers=headers_message, params=params)

        if response.status_code == 200:
            data = response.json()
            if data.get('code') == 0:
                values = data['data']['valueRange']['values']
                return [v[0] for v in values]
            else:
                print(f"API Error: {data['msg']}")
        else:
            print(f"HTTP Error: {response.status_code} - {response.text}")
    except requests.RequestException as e:
        print(f"Network error: {e}")
    return None


# 发送消息到群聊
def send_message(title, plan, actual, link, access_token):
    try:
        url = "https://open.feishu.cn/open-apis/im/v1/messages"
        params = {"receive_id_type": "chat_id"}
        msg = "{}\n{}\n{}\n{}".format(title, plan, actual, link)
        msgContent = {
            "text": msg,
        }
        req = {
            "receive_id": receive_id,
            "msg_type": "text",
            "content": json.dumps(msgContent)
        }
        payload = json.dumps(req)
        headers = {
            'Authorization': f"Bearer {access_token}",
            'Content-Type': 'application/json'
        }
        response = requests.post(url, params=params, headers=headers, data=payload)
        if response.status_code == 200:
            print("Message sent successfully.")
        else:
            print(f"Message send error: {response.status_code} - {response.text}")
        print(response.headers.get('X-Tt-Logid'))
    except requests.RequestException as e:
        print(f"Network error: {e}")


# 定时任务
def job():
    access_token = get_access_token(app_id, app_secret)
    if not access_token:
        print("Failed to obtain access token. Exiting...")
        return

    messages = get_message(access_token)
    if messages and len(messages) == 3:
        title, plan, actual = messages
        send_message(title, plan, actual, link, access_token)
    else:
        print("Failed to retrieve messages or message format is incorrect.")


# 使用 APScheduler 定时每天执行一次任务
scheduler = BlockingScheduler()
scheduler.add_job(job, 'cron', hour=11, minute=23)  # 每天早上9点执行
scheduler.start()

# 进入终端输入
# cd /Users/zhangbowen7/MyCode/Daily_Report/forbabe
# caffeinate -i /Users/zhangbowen7/MyCode/Daily_Report/venv/bin/python3 version1.6.py