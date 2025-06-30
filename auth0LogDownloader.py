import pandas as pd
import requests
import os
import sys
import configparser
from dotenv import load_dotenv
import ctypes  # ポップアップ用

# メッセージボックス関数
def show_message(title, message):
    ctypes.windll.user32.MessageBoxW(0, message, title, 0)

# .envファイルの読み込み
if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(sys.executable)
else:
    base_path = os.path.dirname(__file__)

env_path = os.path.join(base_path, '.env')
load_dotenv(dotenv_path=env_path)

# 環境変数から認証情報を取得
domain = os.getenv("DOMAIN")
client_id = os.getenv("CLIENT_ID")
client_secret_key = os.getenv("CLIENT_SECRET")
audience = f"https://{domain}/api/v2/"

# Excel生成開始メッセージ
show_message("処理開始", "Excelファイルの生成を開始します。")

# Token取得
url = f'https://{domain}/oauth/token'
headers = {'Content-Type': 'application/json'}
body = {
    "grant_type": "client_credentials",
    "client_id": client_id,
    "client_secret": client_secret_key,
    "audience": audience
}

with requests.Session() as session:
    token_response = session.post(url, json=body, headers=headers)
    token_response.raise_for_status()
    token_data = token_response.json()
    api_token = token_data['access_token']

    headers = {'Authorization': f'Bearer {api_token}'}

    config = configparser.ConfigParser()
    config.read('last_log_id.ini')

    if 'DEFAULT' in config and 'last_log_id' in config['DEFAULT']:
        log_id = config['DEFAULT']['last_log_id']
    else:
        log_id = input("最後に取得したLogIDを入力してください：")

    params = {'from': log_id, 'take': 100}
    logs_list = []

    while True:
        logs_url = f'https://{domain}/api/v2/logs'
        response = session.get(logs_url, headers=headers, params=params)
        response.raise_for_status()
        logs = response.json()

        if not logs:
            break

        logs_list.extend(logs)

        if len(logs) < 100:
            break

        params['from'] = logs[-1]['log_id']

    log_columns = ["date", "type", "description", "connection_id", "client_id", "client_name",
                   "ip", "user_agent", "details", "hostname", "user_id", "user_name", "auth0_client",
                   "log_id", "_id", "isMobile", "audience", "scope", "connection", "strategy", "strategy_type",
                   "session_connection", "organization_id", "organization_name"]
    df_all = pd.DataFrame(logs_list, columns=log_columns)

    last_log_id = df_all.iloc[-1]['log_id'] if df_all.shape[0] > 0 else "ぴったりでした"
    config['DEFAULT']['last_log_id'] = last_log_id
    with open('last_log_id.ini', 'w') as configfile:
        config.write(configfile)

    def extract_date_from_log_id(log_id):
        return log_id[3:11]

    output_dir = os.path.join(base_path, "output")
    os.makedirs(output_dir, exist_ok=True)

    if df_all.shape[0] > 0:
        first_date = extract_date_from_log_id(df_all.iloc[0]['log_id'])
        last_date = extract_date_from_log_id(df_all.iloc[-1]['log_id'])
        filename = f"Logs_{first_date}-{last_date}.xlsx"
    else:
        filename = "Logs_empty.xlsx"

    f_path = os.path.join(output_dir, filename)
    df_all.to_excel(f_path, index=False)

    # 完了メッセージ
    show_message("完了", f"Excelファイルを保存しました:\n{f_path}")
