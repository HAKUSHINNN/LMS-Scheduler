import requests
from bs4 import BeautifulSoup
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import datetime
import json

today = datetime.datetime.now()
print(today)

hp_url = "https://lms-tokyo.iput.ac.jp/"
login_url = f"{hp_url}login/index.php"
calendar_url = f"{hp_url}calendar/view.php?view=month"

# config.jsonからキーワードを読み込む
with open('config.json', 'r', encoding='utf-8') as file:
    config = json.load(file)

# セッションを開始してログインページにアクセス
session = requests.Session()
response = session.get(login_url)

# レスポンスからHTMLを解析してトークンを取得
soup = BeautifulSoup(response.content, 'html.parser')
logintoken = soup.find('input', {'name': 'logintoken'})['value']

username = config["username"]
password = config["password"]

# ログインに使用する情報
payload = {
    'username': username,  # ユーザ名
    'password': password,  # パスワード
    'logintoken': logintoken  # 取得したトークン
}

# ログインリクエストを送信
login_response = session.post(login_url, data=payload)

if login_response.url == hp_url:
    print("ログインに成功しました！")
    
    # カレンダーページにアクセス
    calendar_response = session.get(calendar_url)
    calendar_soup = BeautifulSoup(calendar_response.text, 'html.parser')
    
    assignments = []
    
    # 月日を取得
    calendar_div = calendar_soup.find("div", class_="calendarwrapper")
    year = calendar_div["data-year"]
    month = calendar_div["data-month"]
    
    weeks = calendar_soup.find_all("tr", {"data-region": "month-view-week"})

    keywords = config["keywords"]

    # 各週のデータから提出課題イベントを抽出
    for week in weeks:
        days = week.find_all('td', {"data-region": "day"})
        for day in days:
            date = day["data-day"]
            events = day.find_all('li', {"data-region": "event-item"})
            for event in events:
                title = event.find('span', class_="eventname").text.strip()
                if any(keyword in title for keyword in keywords):  # 提出課題，レポートのみをフィルタリング
                    title = title.replace("の提出期限が到来しています。", "")
                    link = event.find('a', {"data-action": "view-event"})["href"]
                    assignments.append({"date": f"{year}-{month}-{date}", "title": title, "link": link})

    # 課題の説明文とコース名を取得
    for assignment in assignments:
        event_response = session.get(assignment["link"])
        event_soup = BeautifulSoup(event_response.content, 'html.parser')
        description_div = event_soup.find('div', class_="no-overflow")
        if description_div:
            description = description_div.get_text(separator='\n').strip()  # 課題の説明文を抽出
            assignment["description"] = description + f"\n<br><a href='{assignment['link']}'>詳細はこちら</a>"

        else:
            assignment["description"] = f"説明文が取得出来ませんでした．下記URLより内容を確認して下さい．\n<br><a href='{assignment['link']}'>詳細はこちら</a>"
            
        course_name_tag = event_soup.find("h1").find("a")
        if course_name_tag:
            course_name = course_name_tag.text.strip()
            assignment["course"] = course_name
        else:
            assignment["course"] = "Unknown Course"
        
        # "完了"のメッセージがあるかチェック
        completion_info = event_soup.find('div', class_='completion-info')

        if completion_info and "完了としてマークする" in completion_info.text:
            assignment["completed"] = False

        else:
            assignment["completed"] = True


    # カレンダーAPIのスコープを設定
    SCOPES = ['https://www.googleapis.com/auth/calendar']

    # ユーザー認証を行いAPIクライアントを作成
    flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
    creds = flow.run_local_server(port=0)

    service = build('calendar', 'v3', credentials=creds)

    # 課題をGoogleカレンダーに追加または更新
    for assignment in assignments:
        event_start = f"{assignment['date']}T09:00:00+09:00"
        event_end = f"{assignment['date']}T10:00:00+09:00"
        event_summary = f"{assignment['course']} - {assignment['title']}"
        
        assignment_date = datetime.datetime.strptime(assignment["date"], "%Y-%m-%d")

        # 色を設定するロジック
        if assignment["completed"]:
            color_id = "2"  # ライトグリーン
        elif assignment_date < today and not assignment["completed"]:
            color_id = "11"
        else:
            color_id = "1"  # 青

        # 既存のイベントを検索
        events_result = service.events().list(
            calendarId='primary',
            timeMin=event_start,
            timeMax=event_end,
            q=event_summary,
            singleEvents=True
        ).execute()
        
        events = events_result.get('items', [])
        
        if events:
            # 既存のイベントがある場合は更新
            event_id = events[0]['id']
            event = {
                'summary': event_summary,
                'description': assignment.get('description', 'No description found'),
                'start': {
                    'dateTime': event_start,
                    'timeZone': 'Asia/Tokyo',
                },
                'end': {
                    'dateTime': event_end,
                    'timeZone': 'Asia/Tokyo',
                },
                'colorId': color_id
            }
            service.events().update(calendarId='primary', eventId=event_id, body=event).execute()
            print(f"Event updated: {event_summary}")
        else:
            # 新しいイベントを作成
            event = {
                'summary': event_summary,
                'description': assignment.get('description', 'No description found'),
                'start': {
                    'dateTime': event_start,
                    'timeZone': 'Asia/Tokyo',
                },
                'end': {
                    'dateTime': event_end,
                    'timeZone': 'Asia/Tokyo',
                },
                'colorId': color_id
            }
            service.events().insert(calendarId='primary', body=event).execute()
            print(f"Event created: {event_summary}")

    print("課題提出日をGoogleカレンダーに追加または更新しました。")
