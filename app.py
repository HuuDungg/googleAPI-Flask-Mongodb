import random

from flask import Flask, redirect, request, session, url_for, render_template
from flask_session import Session
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from pymongo import MongoClient
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import random
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import datetime as dt

app = Flask(__name__)
app.secret_key = 'huudungisthebest'
app.config['SESSION_TYPE'] = 'filesystem'
Session(app)

# URL của bạn
LOCAL_URL = 'https://127.0.0.1:5000'

# Cấu hình OAuth 2.0
CLIENT_SECRETS_FILE = 'client_secret.json'
SCOPES = [
    'openid',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile',
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]

# Khởi tạo kết nối MongoDB
client = MongoClient(
    "mongodb+srv://huudung038:1@clusterhuudung.z5tdrft.mongodb.net/?retryWrites=true&w=majority&appName=ClusterHuuDung")
app.db = client.firstflaskapp
collection = app.db.hubData

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/login')
def login():
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE,
        scopes=SCOPES,
        redirect_uri=f'{LOCAL_URL}/oauth2callback'  # Sử dụng URL của bạn
    )
    authorization_url, state = flow.authorization_url(access_type='offline', include_granted_scopes='true')
    session['state'] = state
    return redirect(authorization_url)

@app.route('/oauth2callback')
def oauth2callback():
    state = session['state']
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE,
        scopes=SCOPES,
        state=state,
        redirect_uri=f'{LOCAL_URL}/oauth2callback'  # Sử dụng URL của bạn
    )
    try:
        flow.fetch_token(authorization_response=request.url)
    except Exception as e:
        return f'An error occurred during token fetch: {e}'

    credentials = flow.credentials
    session['credentials'] = credentials_to_dict(credentials)

    return redirect(url_for('select_sheet'))


@app.route('/select_sheet')
def select_sheet():
    credentials_dict = session.get('credentials')
    if not credentials_dict:
        return redirect(url_for('login'))

    credentials = Credentials(
        token=credentials_dict['token'],
        refresh_token=credentials_dict['refresh_token'],
        token_uri=credentials_dict['token_uri'],
        client_id=credentials_dict['client_id'],
        client_secret=credentials_dict['client_secret'],
        scopes=credentials_dict['scopes']
    )

    try:
        drive_service = build('drive', 'v3', credentials=credentials)
        query = "mimeType='application/vnd.google-apps.spreadsheet'"
        results = drive_service.files().list(q=query, fields="files(id, name)").execute()
        sheet_files = results.get('files', [])

        if not sheet_files:
            return 'No Google Sheets files found.'

        # Truyền token vào template
        return render_template('select_sheet.html', sheets=sheet_files, oauth_token=credentials.token)
    except HttpError as error:
        return f'An error occurred while listing files: {error}'


@app.route('/read_sheet/<file_id>')
def read_sheet(file_id):
    credentials_dict = session.get('credentials')
    if not credentials_dict:
        return redirect(url_for('login'))

    credentials = Credentials(
        token=credentials_dict['token'],
        refresh_token=credentials_dict['refresh_token'],
        token_uri=credentials_dict['token_uri'],
        client_id=credentials_dict['client_id'],
        client_secret=credentials_dict['client_secret'],
        scopes=credentials_dict['scopes']
    )

    try:
        sheets_service = build('sheets', 'v4', credentials=credentials)

        # Lấy danh sách các sheet để kiểm tra tên sheet
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=file_id).execute()
        sheet_names = [sheet['properties']['title'] for sheet in spreadsheet.get('sheets', [])]

        # Đọc dữ liệu từ sheet đầu tiên
        if sheet_names:
            sheet_name = sheet_names[0]
            range_name = f"{sheet_name}!A1:Z1000"
            sheet = sheets_service.spreadsheets().values().get(spreadsheetId=file_id, range=range_name).execute()
            values = sheet.get('values', [])
            print(values)
        else:
            return "No sheets found in the spreadsheet."

        # Tạo sheet mới bên cạnh sheet đầu tiên và chèn dữ liệu
        create_sheet_next_to_existing(sheets_service, file_id, sheet_name, values)

        # Lưu dữ liệu vào MongoDB
        if values:
            for row in values:
                collection.insert_one({'data': row})
        # Trả về dữ liệu qua template
        createDiagram(values, filename='static/gantt_chart.png')
        return render_template('show_data.html', data=values)
    except HttpError as error:
        print(f"An error occurred while reading sheet: {error}")
        return f'An error occurred while reading sheet: {error}'
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return f'An unexpected error occurred: {e}'


def credentials_to_dict(credentials):
    return {
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret,
        'scopes': credentials.scopes
    }


def create_sheet_next_to_existing(sheets_service, spreadsheet_id, existing_sheet_name, values):
    try:
        # Lấy danh sách các sheet hiện tại trong bảng tính
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        sheet_names = [sheet['properties']['title'] for sheet in sheets]

        # Kiểm tra nếu tên sheet hiện tại có trong danh sách
        if existing_sheet_name not in sheet_names:
            raise ValueError(f"Sheet named '{existing_sheet_name}' not found in the spreadsheet.")

        # Tìm vị trí của sheet hiện tại
        position = sheet_names.index(existing_sheet_name)

        # Tạo sheet mới
        new_sheet_title = f'New Sheet {random.randint(1000, 9999)}'
        requests = [
            {
                'addSheet': {
                    'properties': {
                        'title': new_sheet_title
                    }
                }
            }
        ]

        # Gửi yêu cầu tạo sheet mới
        response = sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': requests}
        ).execute()

        new_sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']

        # Di chuyển sheet mới đến vị trí sau sheet hiện tại
        move_request = {
            'requests': [
                {
                    'updateSheetProperties': {
                        'properties': {
                            'sheetId': new_sheet_id,
                            'index': position + 1
                        },
                        'fields': 'index'
                    }
                }
            ]
        }

        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=move_request
        ).execute()

        # Chèn dữ liệu vào sheet mới
        range_name = f'{new_sheet_title}!A1'
        body = {
            'values': values
        }
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='RAW',
            body=body
        ).execute()

        # Tạo biểu đồ trong sheet mới
        chart_request = {
            "requests": [
                {
                    "addChart": {
                        "chart": {
                            "spec": {
                                "title": "Sample Chart",
                                "basicChart": {
                                    "chartType": "LINE",
                                    "legendPosition": "BOTTOM_LEGEND",
                                    "axis": [
                                        {"position": "BOTTOM_AXIS", "title": "X-Axis"},
                                        {"position": "LEFT_AXIS", "title": "Y-Axis"}
                                    ],
                                    "domains": [
                                        {
                                            "domain": {
                                                "sourceRange": {
                                                    "sources": [
                                                        {
                                                            "sheetId": new_sheet_id,
                                                            "startRowIndex": 1,
                                                            "startColumnIndex": 0,
                                                            "endColumnIndex": 1
                                                        }
                                                    ]
                                                }
                                            }
                                        }
                                    ],
                                    "series": [
                                        {
                                            "series": {
                                                "sourceRange": {
                                                    "sources": [
                                                        {
                                                            "sheetId": new_sheet_id,
                                                            "startRowIndex": 1,
                                                            "startColumnIndex": 1,
                                                            "endColumnIndex": 2
                                                        }
                                                    ]
                                                }
                                            }
                                        }
                                    ]
                                }
                            },
                            "position": {
                                "overlayPosition": {
                                    "anchorCell": {
                                        "sheetId": new_sheet_id,
                                        "rowIndex": 0,
                                        "columnIndex": 4
                                    },
                                    "offsetXPixels": 0,
                                    "offsetYPixels": 0
                                }
                            }
                        }
                    }
                }
            ]
        }

        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": chart_request}
        ).execute()

        print(f'Sheet "{new_sheet_title}" created, data inserted, and chart created successfully.')
    except HttpError as error:
        print(f'An error occurred: {error}')
    except ValueError as ve:
        print(f'Value error: {ve}')
    except Exception as e:
        print(f'An unexpected error occurred: {e}')


def createDiagram(values, filename='static/gantt_chart.png'):
    data = []
    for row in values:
        if len(row) > 6 and row[2] and row[2][0].isdigit():
            task = row[3] if row[3] else data[-1][0]
            subtask = row[4]
            hours = float(row[6])
            data.append([task, subtask, hours])

    df = pd.DataFrame(data, columns=['Task', 'Subtask', 'Duration'])

    # Thêm cột Start và End cho mỗi subtask
    start_date = dt.datetime(2024, 7, 1)
    df['Start'] = [start_date + dt.timedelta(days=i // 3) for i in range(len(df))]
    df['End'] = df['Start'] + df['Duration'].apply(lambda x: dt.timedelta(hours=x))
    df['Duration'] = df['End'] - df['Start']

    # Tạo danh sách các màu sắc
    unique_subtasks = df['Subtask'].unique()
    colors = plt.cm.tab20(range(len(unique_subtasks)))
    color_map = dict(zip(unique_subtasks, colors))

    # Thiết lập figure và axis
    fig, ax = plt.subplots(figsize=(12, 8))

    # Vẽ các task trên biểu đồ Gantt
    patches = []
    for i, (task, subtask, start, end) in enumerate(
            zip(df['Task'], df['Subtask'], df['Start'], df['End'])):
        color = color_map[subtask]
        patch = ax.barh(task, (end - start).total_seconds() / 3600, left=start, color=color, edgecolor='black')
        patches.append(patch)

    # Thiết lập tiêu đề
    plt.title('Project Management Schedule of Project X', fontsize=15)

    # Định dạng trục x
    ax.xaxis_date()
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))

    # Thay đổi định dạng trục y và thêm lưới
    plt.gca().invert_yaxis()
    ax.xaxis.grid(True, alpha=0.5)

    # Thêm chú giải
    labels = [subtask for subtask in unique_subtasks]
    handles = [patch for patch in patches]
    ax.legend(handles, labels, fontsize=5)

    # Lưu hình ảnh
    plt.tight_layout()
    plt.savefig(filename)
    plt.close()

if __name__ == '__main__':
    app.run(port=5000, debug=True, ssl_context='adhoc')
