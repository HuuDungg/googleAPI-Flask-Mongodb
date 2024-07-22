from flask import Flask, redirect, request, session, url_for, render_template
from flask_session import Session
from redis import Redis
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from pymongo import MongoClient
import os

app = Flask(__name__)
app.secret_key = 'huudungisthebest'
app.config['SESSION_TYPE'] = 'redis'
app.config['SESSION_REDIS'] = Redis.from_url('redis://redis-12375.c12.us-east-1-4.ec2.redns.redis-cloud.com:12375')
Session(app)

# URL của bạn
LOCAL_URL = 'https://google-api-flask-mongodb.vercel.app'

# Cấu hình OAuth 2.0
CLIENT_SECRETS_FILE = 'client_secret.json'
SCOPES = [
    'openid',
    'https://www.googleapis.com/auth/userinfo.email',
    'https://www.googleapis.com/auth/userinfo.profile',
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/spreadsheets.readonly',
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
        redirect_uri=f'{LOCAL_URL}/oauth2callback'
    )
    authorization_url, state = flow.authorization_url(access_type='offline', include_granted_scopes='true')
    session['state'] = state
    return redirect(authorization_url)

@app.route('/oauth2callback')
def oauth2callback():
    state = session.get('state')
    if not state:
        return 'State not found. Possible session expiration or tampering.'

    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE,
        scopes=SCOPES,
        state=state,
        redirect_uri=f'{LOCAL_URL}/oauth2callback'
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
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=file_id).execute()
        sheet_names = [sheet['properties']['title'] for sheet in spreadsheet.get('sheets', [])]

        if sheet_names:
            sheet_name = sheet_names[0]
            range_name = f"{sheet_name}!A1:Z1000"
            sheet = sheets_service.spreadsheets().values().get(spreadsheetId=file_id, range=range_name).execute()
            values = sheet.get('values', [])
        else:
            return "No sheets found in the spreadsheet."

        if values:
            for row in values:
                collection.insert_one({'data': row})

        return render_template('show_data.html', data=values)
    except HttpError as error:
        return f'An error occurred while reading sheet: {error}'
    except Exception as e:
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

if __name__ == '__main__':
    app.run(port=5000, debug=True, ssl_context='adhoc')
