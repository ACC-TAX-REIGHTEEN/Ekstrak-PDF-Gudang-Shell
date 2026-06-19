import os
import base64
import configparser
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def get_gmail_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print("--> Error: File credentials.json tidak ditemukan")
                return None
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)

def load_config():
    config = configparser.ConfigParser()
    if not os.path.exists('gmail.conf'):
        print("--> CRITICAL: File gmail.conf tidak ditemukan")
        return None

    config.read('gmail.conf')
    
    try:
        cfg = {}
        cfg['output_folder'] = config['DEFAULT'].get('output_folder', 'download')
        cfg['query'] = config['SEARCH_CONFIG'].get('gmail_query', '').replace("'", "").replace('"', '')
        cfg['filename_filter'] = config['SEARCH_CONFIG'].get('filename_must_contain', '').strip()

        s_start = config['SEARCH_CONFIG'].get('strict_start_date', '').strip()
        s_end = config['SEARCH_CONFIG'].get('strict_end_date', '').strip()

        cfg['strict_start'] = datetime.strptime(s_start, "%Y-%m-%d") if s_start else None
        cfg['strict_end']   = datetime.strptime(s_end, "%Y-%m-%d") if s_end else None

        return cfg
    except Exception as e:
        print(f"--> Error pada gmail.conf: {e}")
        return None

def download_attachments(service, cfg):
    folder = cfg['output_folder']
    if not os.path.exists(folder):
        os.makedirs(folder)

    print("--> MEMULAI DOWNLOAD")
    print(f"--> Query API Gmail : {cfg['query']}")
    print(f"--> Filter Strict   : {cfg['strict_start']} s/d {cfg['strict_end']}")
    print("--> ------------------------")

    try:
        count_download = 0
        count_skip = 0
        count_processed_emails = 0
        next_page_token = None
        processed_threads = set()

        while True:
            results = service.users().messages().list(
                userId='me', 
                q=cfg['query'], 
                pageToken=next_page_token
            ).execute()
            
            messages = results.get('messages', [])

            if not messages:
                if count_processed_emails == 0:
                    print("--> Gmail API tidak menemukan email yang cocok")
                break

            for msg in messages:
                thread_id = msg['threadId']
                
                if thread_id in processed_threads:
                    continue
                processed_threads.add(thread_id)
                
                try:
                    thread_detail = service.users().threads().get(userId='me', id=thread_id).execute()
                    thread_messages = thread_detail.get('messages', [])
                except Exception:
                    continue

                for msg_detail in thread_messages:
                    count_processed_emails += 1
                    
                    msg_date_ms = int(msg_detail['internalDate'])
                    msg_date = datetime.fromtimestamp(msg_date_ms / 1000.0)
                    
                    if cfg['strict_start'] and msg_date < cfg['strict_start']:
                        count_skip += 1
                        continue
                    
                    if cfg['strict_end'] and msg_date >= cfg['strict_end']:
                        count_skip += 1
                        continue

                    payload = msg_detail.get('payload', {})
                    parts = payload.get('parts', [])

                    if not parts and 'body' in payload:
                        parts = [payload]

                    all_parts = []
                    def extract_parts(parts_list):
                        for p in parts_list:
                            if p.get('filename'):
                                all_parts.append(p)
                            if p.get('parts'):
                                extract_parts(p['parts'])
                    
                    extract_parts(parts)

                    for part in all_parts:
                        filename = part.get('filename')
                        
                        if filename:
                            if not filename.lower().endswith('.pdf'):
                                continue

                            if cfg['filename_filter'] and cfg['filename_filter'] not in filename:
                                continue

                            data = None
                            if 'data' in part['body']:
                                data = part['body']['data']
                            elif 'attachmentId' in part['body']:
                                att_id = part['body']['attachmentId']
                                try:
                                    att = service.users().messages().attachments().get(
                                        userId='me', messageId=msg_detail['id'], id=att_id).execute()
                                    data = att['data']
                                except Exception:
                                    continue
                            
                            if data:
                                file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                                
                                date_prefix = msg_date.strftime("%Y-%m-%d")
                                clean_name = "".join([c for c in filename if c.isalpha() or c.isdigit() or c in "._- "]).strip()
                                final_filename = f"{date_prefix}_{msg_detail['id']}_{clean_name}"
                                path = os.path.join(folder, final_filename)
                                
                                with open(path, 'wb') as f:
                                    f.write(file_data)
                                
                                print(f"--> [OK] {date_prefix} | {clean_name}")
                                count_download += 1
            
            next_page_token = results.get('nextPageToken')
            if not next_page_token:
                break
        
        print(f"--> Selesai Memproses total {count_processed_emails} email dalam thread terkait")
        print(f"--> {count_download} file berhasil diunduh")
        print(f"--> {count_skip} email dilewati karena filter tanggal")

    except Exception as error:
        print(f"--> Terjadi error fatal: {error}")

if __name__ == '__main__':
    config_data = load_config()
    
    if config_data:
        service = get_gmail_service()
        if service:
            download_attachments(service, config_data)