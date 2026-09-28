import json, re, ssl, urllib.parse, urllib.request, urllib.error, secrets, imaplib, email, sqlite3, os
from datetime import datetime, timezone
from email import policy
from email.header import decode_header, make_header
from pathlib import Path

ROOT = Path(__file__).resolve().parent

def _user_data_dir():
    base = Path(os.environ.get('LOCALAPPDATA') or os.environ.get('APPDATA') or Path.home())
    d = base / 'RyanNguyen_ShopeeVoucher' / 'mail_native'
    d.mkdir(parents=True, exist_ok=True)
    return d

DB_PATH = _user_data_dir() / 'outlook_data.db'
SESSIONS = {}
TOKEN_URL = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token'
GRAPH = 'https://graph.microsoft.com/v1.0'


def now_iso():
    return datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')


def db_conn():
    con = sqlite3.connect(DB_PATH, timeout=30)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    with db_conn() as con:
        con.executescript('''
        CREATE TABLE IF NOT EXISTS accounts (
            email TEXT PRIMARY KEY COLLATE NOCASE,
            password TEXT NOT NULL DEFAULT '',
            refresh_token TEXT NOT NULL,
            client_id TEXT NOT NULL,
            method TEXT NOT NULL DEFAULT '',
            scope TEXT NOT NULL DEFAULT '',
            last_status TEXT NOT NULL DEFAULT 'SAVED',
            last_error TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            token_updated_at TEXT NOT NULL DEFAULT '',
            last_read_at TEXT NOT NULL DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS messages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            account_email TEXT NOT NULL,
            remote_id TEXT NOT NULL,
            subject TEXT NOT NULL DEFAULT '',
            from_name TEXT NOT NULL DEFAULT '',
            from_address TEXT NOT NULL DEFAULT '',
            received_at TEXT NOT NULL DEFAULT '',
            body_type TEXT NOT NULL DEFAULT 'text',
            body TEXT NOT NULL DEFAULT '',
            body_preview TEXT NOT NULL DEFAULT '',
            shopee_username TEXT NOT NULL DEFAULT '',
            username_manual INTEGER NOT NULL DEFAULT 0,
            shopee_otp TEXT NOT NULL DEFAULT '',
            verify_link TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            UNIQUE(account_email, remote_id)
        );
        CREATE INDEX IF NOT EXISTS idx_messages_account ON messages(account_email);
        CREATE INDEX IF NOT EXISTS idx_messages_username ON messages(shopee_username);
        CREATE INDEX IF NOT EXISTS idx_messages_otp ON messages(shopee_otp);
        ''')


def parse_full_line(raw: str):
    line = (raw or '').strip().splitlines()[0].strip()
    parts = line.split('|')
    if len(parts) < 4:
        raise ValueError('Định dạng cần 4 cột: email|password|refresh_token|client_id')
    email_addr = parts[0].strip()
    password = parts[1].strip()
    refresh_token = '|'.join(parts[2:-1]).strip() if len(parts) > 4 else parts[2].strip()
    client_id = parts[-1].strip()
    if not re.fullmatch(r'[^@\s]+@[^@\s]+\.[^@\s]+', email_addr):
        raise ValueError('Cột 1 không phải email hợp lệ.')
    if not refresh_token:
        raise ValueError('Thiếu refresh_token ở cột 3.')
    if not re.fullmatch(r'[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}', client_id):
        raise ValueError('Cột cuối không phải client_id dạng UUID.')
    return {'email': email_addr, 'password': password, 'refresh_token': refresh_token, 'client_id': client_id}


def full_line(acct):
    return '|'.join([
        str(acct.get('email') or ''),
        str(acct.get('password') or ''),
        str(acct.get('refresh_token') or ''),
        str(acct.get('client_id') or ''),
    ])


def get_account(email_addr):
    with db_conn() as con:
        r = con.execute('SELECT * FROM accounts WHERE email=?', (email_addr,)).fetchone()
        return dict(r) if r else None


def upsert_account(acct, method=None, scope=None, status=None, error=None, token_changed=False, read_ok=False):
    old = get_account(acct['email'])
    now = now_iso()
    created = old['created_at'] if old else now
    method_v = method if method is not None else (old['method'] if old else '')
    scope_v = scope if scope is not None else (old['scope'] if old else '')
    status_v = status if status is not None else (old['last_status'] if old else 'SAVED')
    error_v = error if error is not None else (old['last_error'] if old else '')
    token_updated_at = now if token_changed else (old['token_updated_at'] if old else '')
    last_read_at = now if read_ok else (old['last_read_at'] if old else '')
    with db_conn() as con:
        con.execute('''
        INSERT INTO accounts(email,password,refresh_token,client_id,method,scope,last_status,last_error,created_at,updated_at,token_updated_at,last_read_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(email) DO UPDATE SET
          password=excluded.password,
          refresh_token=excluded.refresh_token,
          client_id=excluded.client_id,
          method=excluded.method,
          scope=excluded.scope,
          last_status=excluded.last_status,
          last_error=excluded.last_error,
          updated_at=excluded.updated_at,
          token_updated_at=excluded.token_updated_at,
          last_read_at=excluded.last_read_at
        ''', (acct['email'], acct.get('password',''), acct['refresh_token'], acct['client_id'], method_v, scope_v,
              status_v, error_v, created, now, token_updated_at, last_read_at))
    return get_account(acct['email'])


def list_accounts():
    with db_conn() as con:
        rows = con.execute('SELECT * FROM accounts ORDER BY updated_at DESC, email COLLATE NOCASE').fetchall()
    out = []
    for r in rows:
        d = dict(r)
        d['full_line'] = full_line(d)
        out.append(d)
    return out


def delete_accounts(emails):
    emails = [str(x).strip() for x in emails if str(x).strip()]
    if not emails:
        return 0
    q = ','.join('?' for _ in emails)
    with db_conn() as con:
        cur = con.execute(f'DELETE FROM accounts WHERE email IN ({q})', emails)
        return cur.rowcount


def save_message(payload):
    msg = payload.get('message') or {}
    account_email = str(payload.get('account_email') or '').strip()
    if not account_email:
        raise ValueError('Thiếu email tài khoản của mail cần lưu.')
    remote_id = str(msg.get('id') or '').strip()
    if not remote_id:
        remote_id = 'local:' + secrets.token_hex(12)
    frm = msg.get('from') or {}
    addr = frm.get('emailAddress') or {}
    body = msg.get('body') or {}
    now = now_iso()
    username = str(payload.get('shopee_username') or '')
    username_manual = 1 if payload.get('username_manual') else 0
    otp = str(payload.get('shopee_otp') or '')
    verify_link = str(payload.get('verify_link') or '')
    with db_conn() as con:
        old = con.execute('SELECT * FROM messages WHERE account_email=? AND remote_id=?', (account_email, remote_id)).fetchone()
        if old and int(old['username_manual'] or 0) == 1 and not username_manual:
            username = old['shopee_username']
            username_manual = 1
        created = old['created_at'] if old else now
        con.execute('''
        INSERT INTO messages(account_email,remote_id,subject,from_name,from_address,received_at,body_type,body,body_preview,
                             shopee_username,username_manual,shopee_otp,verify_link,created_at,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(account_email,remote_id) DO UPDATE SET
          subject=excluded.subject,
          from_name=excluded.from_name,
          from_address=excluded.from_address,
          received_at=excluded.received_at,
          body_type=excluded.body_type,
          body=excluded.body,
          body_preview=excluded.body_preview,
          shopee_username=excluded.shopee_username,
          username_manual=excluded.username_manual,
          shopee_otp=excluded.shopee_otp,
          verify_link=excluded.verify_link,
          updated_at=excluded.updated_at
        ''', (
            account_email, remote_id, str(msg.get('subject') or ''), str(addr.get('name') or ''), str(addr.get('address') or ''),
            str(msg.get('receivedDateTime') or ''), str(body.get('contentType') or 'text'), str(body.get('content') or ''),
            str(msg.get('bodyPreview') or ''), username, username_manual, otp, verify_link, created, now
        ))
        row = con.execute('SELECT * FROM messages WHERE account_email=? AND remote_id=?', (account_email, remote_id)).fetchone()
    return dict(row)


def list_messages():
    with db_conn() as con:
        rows = con.execute('SELECT * FROM messages ORDER BY COALESCE(NULLIF(received_at,\'\'), created_at) DESC, id DESC').fetchall()
    return [dict(r) for r in rows]


def update_message(payload):
    mid = int(payload.get('id') or 0)
    if not mid:
        raise ValueError('Thiếu ID mail đã lưu.')
    fields = []
    vals = []
    if 'shopee_username' in payload:
        fields += ['shopee_username=?', 'username_manual=?']
        vals += [str(payload.get('shopee_username') or '').strip(), 1 if payload.get('username_manual', True) else 0]
    if 'shopee_otp' in payload:
        fields.append('shopee_otp=?'); vals.append(str(payload.get('shopee_otp') or '').strip())
    if 'verify_link' in payload:
        fields.append('verify_link=?'); vals.append(str(payload.get('verify_link') or '').strip())
    if not fields:
        return None
    fields.append('updated_at=?'); vals.append(now_iso()); vals.append(mid)
    with db_conn() as con:
        con.execute('UPDATE messages SET ' + ','.join(fields) + ' WHERE id=?', vals)
        row = con.execute('SELECT * FROM messages WHERE id=?', (mid,)).fetchone()
    return dict(row) if row else None


def delete_messages(ids):
    ids = [int(x) for x in ids if str(x).isdigit()]
    if not ids:
        return 0
    q = ','.join('?' for _ in ids)
    with db_conn() as con:
        cur = con.execute(f'DELETE FROM messages WHERE id IN ({q})', ids)
        return cur.rowcount


def clear_local_data():
    with db_conn() as con:
        con.execute('DELETE FROM messages')
        con.execute('DELETE FROM accounts')


def import_backup(data):
    accounts = data.get('accounts') or []
    messages = data.get('messages') or []
    ac = 0; mc = 0
    for a in accounts:
        try:
            acct = {
                'email': str(a.get('email') or '').strip(),
                'password': str(a.get('password') or ''),
                'refresh_token': str(a.get('refresh_token') or ''),
                'client_id': str(a.get('client_id') or ''),
            }
            parse_full_line(full_line(acct))
            upsert_account(acct, method=str(a.get('method') or ''), scope=str(a.get('scope') or ''),
                           status=str(a.get('last_status') or 'IMPORTED'), error=str(a.get('last_error') or ''))
            ac += 1
        except Exception:
            pass
    with db_conn() as con:
        for m in messages:
            try:
                account_email = str(m.get('account_email') or '').strip()
                remote_id = str(m.get('remote_id') or '').strip()
                if not account_email or not remote_id:
                    continue
                now = now_iso()
                con.execute('''
                INSERT INTO messages(account_email,remote_id,subject,from_name,from_address,received_at,body_type,body,body_preview,
                                     shopee_username,username_manual,shopee_otp,verify_link,created_at,updated_at)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(account_email,remote_id) DO UPDATE SET
                  subject=excluded.subject,from_name=excluded.from_name,from_address=excluded.from_address,
                  received_at=excluded.received_at,body_type=excluded.body_type,body=excluded.body,body_preview=excluded.body_preview,
                  shopee_username=excluded.shopee_username,username_manual=excluded.username_manual,
                  shopee_otp=excluded.shopee_otp,verify_link=excluded.verify_link,updated_at=excluded.updated_at
                ''', (
                    account_email, remote_id, str(m.get('subject') or ''), str(m.get('from_name') or ''), str(m.get('from_address') or ''),
                    str(m.get('received_at') or ''), str(m.get('body_type') or 'text'), str(m.get('body') or ''), str(m.get('body_preview') or ''),
                    str(m.get('shopee_username') or ''), int(m.get('username_manual') or 0), str(m.get('shopee_otp') or ''),
                    str(m.get('verify_link') or ''), str(m.get('created_at') or now), now
                ))
                mc += 1
            except Exception:
                pass
    return {'accounts': ac, 'messages': mc}


def safe_json_loads(body):
    try:
        return json.loads(body)
    except Exception:
        return {'error': 'non_json_response', 'error_description': body[:1600]}


def post_form(url, form):
    data = urllib.parse.urlencode(form).encode('utf-8')
    req = urllib.request.Request(url, data=data, method='POST', headers={
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': 'application/json',
        'User-Agent': 'OutlookFullTokenToolV6/1.0',
    })
    try:
        with urllib.request.urlopen(req, timeout=30, context=ssl.create_default_context()) as r:
            body = r.read().decode('utf-8', 'replace')
            return r.status, safe_json_loads(body)
    except urllib.error.HTTPError as e:
        body = e.read().decode('utf-8', 'replace')
        return e.code, safe_json_loads(body)
    except Exception as e:
        return 0, {'error': 'network_error', 'error_description': str(e)}


def get_json(url, access_token):
    req = urllib.request.Request(url, headers={
        'Authorization': 'Bearer ' + access_token,
        'Accept': 'application/json',
        'Prefer': 'outlook.body-content-type="html"',
        'User-Agent': 'OutlookFullTokenToolV6/1.0',
    })
    try:
        with urllib.request.urlopen(req, timeout=30, context=ssl.create_default_context()) as r:
            body = r.read().decode('utf-8', 'replace')
            return r.status, safe_json_loads(body)
    except urllib.error.HTTPError as e:
        body = e.read().decode('utf-8', 'replace')
        obj = safe_json_loads(body)
        if 'error' not in obj or not isinstance(obj.get('error'), dict):
            obj = {'error': {'message': obj.get('error_description') or body[:1600]}}
        return e.code, obj
    except Exception as e:
        return 0, {'error': {'message': str(e)}}


def exchange(refresh_token, client_id, scope):
    form = {'client_id': client_id, 'grant_type': 'refresh_token', 'refresh_token': refresh_token}
    if scope:
        form['scope'] = scope
    status, obj = post_form(TOKEN_URL, form)
    if status == 200 and obj.get('access_token'):
        return True, obj, status
    return False, obj, status


def graph_latest(token):
    sel = 'id,subject,from,receivedDateTime,isRead,bodyPreview,body,hasAttachments'
    params = {'$top': 1, '$orderby': 'receivedDateTime desc', '$select': sel}
    url = GRAPH + '/me/mailFolders/inbox/messages?' + urllib.parse.urlencode(params)
    status, obj = get_json(url, token)
    if status != 200:
        return status, None, obj
    rows = obj.get('value') or []
    return status, (rows[0] if rows else None), obj

def message_search_text(msg):
    frm = (msg or {}).get('from') or {}
    addr = frm.get('emailAddress') or {}
    body = (msg or {}).get('body') or {}
    return ' '.join([
        str((msg or {}).get('subject') or ''),
        str(addr.get('name') or ''),
        str(addr.get('address') or ''),
        str((msg or {}).get('bodyPreview') or ''),
        str(body.get('content') or ''),
    ]).lower()


def is_shopee_message(msg):
    text = message_search_text(msg)
    return ('shopee' in text) or ('shp.ee' in text)


def is_shopee_verification_message(msg):
    text = message_search_text(msg)
    if not (('shopee' in text) or ('shp.ee' in text)):
        return False
    verify_words = (
        'xác nhận', 'xac nhan', 'xác thực', 'xac thuc', 'xác minh', 'xac minh',
        'đăng nhập', 'dang nhap', 'truy cập', 'truy cap', 'ai đó đang cố gắng',
        'verify', 'verification', 'confirm', 'security', 'login', 'log in', 'authorize',
    )
    if not any(word in text for word in verify_words):
        return False
    link_words = ('vn.shp.ee/dlink/', 'shp.ee/dlink/', 'sendgrid.net/ls/click', 'http://', 'https://')
    return any(word in text for word in link_words)


def graph_latest_shopee(token, limit=30):
    sel = 'id,subject,from,receivedDateTime,isRead,bodyPreview,body,hasAttachments'
    params = {'$top': max(1, min(int(limit or 30), 50)), '$orderby': 'receivedDateTime desc', '$select': sel}
    url = GRAPH + '/me/mailFolders/inbox/messages?' + urllib.parse.urlencode(params)
    status, obj = get_json(url, token)
    if status != 200:
        return status, None, obj
    for row in obj.get('value') or []:
        if is_shopee_message(row):
            return status, row, obj
    return status, None, obj


def graph_latest_shopee_verification(token, limit=50):
    sel = 'id,subject,from,receivedDateTime,isRead,bodyPreview,body,hasAttachments'
    params = {'$top': max(1, min(int(limit or 50), 50)), '$orderby': 'receivedDateTime desc', '$select': sel}
    url = GRAPH + '/me/mailFolders/inbox/messages?' + urllib.parse.urlencode(params)
    status, obj = get_json(url, token)
    if status != 200:
        return status, None, obj
    for row in obj.get('value') or []:
        if is_shopee_verification_message(row):
            return status, row, obj
    return status, None, obj


def decode_mime_header(value):
    try:
        return str(make_header(decode_header(value or '')))
    except Exception:
        return value or ''


def message_body_parts(msg):
    html = ''
    text = ''
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = (part.get('Content-Disposition') or '').lower()
            if 'attachment' in disp:
                continue
            try:
                content = part.get_content()
            except Exception:
                payload = part.get_payload(decode=True) or b''
                charset = part.get_content_charset() or 'utf-8'
                content = payload.decode(charset, 'replace')
            if ctype == 'text/html' and not html:
                html = str(content)
            elif ctype == 'text/plain' and not text:
                text = str(content)
    else:
        try:
            content = msg.get_content()
        except Exception:
            payload = msg.get_payload(decode=True) or b''
            charset = msg.get_content_charset() or 'utf-8'
            content = payload.decode(charset, 'replace')
        if msg.get_content_type() == 'text/html':
            html = str(content)
        else:
            text = str(content)
    return text, html


def imap_latest(email_addr, access_token):
    imap = None
    try:
        imap = imaplib.IMAP4_SSL('outlook.office365.com', 993, ssl_context=ssl.create_default_context(), timeout=30)
        auth = f'user={email_addr}\x01auth=Bearer {access_token}\x01\x01'.encode('utf-8')
        imap.authenticate('XOAUTH2', lambda _: auth)
        typ, _ = imap.select('INBOX', readonly=True)
        if typ != 'OK':
            raise RuntimeError('Không mở được INBOX qua IMAP.')
        typ, data = imap.search(None, 'ALL')
        if typ != 'OK':
            raise RuntimeError('IMAP SEARCH thất bại.')
        ids = (data[0] or b'').split()
        if not ids:
            return None
        latest_id = ids[-1]
        typ, fetched = imap.fetch(latest_id, '(RFC822)')
        if typ != 'OK' or not fetched:
            raise RuntimeError('IMAP FETCH mail mới nhất thất bại.')
        raw = None
        for item in fetched:
            if isinstance(item, tuple) and len(item) > 1 and isinstance(item[1], (bytes, bytearray)):
                raw = bytes(item[1]); break
        if not raw:
            raise RuntimeError('IMAP không trả nội dung RFC822.')
        msg = email.message_from_bytes(raw, policy=policy.default)
        text, html = message_body_parts(msg)
        return {
            'id': latest_id.decode('ascii', 'replace'),
            'subject': decode_mime_header(msg.get('Subject')),
            'from': {'emailAddress': {'name': decode_mime_header(msg.get('From')), 'address': decode_mime_header(msg.get('From'))}},
            'receivedDateTime': msg.get('Date') or '',
            'bodyPreview': (text or re.sub(r'<[^>]+>', ' ', html))[:500],
            'body': {'contentType': 'html' if html else 'text', 'content': html or text},
            'hasAttachments': any((p.get('Content-Disposition') or '').lower().startswith('attachment') for p in msg.walk()) if msg.is_multipart() else False,
        }
    finally:
        if imap is not None:
            try: imap.logout()
            except Exception: pass


def imap_latest_shopee(email_addr, access_token, limit=30, verification_only=False):
    imap = None
    try:
        imap = imaplib.IMAP4_SSL('outlook.office365.com', 993, ssl_context=ssl.create_default_context(), timeout=30)
        auth = f'user={email_addr}\x01auth=Bearer {access_token}\x01\x01'.encode('utf-8')
        imap.authenticate('XOAUTH2', lambda _: auth)
        typ, _ = imap.select('INBOX', readonly=True)
        if typ != 'OK':
            raise RuntimeError('Không mở được INBOX qua IMAP.')
        typ, data = imap.search(None, 'ALL')
        if typ != 'OK':
            raise RuntimeError('IMAP SEARCH thất bại.')
        ids = (data[0] or b'').split()
        for latest_id in reversed(ids[-max(1, min(int(limit or 30), 50)):]):
            typ, fetched = imap.fetch(latest_id, '(RFC822)')
            if typ != 'OK' or not fetched:
                continue
            raw = None
            for item in fetched:
                if isinstance(item, tuple) and len(item) > 1 and isinstance(item[1], (bytes, bytearray)):
                    raw = bytes(item[1]); break
            if not raw:
                continue
            msg = email.message_from_bytes(raw, policy=policy.default)
            text, html = message_body_parts(msg)
            row = {
                'id': latest_id.decode('ascii', 'replace'),
                'subject': decode_mime_header(msg.get('Subject')),
                'from': {'emailAddress': {'name': decode_mime_header(msg.get('From')), 'address': decode_mime_header(msg.get('From'))}},
                'receivedDateTime': msg.get('Date') or '',
                'bodyPreview': (text or re.sub(r'<[^>]+>', ' ', html))[:500],
                'body': {'contentType': 'html' if html else 'text', 'content': html or text},
                'hasAttachments': any((p.get('Content-Disposition') or '').lower().startswith('attachment') for p in msg.walk()) if msg.is_multipart() else False,
            }
            predicate = is_shopee_verification_message if verification_only else is_shopee_message
            if predicate(row):
                return row
        return None
    finally:
        if imap is not None:
            try: imap.logout()
            except Exception: pass


def read_latest_shopee_for_account(acct):
    attempts = []
    ok, tok, ts = exchange(acct['refresh_token'], acct['client_id'], 'https://graph.microsoft.com/Mail.Read offline_access')
    if ok:
        gs, message, gobj = graph_latest_shopee(tok['access_token'], 30)
        if gs == 200:
            attempts.append({'step': 'Graph Shopee', 'ok': True, 'detail': 'Đã quét 30 mail gần nhất.'})
            return {'ok': True, 'method': 'GRAPH', 'message': message, 'access_token': tok['access_token'],
                    'refresh_token': tok.get('refresh_token') or acct['refresh_token'], 'scope': tok.get('scope',''),
                    'expires_in': tok.get('expires_in'), 'attempts': attempts}
        code, detail = short_error(gobj)
        attempts.append({'step':'Graph Shopee','ok':False,'error':f'Graph HTTP {gs or "network"} {code}','detail':detail[:900]})
    else:
        code, detail = short_error(tok)
        attempts.append({'step':'Graph token','ok':False,'error':f'HTTP {ts or "network"} {code}','detail':detail[:900]})

    ok, tok, ts = exchange(acct['refresh_token'], acct['client_id'], 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access')
    if ok:
        try:
            msg = imap_latest_shopee(acct['email'], tok['access_token'], 30)
            attempts.append({'step':'IMAP Shopee','ok':True,'detail':'Đã quét tối đa 30 mail gần nhất.'})
            return {'ok':True,'method':'IMAP','message':msg,'access_token':tok['access_token'],
                    'refresh_token':tok.get('refresh_token') or acct['refresh_token'],'scope':tok.get('scope',''),
                    'expires_in':tok.get('expires_in'),'attempts':attempts}
        except Exception as e:
            attempts.append({'step':'IMAP Shopee','ok':False,'error':type(e).__name__,'detail':str(e)[:1000]})
    else:
        code, detail = short_error(tok)
        attempts.append({'step':'IMAP token','ok':False,'error':f'HTTP {ts or "network"} {code}','detail':detail[:900]})
    return {'ok':False,'attempts':attempts,'message':'Không đọc được Inbox bằng Graph hoặc IMAP OAuth2.'}


def read_latest_shopee_verification_for_account(acct):
    attempts = []
    ok, tok, ts = exchange(acct['refresh_token'], acct['client_id'], 'https://graph.microsoft.com/Mail.Read offline_access')
    if ok:
        gs, message, gobj = graph_latest_shopee_verification(tok['access_token'], 50)
        if gs == 200:
            attempts.append({'step': 'Graph Shopee Verify', 'ok': True, 'detail': 'Đã quét 50 mail gần nhất để tìm link xác thực.'})
            return {'ok': True, 'method': 'GRAPH', 'message': message, 'access_token': tok['access_token'],
                    'refresh_token': tok.get('refresh_token') or acct['refresh_token'], 'scope': tok.get('scope',''),
                    'expires_in': tok.get('expires_in'), 'attempts': attempts}
        code, detail = short_error(gobj)
        attempts.append({'step':'Graph Shopee Verify','ok':False,'error':f'Graph HTTP {gs or "network"} {code}','detail':detail[:900]})
    else:
        code, detail = short_error(tok)
        attempts.append({'step':'Graph token','ok':False,'error':f'HTTP {ts or "network"} {code}','detail':detail[:900]})

    ok, tok, ts = exchange(acct['refresh_token'], acct['client_id'], 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access')
    if ok:
        try:
            msg = imap_latest_shopee(acct['email'], tok['access_token'], 50, verification_only=True)
            attempts.append({'step':'IMAP Shopee Verify','ok':True,'detail':'Đã quét tối đa 50 mail gần nhất để tìm link xác thực.'})
            return {'ok':True,'method':'IMAP','message':msg,'access_token':tok['access_token'],
                    'refresh_token':tok.get('refresh_token') or acct['refresh_token'],'scope':tok.get('scope',''),
                    'expires_in':tok.get('expires_in'),'attempts':attempts}
        except Exception as e:
            attempts.append({'step':'IMAP Shopee Verify','ok':False,'error':type(e).__name__,'detail':str(e)[:1000]})
    else:
        code, detail = short_error(tok)
        attempts.append({'step':'IMAP token','ok':False,'error':f'HTTP {ts or "network"} {code}','detail':detail[:900]})
    return {'ok':False,'attempts':attempts,'message':'Không đọc được Inbox bằng Graph hoặc IMAP OAuth2.'}


def short_error(obj):
    if isinstance(obj.get('error'), dict):
        return (obj['error'].get('code') or 'error', obj['error'].get('message') or '')
    return (str(obj.get('error') or 'error'), str(obj.get('error_description') or obj.get('message') or ''))


def read_latest_for_account(acct):
    attempts = []
    ok, tok, ts = exchange(acct['refresh_token'], acct['client_id'], 'https://graph.microsoft.com/Mail.Read offline_access')
    if ok:
        gs, message, gobj = graph_latest(tok['access_token'])
        if gs == 200:
            attempts.append({'step': 'Graph Mail.Read', 'ok': True, 'detail': 'Đọc được mail mới nhất.' if message else 'Inbox trống.'})
            return {'ok': True, 'method': 'GRAPH', 'message': message, 'access_token': tok['access_token'],
                    'refresh_token': tok.get('refresh_token') or acct['refresh_token'], 'scope': tok.get('scope',''),
                    'expires_in': tok.get('expires_in'), 'attempts': attempts}
        code, detail = short_error(gobj)
        attempts.append({'step':'Graph Mail.Read','ok':False,'error':f'Graph HTTP {gs or "network"} {code}','detail':detail[:900]})
    else:
        code, detail = short_error(tok)
        attempts.append({'step':'Graph token','ok':False,'error':f'HTTP {ts or "network"} {code}','detail':detail[:900]})

    ok, tok, ts = exchange(acct['refresh_token'], acct['client_id'], 'https://outlook.office.com/IMAP.AccessAsUser.All offline_access')
    if ok:
        try:
            msg = imap_latest(acct['email'], tok['access_token'])
            attempts.append({'step':'IMAP OAuth2','ok':True,'detail':'Đọc được mail mới nhất.' if msg else 'Inbox trống.'})
            return {'ok':True,'method':'IMAP','message':msg,'access_token':tok['access_token'],
                    'refresh_token':tok.get('refresh_token') or acct['refresh_token'],'scope':tok.get('scope',''),
                    'expires_in':tok.get('expires_in'),'attempts':attempts}
        except Exception as e:
            attempts.append({'step':'IMAP OAuth2','ok':False,'error':type(e).__name__,'detail':str(e)[:1000]})
    else:
        code, detail = short_error(tok)
        attempts.append({'step':'IMAP token','ok':False,'error':f'HTTP {ts or "network"} {code}','detail':detail[:900]})
    return {'ok':False,'attempts':attempts,'message':'Không đọc được mail bằng Graph hoặc IMAP OAuth2. Xem Nhật ký để biết lỗi thật từ Microsoft.'}


def renew_refresh_token(acct, preferred_method=''):
    tries = []
    candidates = []
    pm = (preferred_method or '').upper()
    if pm == 'GRAPH': candidates.append(('Graph Mail.Read','https://graph.microsoft.com/Mail.Read offline_access'))
    elif pm == 'IMAP': candidates.append(('IMAP OAuth2','https://outlook.office.com/IMAP.AccessAsUser.All offline_access'))
    candidates += [
        ('Quyền gốc của refresh token',''),
        ('Graph Mail.Read','https://graph.microsoft.com/Mail.Read offline_access'),
        ('IMAP OAuth2','https://outlook.office.com/IMAP.AccessAsUser.All offline_access')
    ]
    seen = set()
    for label, scope in candidates:
        key = scope or '<original>'
        if key in seen: continue
        seen.add(key)
        ok, tok, status = exchange(acct['refresh_token'], acct['client_id'], scope)
        if ok:
            new_rt = tok.get('refresh_token') or ''
            changed = bool(new_rt and new_rt != acct['refresh_token'])
            tries.append({'step':'Làm mới token — '+label,'ok':True,
                          'detail':'Microsoft trả refresh_token mới.' if changed else 'Đổi token thành công nhưng Microsoft không trả refresh_token khác; giữ token hiện tại.'})
            return {'ok':True,'access_token':tok.get('access_token',''),'refresh_token':new_rt or acct['refresh_token'],
                    'refresh_token_changed':changed,'scope':tok.get('scope','') or scope,'expires_in':tok.get('expires_in'),'attempts':tries}
        code, detail = short_error(tok)
        tries.append({'step':'Làm mới token — '+label,'ok':False,'error':f'HTTP {status or "network"} {code}','detail':detail[:1000]})
    return {'ok':False,'attempts':tries,'message':'Không làm mới được refresh token. Xem Nhật ký để biết lỗi Microsoft.'}


def resolve_account(data):
    email_addr = str(data.get('email') or '').strip()
    if email_addr:
        acct = get_account(email_addr)
        if not acct:
            raise ValueError('Không tìm thấy tài khoản đã lưu: ' + email_addr)
        return {'email':acct['email'],'password':acct['password'],'refresh_token':acct['refresh_token'],'client_id':acct['client_id']}
    return parse_full_line(data.get('full_line',''))



# Native backend only: no HTTP server, no HTML, no browser launch.
