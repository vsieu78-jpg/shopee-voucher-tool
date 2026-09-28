from __future__ import annotations

import csv
import html
import json
import os
import random
import re
import secrets
import sqlite3
import string
import threading
import time
import webbrowser
from html.parser import HTMLParser
from urllib.parse import urlsplit
from datetime import datetime, timezone
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
import tkinter as tk
from typing import Any

try:
    import httpx
except Exception:
    httpx = None

try:
    import outlook_native_backend as outlook_backend
except Exception:
    outlook_backend = None


MAILTM_API = "https://api.mail.tm"
TINYHOST_API = "https://tinyhost.shop"


def _now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def _fmt_time(v: str) -> str:
    if not v:
        return ""
    try:
        return datetime.fromisoformat(v.replace("Z", "+00:00")).astimezone().strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        return str(v)


def _user_data_dir() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or Path.home())
    d = base / "RyanNguyen_ShopeeVoucher" / "mail_native"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _html_to_text(value: str) -> str:
    s = html.unescape(str(value or ""))
    s = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", s)
    s = re.sub(r"(?i)<br\s*/?>", "\n", s)
    s = re.sub(r"(?i)</p\s*>", "\n", s)
    s = re.sub(r"(?s)<[^>]+>", " ", s)
    s = s.replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n\s*\n+", "\n\n", s)
    return s.strip()


def _clean_email_url(value: str) -> str:
    url = html.unescape(str(value or "").strip()).replace("&amp;", "&")
    return url.rstrip("\"'<>).,;!]}")


def _extract_email_urls(body: str = "", raw_html: str = "") -> list[str]:
    """Lay URL trong text/href, giu tracking link de co the copy/mo truc tiep."""
    raw = html.unescape(str(raw_html or ""))
    text = html.unescape(str(body or ""))
    found: list[str] = []
    href_re = re.compile(r"(?is)href\s*=\s*['\"](https?://[^'\"]+)['\"]")
    url_re = re.compile(r"https?://[^\s<>\"']+", re.I)
    for m in href_re.finditer(raw):
        found.append(_clean_email_url(m.group(1)))
    for source in (raw, text, _html_to_text(raw)):
        for m in url_re.finditer(source):
            found.append(_clean_email_url(m.group(0)))
    out: list[str] = []
    seen = set()
    for url in found:
        key = url.casefold()
        if url and key not in seen:
            seen.add(key)
            out.append(url)
    return out


def _verification_context(subject: str = "", body: str = "", raw_html: str = "") -> bool:
    text = " ".join([str(subject or ""), str(body or ""), _html_to_text(raw_html or "")]).casefold()
    shopee_hint = ("shopee" in text) or ("shp.ee" in text)
    verify_words = (
        "xác nhận", "xac nhan", "xác thực", "xac thuc", "xác minh", "xac minh",
        "đăng nhập", "dang nhap", "truy cập", "truy cap", "ai đó đang cố gắng",
        "verify", "verification", "confirm", "security", "login", "log in", "authorize",
    )
    return shopee_hint and any(word in text for word in verify_words)


class _ConfirmationAnchors(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.links = []
        self.href = None
        self.label = []

    def handle_starttag(self, tag, attrs):
        if tag.lower() == "a":
            self.href = dict(attrs).get("href")
            self.label = []

    def handle_data(self, data):
        if self.href is not None:
            self.label.append(data)

    def handle_endtag(self, tag):
        if tag.lower() == "a" and self.href is not None:
            self.links.append((self.href.strip(), " ".join(" ".join(self.label).split()).casefold()))
            self.href = None


def detect_shopee_verification_link(subject: str = "", body: str = "", raw_html: str = "") -> str:
    """Uu tien link xac thuc Shopee, ke ca dlink va SendGrid tracking cua mail Shopee."""
    # Giữ nguyên href, kể cả query token phân biệt hoa/thường và HTML lồng trong nút.
    parser = _ConfirmationAnchors()
    parser.feed(str(raw_html or body or ""))
    if _verification_context(subject, body, raw_html):
        for href, label in parser.links:
            if label in {"tại đây", "tai day", "xác nhận tại đây", "xác nhận", "click here", "verify", "confirm"}:
                parsed = urlsplit(href)
                if parsed.scheme.lower() in {"http", "https"} and parsed.hostname:
                    return href
    urls = _extract_email_urls(body, raw_html)
    if not urls:
        return ""

    def score(url: str) -> int:
        low = url.casefold()
        if any(x in low for x in ("/file/", ".png", ".jpg", ".jpeg", ".gif", "pixel", "unsubscribe")):
            return -100
        if "vn.shp.ee/dlink/" in low:
            return 100
        if "shp.ee/dlink/" in low:
            return 95
        if "sendgrid.net/ls/click" in low or ".ct.sendgrid.net/ls/click" in low:
            return 90 if _verification_context(subject, body, raw_html) else 20
        if "shopee" in low and any(x in low for x in ("verify", "verification", "confirm", "login", "auth", "security", "dlink")):
            return 85
        if "shopee" in low and _verification_context(subject, body, raw_html):
            return 55
        return 0

    ranked = sorted(((score(u), i, u) for i, u in enumerate(urls)), key=lambda x: (-x[0], x[1]))
    best = ranked[0] if ranked else (0, 0, "")
    return best[2] if best[0] >= 50 else ""


def detect_shopee_data(subject: str = "", body: str = "", raw_html: str = "") -> tuple[str, str, str]:
    text = "\n".join([str(subject or ""), str(body or ""), _html_to_text(raw_html or "")])
    otp = ""
    for pat in (
        r"(?i)(?:mã|ma|otp|verification\s*code|security\s*code)[^0-9]{0,30}(\d{6})",
        r"(?<!\d)(\d{6})(?!\d)",
    ):
        m = re.search(pat, text)
        if m:
            otp = m.group(1)
            break
    username = ""
    for pat in (
        r"(?i)(?:username|tên\s*đăng\s*nhập|ten\s*dang\s*nhap|tài\s*khoản|tai\s*khoan)\s*[:：-]?\s*([A-Za-z0-9._-]{3,64})",
        r"(?i)Shopee\s*(?:ID|user(?:name)?)\s*[:：-]?\s*([A-Za-z0-9._-]{3,64})",
    ):
        m = re.search(pat, text)
        if m:
            username = m.group(1).strip()
            break
    link = detect_shopee_verification_link(subject, body, raw_html)
    return username, otp, link


def random_human_username() -> str:
    first = ["anh","an","bao","binh","chau","chi","dung","duy","giang","ha","hai","han","hang","hieu","hoa","hoang","hung","huong","khanh","khoa","lam","lan","linh","long","mai","minh","nam","nga","ngan","ngoc","nhan","nhi","phong","phuc","phuong","quan","quang","son","tai","tam","thanh","thao","thu","thuy","tien","trang","tuan","tung","uyen","van","viet","vy"]
    last = ["bui","cao","dang","dao","dinh","do","duong","ho","hoang","huynh","lam","le","luu","ly","ngo","nguyen","pham","phan","truong","tran","vo","vu"]
    f, l = random.choice(first), random.choice(last)
    yy = f"{random.randint(0,99):02d}"
    year = str(random.randint(1987, 2005))
    return random.choice([f"{f}{l}", f"{f}{l}{yy}", f"{f}.{l}", f"{f}.{l}{yy}", f"{f}_{l}", f"{l}{f}", f"{f}{l}{year}"])


def random_password(length: int = 12) -> str:
    chars = string.ascii_letters + string.digits + "!@#$%"
    return "".join(secrets.choice(chars) for _ in range(max(8, length)))


def apply_suffix(local: str, enabled: bool, mode: str, length: int) -> str:
    local = re.sub(r"[^a-zA-Z0-9._-]", "", local).strip("._-").lower()
    if not enabled:
        return local
    sets = {"digits": string.digits, "letters": string.ascii_lowercase, "alnum": string.ascii_lowercase + string.digits}
    chars = sets.get(mode, sets["alnum"])
    suffix = "".join(secrets.choice(chars) for _ in range(max(1, min(12, int(length or 4)))))
    return local + suffix


class NativeMailStore:
    def __init__(self):
        self.path = _user_data_dir() / "mail_native_data.db"
        self._init_db()

    def conn(self):
        con = sqlite3.connect(self.path, timeout=30)
        con.row_factory = sqlite3.Row
        return con

    def _init_db(self):
        with self.conn() as con:
            con.executescript("""
            CREATE TABLE IF NOT EXISTS accounts(
              provider TEXT NOT NULL,
              address TEXT NOT NULL COLLATE NOCASE,
              password TEXT NOT NULL DEFAULT '',
              domain TEXT NOT NULL DEFAULT '',
              meta_json TEXT NOT NULL DEFAULT '{}',
              status TEXT NOT NULL DEFAULT 'SAVED',
              last_error TEXT NOT NULL DEFAULT '',
              mail_count INTEGER,
              last_check TEXT NOT NULL DEFAULT '',
              created_at TEXT NOT NULL,
              updated_at TEXT NOT NULL,
              PRIMARY KEY(provider,address)
            );
            CREATE TABLE IF NOT EXISTS messages(
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              provider TEXT NOT NULL,
              account_address TEXT NOT NULL,
              remote_id TEXT NOT NULL,
              sender TEXT NOT NULL DEFAULT '',
              recipient TEXT NOT NULL DEFAULT '',
              subject TEXT NOT NULL DEFAULT '',
              received_at TEXT NOT NULL DEFAULT '',
              body TEXT NOT NULL DEFAULT '',
              html_body TEXT NOT NULL DEFAULT '',
              shopee_username TEXT NOT NULL DEFAULT '',
              username_manual INTEGER NOT NULL DEFAULT 0,
              shopee_otp TEXT NOT NULL DEFAULT '',
              verify_link TEXT NOT NULL DEFAULT '',
              raw_json TEXT NOT NULL DEFAULT '{}',
              saved_at TEXT NOT NULL,
              updated_at TEXT NOT NULL,
              UNIQUE(provider,account_address,remote_id)
            );
            CREATE INDEX IF NOT EXISTS idx_native_messages_provider ON messages(provider);
            CREATE INDEX IF NOT EXISTS idx_native_messages_user ON messages(shopee_username);
            CREATE TABLE IF NOT EXISTS linked_emails(
              address TEXT NOT NULL COLLATE NOCASE PRIMARY KEY,
              source TEXT NOT NULL DEFAULT 'shopee_addmail',
              shopee_user TEXT NOT NULL DEFAULT '',
              provider_hint TEXT NOT NULL DEFAULT '',
              meta_json TEXT NOT NULL DEFAULT '{}',
              linked_at TEXT NOT NULL,
              updated_at TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_linked_emails_updated ON linked_emails(updated_at);
            """)

    def save_account(self, provider: str, address: str, password: str = "", domain: str = "", meta: dict | None = None,
                     status: str = "SAVED", error: str = "", mail_count: int | None = None, checked: bool = False):
        now = _now_iso()
        with self.conn() as con:
            old = con.execute("SELECT created_at FROM accounts WHERE provider=? AND address=?", (provider,address)).fetchone()
            created = old[0] if old else now
            con.execute("""
            INSERT INTO accounts(provider,address,password,domain,meta_json,status,last_error,mail_count,last_check,created_at,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(provider,address) DO UPDATE SET
              password=excluded.password,domain=excluded.domain,meta_json=excluded.meta_json,status=excluded.status,
              last_error=excluded.last_error,mail_count=excluded.mail_count,
              last_check=CASE WHEN excluded.last_check<>'' THEN excluded.last_check ELSE accounts.last_check END,
              updated_at=excluded.updated_at
            """, (provider,address,password,domain,json.dumps(meta or {}, ensure_ascii=False),status,error,mail_count,now if checked else "",created,now))

    def list_accounts(self, provider: str, search: str = "") -> list[dict]:
        q = f"%{search.strip().lower()}%"
        with self.conn() as con:
            rows = con.execute("""
              SELECT a.*,
                COALESCE((SELECT group_concat(DISTINCT m.shopee_username) FROM messages m
                          WHERE m.provider=a.provider AND lower(m.account_address)=lower(a.address) AND m.shopee_username<>''),'') AS shopee_usernames
              FROM accounts a
              WHERE a.provider=? AND (
                ?='' OR lower(a.address) LIKE ? OR EXISTS (
                  SELECT 1 FROM messages mx WHERE mx.provider=a.provider
                    AND lower(mx.account_address)=lower(a.address) AND lower(mx.shopee_username) LIKE ?
                )
              )
              ORDER BY a.updated_at DESC
            """, (provider,search.strip(),q,q)).fetchall()
        return [dict(r) for r in rows]

    def delete_accounts(self, provider: str, addresses: list[str]):
        addresses = [x for x in addresses if x]
        if not addresses: return 0
        marks = ",".join("?" for _ in addresses)
        with self.conn() as con:
            cur = con.execute(f"DELETE FROM accounts WHERE provider=? AND address IN ({marks})", [provider,*addresses])
            return cur.rowcount

    def save_message(self, provider: str, address: str, remote_id: str, subject: str, sender: str, received_at: str,
                     body: str, html_body: str = "", recipient: str = "", raw: dict | None = None, manual: bool = False):
        username, otp, link = detect_shopee_data(subject, body, html_body)
        now = _now_iso()
        with self.conn() as con:
            old = con.execute("SELECT * FROM messages WHERE provider=? AND account_address=? AND remote_id=?", (provider,address,remote_id)).fetchone()
            if old and int(old["username_manual"] or 0):
                username = old["shopee_username"]
                manual = True
            con.execute("""
            INSERT INTO messages(provider,account_address,remote_id,sender,recipient,subject,received_at,body,html_body,
              shopee_username,username_manual,shopee_otp,verify_link,raw_json,saved_at,updated_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(provider,account_address,remote_id) DO UPDATE SET sender=excluded.sender,recipient=excluded.recipient,
              subject=excluded.subject,received_at=excluded.received_at,body=excluded.body,html_body=excluded.html_body,
              shopee_username=CASE WHEN messages.username_manual=1 THEN messages.shopee_username ELSE excluded.shopee_username END,
              username_manual=CASE WHEN messages.username_manual=1 THEN 1 ELSE excluded.username_manual END,
              shopee_otp=excluded.shopee_otp,verify_link=excluded.verify_link,raw_json=excluded.raw_json,updated_at=excluded.updated_at
            """, (provider,address,str(remote_id),sender,recipient,subject,received_at,body,html_body,username,1 if manual else 0,otp,link,
                  json.dumps(raw or {},ensure_ascii=False), old["saved_at"] if old else now, now))
            row = con.execute("SELECT * FROM messages WHERE provider=? AND account_address=? AND remote_id=?",(provider,address,str(remote_id))).fetchone()
        return dict(row)

    def list_messages(self, provider: str, search: str = "") -> list[dict]:
        q=f"%{search.strip().lower()}%"
        with self.conn() as con:
            rows=con.execute("""
              SELECT * FROM messages WHERE provider=? AND (?='' OR lower(account_address) LIKE ? OR lower(subject) LIKE ?
                OR lower(shopee_username) LIKE ? OR lower(shopee_otp) LIKE ?)
              ORDER BY COALESCE(NULLIF(received_at,''),saved_at) DESC,id DESC
            """,(provider,search.strip(),q,q,q,q)).fetchall()
        return [dict(r) for r in rows]

    def update_username(self, mid: int, username: str):
        with self.conn() as con:
            con.execute("UPDATE messages SET shopee_username=?,username_manual=1,updated_at=? WHERE id=?",(username.strip(),_now_iso(),mid))

    def rescan(self, provider: str) -> int:
        changed=0
        with self.conn() as con:
            rows=con.execute("SELECT * FROM messages WHERE provider=?",(provider,)).fetchall()
            for r in rows:
                if int(r["username_manual"] or 0):
                    continue
                u,o,l=detect_shopee_data(r["subject"],r["body"],r["html_body"])
                if u!=r["shopee_username"] or o!=r["shopee_otp"] or l!=r["verify_link"]:
                    con.execute("UPDATE messages SET shopee_username=?,shopee_otp=?,verify_link=?,updated_at=? WHERE id=?",(u,o,l,_now_iso(),r["id"]))
                    changed+=1
        return changed

    def delete_messages(self, provider: str, ids: list[int]):
        ids=[int(x) for x in ids if str(x).isdigit()]
        if not ids:return 0
        marks=",".join("?" for _ in ids)
        with self.conn() as con:
            cur=con.execute(f"DELETE FROM messages WHERE provider=? AND id IN ({marks})",[provider,*ids])
            return cur.rowcount

    def save_linked_email(self, address: str, source: str = "shopee_addmail", shopee_user: str = "",
                          provider_hint: str = "", meta: dict | None = None) -> dict:
        """Lưu email vừa liên kết Shopee mà KHÔNG ghi đè credential mailbox đã có."""
        address = str(address or "").strip().lower()
        if not address or "@" not in address:
            raise ValueError("Email liên kết không hợp lệ")
        now = _now_iso()
        with self.conn() as con:
            old = con.execute("SELECT linked_at FROM linked_emails WHERE address=?", (address,)).fetchone()
            linked_at = old[0] if old else now
            con.execute("""
              INSERT INTO linked_emails(address,source,shopee_user,provider_hint,meta_json,linked_at,updated_at)
              VALUES(?,?,?,?,?,?,?)
              ON CONFLICT(address) DO UPDATE SET
                source=excluded.source,
                shopee_user=CASE WHEN excluded.shopee_user<>'' THEN excluded.shopee_user ELSE linked_emails.shopee_user END,
                provider_hint=CASE WHEN excluded.provider_hint<>'' THEN excluded.provider_hint ELSE linked_emails.provider_hint END,
                meta_json=excluded.meta_json,updated_at=excluded.updated_at
            """, (address, source, shopee_user, provider_hint, json.dumps(meta or {}, ensure_ascii=False), linked_at, now))
            row = con.execute("SELECT * FROM linked_emails WHERE address=?", (address,)).fetchone()
        return dict(row)

    def list_linked_emails(self, search: str = "") -> list[dict]:
        q = f"%{str(search or '').strip().lower()}%"
        with self.conn() as con:
            rows = con.execute("""
              SELECT * FROM linked_emails
              WHERE ?='' OR lower(address) LIKE ? OR lower(shopee_user) LIKE ?
              ORDER BY updated_at DESC
            """, (str(search or '').strip(), q, q)).fetchall()
        return [dict(r) for r in rows]

    def get_linked_email(self, address: str) -> dict | None:
        with self.conn() as con:
            row = con.execute("SELECT * FROM linked_emails WHERE lower(address)=lower(?)", (str(address or '').strip(),)).fetchone()
        return dict(row) if row else None

    def find_mail_account(self, address: str) -> dict | None:
        """Tìm credential mailbox đã lưu cho email trong Mail.tm/TinyHost/Outlook."""
        address = str(address or "").strip().lower()
        if not address:
            return None
        for provider in ("mailtm", "tinyhost"):
            rows = self.list_accounts(provider, address)
            for row in rows:
                if str(row.get("address") or "").strip().lower() == address:
                    return {"provider": provider, "account": row}
        if outlook_backend is not None:
            try:
                row = outlook_backend.get_account(address)
                if row:
                    return {"provider": "outlook", "account": row}
            except Exception:
                pass
        return None

    def export_backup(self, provider: str) -> dict:
        return {"version":1,"provider":provider,"exported_at":_now_iso(),"accounts":self.list_accounts(provider),"messages":self.list_messages(provider)}

    def import_backup(self, provider: str, data: dict) -> tuple[int,int]:
        ac=mc=0
        for a in data.get("accounts") or []:
            try:
                self.save_account(provider,a.get("address",''),a.get("password",''),a.get("domain",''),json.loads(a.get("meta_json") or '{}'),a.get("status",'IMPORTED'),a.get("last_error",''),a.get("mail_count"),bool(a.get("last_check")))
                ac+=1
            except Exception: pass
        for m in data.get("messages") or []:
            try:
                row=self.save_message(provider,m.get("account_address",''),m.get("remote_id",secrets.token_hex(8)),m.get("subject",''),m.get("sender",''),m.get("received_at",''),m.get("body",''),m.get("html_body",''),m.get("recipient",''),json.loads(m.get("raw_json") or '{}'))
                if m.get("shopee_username") and m.get("username_manual"):
                    self.update_username(row["id"],m["shopee_username"])
                mc+=1
            except Exception: pass
        return ac,mc


def register_linked_email_global(address: str, shopee_user: str = "", source: str = "shopee_addmail",
                                 provider_hint: str = "", meta: dict | None = None) -> dict | None:
    """Cho trang AddMail lưu email ngay cả khi trang Đọc Mail chưa được mở."""
    address = str(address or "").strip().lower()
    if not address or "@" not in address:
        return None
    store = NativeMailStore()
    return store.save_linked_email(address, source=source, shopee_user=shopee_user, provider_hint=provider_hint, meta=meta)


class BaseMailTab(tk.Frame):
    provider = "base"
    def __init__(self, parent, store: NativeMailStore):
        super().__init__(parent, bg="#f6f7f9")
        self.store=store
        self.stop_event=threading.Event()
        self.log_var=None
        self.status_var=tk.StringVar(value="Sẵn sàng.")
        self.auto_refresh_seconds = tk.StringVar(value="0")
        self.auto_refresh_status = tk.StringVar(value="Tắt")
        self._refresh_remaining = 0
        self._active_workers = 0
        bar = tk.Frame(self, bg="#f6f7f9")
        bar.pack(fill="x", padx=12, pady=4)
        tk.Label(bar, text="Tự làm mới inbox (giây, 0 = tắt):", bg="#f6f7f9").pack(side="left")
        ttk.Spinbox(bar, from_=0, to=86400, width=8, textvariable=self.auto_refresh_seconds).pack(side="left", padx=6)
        tk.Label(bar, textvariable=self.auto_refresh_status, bg="#f6f7f9").pack(side="left")
        self.auto_refresh_seconds.trace_add("write", self._reset_auto_refresh)
        self.after(1000, self._auto_refresh_tick)

    def _reset_auto_refresh(self, *_):
        try:
            seconds = int(self.auto_refresh_seconds.get())
            if not 0 <= seconds <= 86400:
                raise ValueError
        except ValueError:
            self._refresh_remaining = 0
            self.auto_refresh_status.set("Nhập 0–86400 giây")
            return
        self._refresh_remaining = seconds
        self.auto_refresh_status.set(f"Còn {seconds} giây" if seconds else "Tắt")

    def _auto_refresh_tick(self):
        if self.stop_event.is_set():
            return
        try:
            seconds = int(self.auto_refresh_seconds.get())
        except ValueError:
            seconds = 0
        if 0 < seconds <= 86400:
            ready = bool(getattr(self, "session_acct", None)) if self.provider == "outlook" else bool(getattr(self, "address", ""))
            if not ready:
                self.auto_refresh_status.set("Chờ mở hộp thư")
                self._refresh_remaining = seconds
            elif self._active_workers:
                self.auto_refresh_status.set("Đang xử lý…")
            else:
                self._refresh_remaining -= 1
                if self._refresh_remaining <= 0:
                    self._refresh_remaining = seconds
                    self.worker(self.refresh_current if self.provider == "outlook" else self.refresh_inbox)
                self.auto_refresh_status.set(f"Còn {self._refresh_remaining} giây")
        self.after(1000, self._auto_refresh_tick)

    def ui(self, func, *args):
        try:self.after(0,func,*args)
        except Exception:pass

    def open_email_link(self, value):
        link = str(value or "").strip()
        try:
            parsed = urlsplit(link)
            valid = parsed.scheme.lower() in {"http", "https"} and bool(parsed.hostname)
        except ValueError:
            valid = False
        if not valid:
            messagebox.showinfo("Chưa có link", "Chọn hoặc đọc email có link xác thực HTTP/HTTPS trước.", parent=self)
            return
        try:
            if not webbrowser.open(link, new=2):
                self.status("Không mở được trình duyệt. Hãy dùng Copy link.")
        except Exception as exc:
            self.status(f"Không mở được link: {exc}")

    def worker(self, func, *args):
        self._active_workers += 1
        def run():
            try:
                func(*args)
            finally:
                self.ui(self._worker_finished)
        threading.Thread(target=run,daemon=True).start()

    def _worker_finished(self):
        self._active_workers = max(0, self._active_workers - 1)

    def log(self, msg: str):
        if not hasattr(self,"log_text"):return
        stamp=datetime.now().strftime("%H:%M:%S")
        def add():
            try:
                self.log_text.configure(state="normal")
                self.log_text.insert("end",f"[{stamp}] {msg}\n")
                self.log_text.see("end")
                self.log_text.configure(state="disabled")
            except Exception: pass
        self.ui(add)

    def status(self,msg):
        text=str(msg)
        self.ui(self.status_var.set,text)
        low=text.lower()
        if any(k in low for k in ("lỗi", "không hoạt động", "offline", "không hợp lệ", "không kiểm tra được", "thất bại", "error", "failed")):
            self.log("⚠ " + text)

    def build_log_panel(self):
        panel=tk.LabelFrame(self,text="NHẬT KÝ / LOG",bg="#f6f7f9",fg="#333",bd=1,relief="solid")
        # Đặt LOG trước notebook trong thứ tự pack để LOG luôn được giữ chiều cao,
        # tránh bị notebook expand ép xuống còn 1px.
        if hasattr(self,"nb"):
            try:self.nb.pack_forget()
            except Exception:pass
        panel.pack(side="bottom",fill="x",padx=12,pady=(0,10))
        if hasattr(self,"nb"):
            self.nb.pack(fill="both",expand=True,padx=12,pady=6)
        row=tk.Frame(panel,bg="#f6f7f9")
        row.pack(fill="x",padx=8,pady=(5,2))
        tk.Label(row,textvariable=self.status_var,bg="#f6f7f9",fg="#666",font=("Segoe UI",9,"bold")).pack(side="left",fill="x",expand=True)
        ttk.Button(row,text="📋 Copy log",command=self.copy_log,style="Soft.TButton").pack(side="right",padx=2)
        ttk.Button(row,text="🧹 Xóa log",command=self.clear_log,style="Soft.TButton").pack(side="right",padx=2)
        self.log_text=scrolledtext.ScrolledText(panel,height=7,state="disabled",font=("Consolas",8),wrap="word")
        self.log_text.pack(fill="x",expand=False,padx=8,pady=(2,8))

    def clear_log(self):
        if not hasattr(self,"log_text"):return
        self.log_text.configure(state="normal");self.log_text.delete("1.0","end");self.log_text.configure(state="disabled")
        self.status("Đã xóa log.")

    def copy_log(self):
        if not hasattr(self,"log_text"):return
        text=self.log_text.get("1.0","end-1c")
        if text:self.copy(text)

    @staticmethod
    def _status_is_active(value):
        s=str(value or "").strip().upper()
        if any(x in s for x in ("KHÔNG", "OFFLINE", "LỖI", "ERROR", "FAILED", "INVALID", "CHƯA XÁC ĐỊNH", "KHÔNG HỢP LỆ")):
            return False
        return any(x in s for x in ("HOẠT ĐỘNG", "THÀNH CÔNG", "ĐĂNG NHẬP OK", "ONLINE", " OK", "OK"))

    def _account_filter_ok(self,row):
        mode=self.account_filter_var.get() if hasattr(self,"account_filter_var") else "Tất cả"
        raw=row.get("status", row.get("last_status", ""))
        active=self._status_is_active(raw)
        if mode=="Hoạt động":return active
        if mode=="Mail lỗi / không hoạt động":return not active
        return True

    def configure_account_tree(self):
        if not hasattr(self,"account_tree"):return
        self._checked_accounts=getattr(self,"_checked_accounts",set())
        self.account_tree.tag_configure("active",foreground="#16a34a")
        self.account_tree.tag_configure("inactive",foreground="#d4a017")
        self.account_tree.bind("<Button-1>",self.toggle_account_check,add="+")

    def toggle_account_check(self,event):
        if not hasattr(self,"account_tree"):return
        row=self.account_tree.identify_row(event.y);col=self.account_tree.identify_column(event.x)
        if not row or col!="#1":return
        if row in self._checked_accounts:self._checked_accounts.remove(row)
        else:self._checked_accounts.add(row)
        vals=list(self.account_tree.item(row,"values"))
        if vals:
            vals[0]="☑" if row in self._checked_accounts else "☐"
            self.account_tree.item(row,values=vals)
        return "break"

    def check_all_visible_accounts(self):
        if not hasattr(self,"account_tree"):return
        for iid in self.account_tree.get_children():self._checked_accounts.add(iid)
        self.load_accounts();self.status(f"Đã tích {len(self._checked_accounts)} mail/tài khoản.")

    def clear_account_checks(self):
        self._checked_accounts.clear();self.load_accounts();self.status("Đã bỏ toàn bộ tích chọn.")

    def delete_checked_accounts(self):
        ids=[x for x in self._checked_accounts if x]
        if not ids:return self.status("Chưa tích chọn mail/tài khoản để xóa.")
        if not messagebox.askyesno("Xóa mail/tài khoản",f"Xóa {len(ids)} mail/tài khoản đã tích? Dữ liệu đã lưu của các tài khoản này sẽ không tự xóa."):
            return
        if self.provider=="outlook":
            if outlook_backend is None:return self.status("Lỗi: thiếu backend Outlook.")
            n=outlook_backend.delete_accounts(ids)
        else:
            n=self.store.delete_accounts(self.provider,ids)
        self._checked_accounts.difference_update(ids);self.load_accounts();self.status(f"Đã xóa {n} mail/tài khoản đã tích.")

    def copy(self,value):
        if not value:return
        try:
            self.clipboard_clear();self.clipboard_append(str(value));self.status("Đã copy vào clipboard.")
        except Exception:pass

    def export_csv(self, filename: str, headers: list[str], rows: list[list[Any]]):
        path=filedialog.asksaveasfilename(title="Xuất CSV",defaultextension=".csv",initialfile=filename,filetypes=[("CSV","*.csv"),("Tất cả","*.*")])
        if not path:return
        with open(path,"w",encoding="utf-8-sig",newline="") as f:
            w=csv.writer(f);w.writerow(headers);w.writerows(rows)
        self.status(f"Đã xuất {len(rows)} dòng: {os.path.basename(path)}")

    def backup_json(self):
        path=filedialog.asksaveasfilename(title="Backup JSON",defaultextension=".json",initialfile=f"{self.provider}_backup_{datetime.now():%Y%m%d}.json",filetypes=[("JSON","*.json")])
        if not path:return
        Path(path).write_text(json.dumps(self.store.export_backup(self.provider),ensure_ascii=False,indent=2),encoding="utf-8")
        self.status("Đã backup dữ liệu JSON.")

    def import_json(self):
        path=filedialog.askopenfilename(title="Import JSON",filetypes=[("JSON","*.json"),("Tất cả","*.*")])
        if not path:return
        try:data=json.loads(Path(path).read_text(encoding="utf-8-sig"))
        except Exception as e:return messagebox.showerror("Import lỗi",str(e))
        ac,mc=self.store.import_backup(self.provider,data)
        self.status(f"Import xong: {ac} tài khoản/địa chỉ, {mc} mail.")
        if hasattr(self,"load_saved"):self.load_saved()
        if hasattr(self,"load_accounts"):self.load_accounts()

    def stop_all(self): self.stop_event.set()


class MailTMTab(BaseMailTab):
    provider="mailtm"
    def __init__(self,parent,store):
        super().__init__(parent,store)
        self.token="";self.address="";self.password="";self.account_id="";self.domains=[];self.inbox=[];self.current_msg=None
        self.bulk_stop=threading.Event();self.check_stop=threading.Event()
        self._build()
        self.worker(self.refresh_domains)
        self.load_saved();self.load_accounts()

    def request(self,method,path,data=None,token=None):
        if httpx is None: raise RuntimeError("Thiếu thư viện httpx")
        headers={"Accept":"application/ld+json, application/json"}
        if token:headers["Authorization"]="Bearer "+token
        with httpx.Client(timeout=25,headers=headers,follow_redirects=True) as c:
            r=c.request(method,MAILTM_API+path,json=data if data is not None else None)
        try:j=r.json()
        except Exception:j={"message":r.text[:500]}
        if r.status_code>=400:
            detail=j.get("hydra:description") if isinstance(j,dict) else ""
            if not detail and isinstance(j,dict):detail=j.get("message") or j.get("detail")
            raise RuntimeError(f"HTTP {r.status_code}: {detail or r.reason_phrase}")
        return j

    def _build(self):
        top=tk.Frame(self,bg="#f6f7f9");top.pack(fill="x",padx=12,pady=(12,6))
        tk.Label(top,text="Mail.tm • Tạo / đăng nhập / đọc mail trực tiếp trong tool",bg="#f6f7f9",fg="#222",font=("Segoe UI",13,"bold")).pack(side="left")
        ttk.Button(top,text="💾 Backup JSON",command=self.backup_json,style="Soft.TButton").pack(side="right",padx=3)
        ttk.Button(top,text="📥 Import JSON",command=self.import_json,style="Soft.TButton").pack(side="right",padx=3)

        login=tk.Frame(self,bg="#fff",highlightthickness=1,highlightbackground="#e4e7eb");login.pack(fill="x",padx=12,pady=6)
        for i,w in enumerate((2,2,2,1,1,1)):login.grid_columnconfigure(i,weight=w)
        self.user_var=tk.StringVar();self.domain_var=tk.StringVar();self.pass_var=tk.StringVar();self.suffix_var=tk.BooleanVar();self.suffix_mode=tk.StringVar(value="alnum");self.suffix_len=tk.IntVar(value=4)
        tk.Label(login,text="Tên mail",bg="#fff",fg="#333").grid(row=0,column=0,sticky="w",padx=10,pady=(8,2));tk.Label(login,text="Domain",bg="#fff",fg="#333").grid(row=0,column=1,sticky="w",padx=6,pady=(8,2));tk.Label(login,text="Mật khẩu",bg="#fff",fg="#333").grid(row=0,column=2,sticky="w",padx=6,pady=(8,2))
        ttk.Entry(login,textvariable=self.user_var).grid(row=1,column=0,sticky="ew",padx=(10,5),pady=(0,8))
        self.domain_combo=ttk.Combobox(login,textvariable=self.domain_var,state="readonly");self.domain_combo.grid(row=1,column=1,sticky="ew",padx=5,pady=(0,8))
        ttk.Entry(login,textvariable=self.pass_var,show="•").grid(row=1,column=2,sticky="ew",padx=5,pady=(0,8))
        ttk.Button(login,text="↻ Domain",command=lambda:self.worker(self.refresh_domains),style="Soft.TButton").grid(row=1,column=3,sticky="ew",padx=4,pady=(0,8))
        ttk.Button(login,text="🎲 Random",command=self.random_fill,style="Soft.TButton").grid(row=1,column=4,sticky="ew",padx=4,pady=(0,8))
        ttk.Button(login,text="➕ Tạo",command=lambda:self.worker(self.create_account),style="Accent.TButton").grid(row=1,column=5,sticky="ew",padx=(4,10),pady=(0,8))
        ttk.Checkbutton(login,text="Random cuối tên",variable=self.suffix_var).grid(row=2,column=0,sticky="w",padx=10,pady=(0,8))
        ttk.Combobox(login,textvariable=self.suffix_mode,values=["alnum","digits","letters"],state="readonly",width=10).grid(row=2,column=1,sticky="w",padx=5,pady=(0,8))
        ttk.Spinbox(login,from_=1,to=12,textvariable=self.suffix_len,width=6).grid(row=2,column=2,sticky="w",padx=5,pady=(0,8))
        ttk.Button(login,text="🔑 Đăng nhập",command=lambda:self.worker(self.login_account),style="Soft.TButton").grid(row=2,column=5,sticky="ew",padx=(4,10),pady=(0,8))

        current=tk.Frame(self,bg="#fff");current.pack(fill="x",padx=12,pady=(0,6))
        self.current_var=tk.StringVar(value="Chưa đăng nhập")
        tk.Label(current,textvariable=self.current_var,bg="#fff",fg="#ee4d2d",font=("Segoe UI",9,"bold")).pack(side="left",padx=10,pady=7)
        ttk.Button(current,text="Copy email",command=lambda:self.copy(self.address),style="Soft.TButton").pack(side="right",padx=4,pady=4)
        ttk.Button(current,text="Copy password",command=lambda:self.copy(self.password),style="Soft.TButton").pack(side="right",padx=4,pady=4)
        ttk.Button(current,text="↻ Inbox",command=lambda:self.worker(self.refresh_inbox),style="Soft.TButton").pack(side="right",padx=4,pady=4)

        self.nb=ttk.Notebook(self);self.nb.pack(fill="both",expand=True,padx=12,pady=6)
        self.inbox_tab=tk.Frame(self.nb,bg="#f6f7f9");self.saved_tab=tk.Frame(self.nb,bg="#f6f7f9");self.bulk_tab=tk.Frame(self.nb,bg="#f6f7f9")
        self.nb.add(self.inbox_tab,text="📥 Inbox / Đọc mail");self.nb.add(self.saved_tab,text="💾 Mail đã lưu / Tìm Shopee");self.nb.add(self.bulk_tab,text="⚡ Tạo hàng loạt / DS tài khoản")
        self._build_inbox();self._build_saved();self._build_bulk()
        self.build_log_panel()

    def _build_inbox(self):
        self.inbox_tab.grid_columnconfigure(0,weight=1);self.inbox_tab.grid_columnconfigure(1,weight=1);self.inbox_tab.grid_rowconfigure(0,weight=1)
        left=tk.Frame(self.inbox_tab,bg="#fff");left.grid(row=0,column=0,sticky="nsew",padx=(0,5),pady=5);left.grid_rowconfigure(0,weight=1);left.grid_columnconfigure(0,weight=1)
        self.inbox_tree=ttk.Treeview(left,columns=("from","subject","date"),show="headings");
        for c,t,w in (("from","Người gửi",180),("subject","Tiêu đề",300),("date","Thời gian",155)):self.inbox_tree.heading(c,text=t);self.inbox_tree.column(c,width=w)
        self.inbox_tree.grid(row=0,column=0,sticky="nsew");ttk.Scrollbar(left,orient="vertical",command=self.inbox_tree.yview).grid(row=0,column=1,sticky="ns");self.inbox_tree.configure(yscrollcommand=lambda a,b:None);self.inbox_tree.bind("<<TreeviewSelect>>",self.open_message)
        right=tk.Frame(self.inbox_tab,bg="#fff");right.grid(row=0,column=1,sticky="nsew",padx=(5,0),pady=5);right.grid_columnconfigure(0,weight=1);right.grid_rowconfigure(4,weight=1)
        self.subject_var=tk.StringVar(value="Chọn một mail để đọc");tk.Label(right,textvariable=self.subject_var,bg="#fff",fg="#222",font=("Segoe UI",10,"bold"),wraplength=520,justify="left").grid(row=0,column=0,sticky="ew",padx=10,pady=(10,4))
        meta=tk.Frame(right,bg="#fff");meta.grid(row=1,column=0,sticky="ew",padx=10);self.otp_var=tk.StringVar();self.username_var=tk.StringVar();self.link_var=tk.StringVar()
        ttk.Entry(meta,textvariable=self.otp_var,width=12).pack(side="left");ttk.Button(meta,text="📋 OTP",command=lambda:self.copy(self.otp_var.get()),style="Soft.TButton").pack(side="left",padx=3);ttk.Entry(meta,textvariable=self.username_var,width=22).pack(side="left",padx=(8,2));ttk.Button(meta,text="💾 Username",command=self.save_current_username,style="Soft.TButton").pack(side="left")
        linkrow=tk.Frame(right,bg="#fff");linkrow.grid(row=2,column=0,sticky="ew",padx=10,pady=5);ttk.Entry(linkrow,textvariable=self.link_var).pack(side="left",fill="x",expand=True);ttk.Button(linkrow,text="Copy link",command=lambda:self.copy(self.link_var.get()),style="Soft.TButton").pack(side="left",padx=4)
        ttk.Button(linkrow, text="Mở link", command=lambda: self.open_email_link(self.link_var.get()), style="Soft.TButton").pack(side="left", padx=4)
        btn=tk.Frame(right,bg="#fff");btn.grid(row=3,column=0,sticky="ew",padx=10,pady=3);ttk.Button(btn,text="📩 Đọc lại",command=lambda:self.worker(self.refresh_inbox),style="Soft.TButton").pack(side="left");ttk.Button(btn,text="🔐 Tìm link xác thực",command=lambda:self.worker(self.find_latest_verify_mail),style="Soft.TButton").pack(side="left",padx=4);ttk.Button(btn,text="💾 Lưu mail",command=self.save_current_mail,style="Accent.TButton").pack(side="left",padx=4);ttk.Button(btn,text="📋 Copy nội dung",command=lambda:self.copy(self.body_text.get("1.0","end-1c")),style="Soft.TButton").pack(side="left")
        self.body_text=scrolledtext.ScrolledText(right,wrap="word",font=("Segoe UI",9));self.body_text.grid(row=4,column=0,sticky="nsew",padx=10,pady=(5,10))

    def _build_saved(self):
        tab=self.saved_tab;tab.grid_columnconfigure(0,weight=1);tab.grid_rowconfigure(1,weight=1)
        bar=tk.Frame(tab,bg="#fff");bar.grid(row=0,column=0,sticky="ew",pady=5);self.saved_search=tk.StringVar();self.verify_only_var=tk.BooleanVar(value=False);ttk.Entry(bar,textvariable=self.saved_search).pack(side="left",fill="x",expand=True,padx=(8,4),pady=6);ttk.Button(bar,text="🔎 Tìm",command=self.load_saved,style="Soft.TButton").pack(side="left",padx=2);ttk.Checkbutton(bar,text="🔐 Chỉ link xác thực",variable=self.verify_only_var,command=self.load_saved).pack(side="left",padx=3);ttk.Button(bar,text="📋 Copy link",command=self.copy_selected_saved_link,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(bar,text="Hiện tất cả",command=lambda:(self.saved_search.set(""),self.verify_only_var.set(False),self.load_saved()),style="Soft.TButton").pack(side="left",padx=2);ttk.Button(bar,text="🔄 Quét lại",command=self.rescan_saved,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(bar,text="📤 CSV",command=self.export_saved,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(bar,text="🗑 Xóa",command=self.delete_saved,style="Danger.TButton").pack(side="left",padx=(2,8))
        box=tk.Frame(tab,bg="#fff");box.grid(row=1,column=0,sticky="nsew");box.grid_rowconfigure(0,weight=1);box.grid_columnconfigure(0,weight=1)
        self.saved_tree=ttk.Treeview(box,columns=("email","user","otp","subject","date"),show="headings",selectmode="extended")
        for c,t,w in (("email","Email",220),("user","Username Shopee",160),("otp","OTP",80),("subject","Tiêu đề",300),("date","Thời gian",150)):self.saved_tree.heading(c,text=t);self.saved_tree.column(c,width=w)
        self.saved_tree.grid(row=0,column=0,sticky="nsew");ttk.Scrollbar(box,orient="vertical",command=self.saved_tree.yview).grid(row=0,column=1,sticky="ns");self.saved_tree.configure(yscrollcommand=lambda a,b:None);self.saved_tree.bind("<Double-1>",self.open_saved)

    def _build_bulk(self):
        tab=self.bulk_tab;tab.grid_columnconfigure(0,weight=1);tab.grid_columnconfigure(1,weight=1);tab.grid_rowconfigure(1,weight=1)
        bulk=tk.LabelFrame(tab,text="Tạo hàng loạt",bg="#fff",fg="#222");bulk.grid(row=0,column=0,sticky="ew",padx=(0,5),pady=5);self.bulk_count=tk.IntVar(value=10);self.bulk_prefix=tk.StringVar();self.bulk_domain=tk.StringVar();self.bulk_password=tk.StringVar();self.bulk_delay=tk.DoubleVar(value=0.7)
        for i in range(5):bulk.grid_columnconfigure(i,weight=1)
        ttk.Spinbox(bulk,from_=1,to=200,textvariable=self.bulk_count,width=7).grid(row=0,column=0,padx=4,pady=7);ttk.Entry(bulk,textvariable=self.bulk_prefix).grid(row=0,column=1,sticky="ew",padx=4);self.bulk_domain_combo=ttk.Combobox(bulk,textvariable=self.bulk_domain,state="readonly");self.bulk_domain_combo.grid(row=0,column=2,sticky="ew",padx=4);ttk.Entry(bulk,textvariable=self.bulk_password).grid(row=0,column=3,sticky="ew",padx=4);ttk.Spinbox(bulk,from_=0.2,to=10,increment=.1,textvariable=self.bulk_delay,width=7).grid(row=0,column=4,padx=4)
        ttk.Button(bulk,text="▶ Bắt đầu tạo",command=lambda:self.worker(self.bulk_create),style="Accent.TButton").grid(row=1,column=0,columnspan=4,sticky="ew",padx=4,pady=(0,7));ttk.Button(bulk,text="■ Dừng",command=self.bulk_stop.set,style="Danger.TButton").grid(row=1,column=4,sticky="ew",padx=4,pady=(0,7))
        checker=tk.LabelFrame(tab,text="Check tài khoản",bg="#fff",fg="#222");checker.grid(row=0,column=1,sticky="ew",padx=(5,0),pady=5);checker.grid_columnconfigure(0,weight=1)
        self.check_input=scrolledtext.ScrolledText(checker,height=4,font=("Consolas",8));self.check_input.grid(row=0,column=0,columnspan=4,sticky="ew",padx=5,pady=5);ttk.Button(checker,text="📥 Nạp DS đã tạo",command=self.load_created_to_check,style="Soft.TButton").grid(row=1,column=0,sticky="ew",padx=3,pady=4);ttk.Button(checker,text="✓ Check",command=lambda:self.worker(self.check_accounts),style="Accent.TButton").grid(row=1,column=1,sticky="ew",padx=3,pady=4);ttk.Button(checker,text="■ Dừng",command=self.check_stop.set,style="Danger.TButton").grid(row=1,column=2,sticky="ew",padx=3,pady=4);ttk.Button(checker,text="📤 CSV",command=self.export_checks,style="Soft.TButton").grid(row=1,column=3,sticky="ew",padx=3,pady=4)
        acc=tk.Frame(tab,bg="#fff");acc.grid(row=1,column=0,columnspan=2,sticky="nsew",pady=5);acc.grid_rowconfigure(1,weight=1);acc.grid_columnconfigure(0,weight=1)
        abar=tk.Frame(acc,bg="#fff");abar.grid(row=0,column=0,sticky="ew");self.account_search=tk.StringVar();self.account_filter_var=tk.StringVar(value="Tất cả");ttk.Entry(abar,textvariable=self.account_search).pack(side="left",fill="x",expand=True,padx=6,pady=5);ttk.Button(abar,text="🔎 Tìm",command=self.load_accounts,style="Soft.TButton").pack(side="left",padx=2);ttk.Combobox(abar,textvariable=self.account_filter_var,values=["Tất cả","Hoạt động","Mail lỗi / không hoạt động"],state="readonly",width=22).pack(side="left",padx=2);ttk.Button(abar,text="Lọc",command=self.load_accounts,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(abar,text="☑ Chọn tất cả",command=self.check_all_visible_accounts,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(abar,text="☐ Bỏ chọn",command=self.clear_account_checks,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(abar,text="🗑 Xóa tích",command=self.delete_checked_accounts,style="Danger.TButton").pack(side="left",padx=2);ttk.Button(abar,text="📤 CSV",command=self.export_accounts,style="Soft.TButton").pack(side="left",padx=2)
        self.account_tree=ttk.Treeview(acc,columns=("pick","email","user","status","error","count","checked","created"),show="headings",selectmode="extended");
        for c,t,w in (("pick","✓",42),("email","Email",230),("user","Username Shopee",160),("status","Trạng thái",135),("error","Lỗi / ghi chú",250),("count","Số mail",70),("checked","Check lúc",145),("created","Tạo lúc",145)):self.account_tree.heading(c,text=t);self.account_tree.column(c,width=w,anchor="center" if c=="pick" else "w")
        self.account_tree.grid(row=1,column=0,sticky="nsew");self.account_tree.bind("<Double-1>",self.use_account);self.configure_account_tree()
        self.check_results=[]

    def random_fill(self):
        self.user_var.set(random_human_username());self.pass_var.set(random_password())

    def refresh_domains(self):
        try:
            j=self.request("GET","/domains?page=1");rows=j.get("hydra:member") or j.get("member") or []
            domains=[d.get("domain") for d in rows if d.get("domain") and d.get("isActive",True) and not d.get("isPrivate",False)]
            self.domains=domains
            self.ui(lambda:(self.domain_combo.configure(values=domains),self.bulk_domain_combo.configure(values=domains),self.domain_var.set(self.domain_var.get() or (domains[0] if domains else "")),self.bulk_domain.set(self.bulk_domain.get() or (domains[0] if domains else ""))))
            self.status(f"Đã tải {len(domains)} domain Mail.tm.")
        except Exception as e:self.status(f"Lỗi domain: {e}");self.log(str(e))

    def _address_from_inputs(self):
        user=self.user_var.get().strip() or random_human_username();domain=self.domain_var.get().strip()
        if "@" in user:user=user.split("@",1)[0]
        user=apply_suffix(user,self.suffix_var.get(),self.suffix_mode.get(),self.suffix_len.get())
        if not domain:raise RuntimeError("Chưa chọn domain")
        return f"{user}@{domain}"

    def create_account(self):
        try:
            address=self._address_from_inputs();password=self.pass_var.get().strip() or random_password();self.status(f"Đang tạo {address}...")
            self.request("POST","/accounts",{"address":address,"password":password});tok=self.request("POST","/token",{"address":address,"password":password})
            self.address=address;self.password=password;self.token=tok.get("token","");self.account_id=tok.get("id","")
            self.store.save_account(self.provider,address,password,address.split("@",1)[1],{"account_id":self.account_id},"THÀNH CÔNG")
            self.ui(lambda:(self.pass_var.set(password),self.current_var.set(address)))
            self.status("Tạo tài khoản và đăng nhập thành công.");self.log(f"✅ {address}");self.load_accounts();self.refresh_inbox()
        except Exception as e:self.status(f"Tạo lỗi: {e}");self.log(f"❌ {e}")

    def login_account(self):
        try:
            user=self.user_var.get().strip();domain=self.domain_var.get().strip();address=user if "@" in user else f"{user}@{domain}";password=self.pass_var.get().strip()
            if not address or not password:raise RuntimeError("Thiếu email hoặc mật khẩu")
            tok=self.request("POST","/token",{"address":address,"password":password});self.address=address;self.password=password;self.token=tok.get("token","");self.current_var.set(address)
            self.store.save_account(self.provider,address,password,address.split("@",1)[1],{},"ĐĂNG NHẬP OK")
            self.status("Đăng nhập Mail.tm thành công.");self.refresh_inbox();self.load_accounts()
        except Exception as e:self.status(f"Đăng nhập lỗi: {e}");self.log(str(e))

    def refresh_inbox(self):
        if not self.token:return self.status("Chưa đăng nhập Mail.tm.")
        try:
            j=self.request("GET","/messages?page=1",token=self.token);self.inbox=j.get("hydra:member") or j.get("member") or []
            def render():
                self.inbox_tree.delete(*self.inbox_tree.get_children())
                for i,m in enumerate(self.inbox):
                    frm=m.get("from") or {};sender=frm.get("address") or frm.get("name") or ""
                    self.inbox_tree.insert("","end",iid=f"m{i}",values=(sender,m.get("subject") or "(Không tiêu đề)",_fmt_time(m.get("createdAt") or m.get("created_at") or "")))
            self.ui(render);self.store.save_account(self.provider,self.address,self.password,self.address.split("@",1)[1],{},"HOẠT ĐỘNG",mail_count=len(self.inbox),checked=True);self.status(f"Inbox có {len(self.inbox)} mail.");self.load_accounts()
        except Exception as e:self.status(f"Inbox lỗi: {e}");self.log(str(e))

    def open_message(self,_evt=None):
        sel=self.inbox_tree.selection()
        if not sel:return
        idx=int(sel[0][1:]);row=self.inbox[idx];mid=row.get("id")
        self.worker(self._load_message,mid,row)

    def _load_message(self,mid,row):
        try:
            m=self.request("GET",f"/messages/{mid}",token=self.token);self.current_msg=m
            body=m.get("text") or m.get("intro") or "";html_body=m.get("html") or ""
            if isinstance(html_body,list):html_body="\n".join(map(str,html_body))
            if isinstance(body,list):body="\n".join(map(str,body))
            if not body:body=_html_to_text(html_body)
            u,o,l=detect_shopee_data(m.get("subject",""),body,html_body)
            def render():
                self.subject_var.set(m.get("subject") or "(Không tiêu đề)");self.otp_var.set(o);self.username_var.set(u);self.link_var.set(l);self.body_text.delete("1.0","end");self.body_text.insert("1.0",body)
            self.ui(render);self.status("Đã đọc mail.")
        except Exception as e:self.status(f"Đọc mail lỗi: {e}")

    def find_latest_verify_mail(self):
        if not self.token:return self.status("Chưa đăng nhập Mail.tm.")
        if not self.inbox:self.refresh_inbox()
        for i,row in enumerate(list(self.inbox)[:30]):
            try:
                mid=row.get("id");m=self.request("GET",f"/messages/{mid}",token=self.token)
                body=m.get("text") or m.get("intro") or "";h=m.get("html") or ""
                if isinstance(h,list):h="\n".join(map(str,h))
                if isinstance(body,list):body="\n".join(map(str,body))
                body=body or _html_to_text(h);u,o,l=detect_shopee_data(m.get("subject",""),body,h)
                if not l:continue
                self.current_msg=m
                def render():
                    self.subject_var.set(m.get("subject") or "(Không tiêu đề)");self.otp_var.set(o);self.username_var.set(u);self.link_var.set(l);self.body_text.delete("1.0","end");self.body_text.insert("1.0",body)
                    iid=f"m{i}"
                    if iid in self.inbox_tree.get_children():self.inbox_tree.selection_set(iid);self.inbox_tree.see(iid)
                self.ui(render);self.status("Đã tìm thấy link xác thực Shopee. Bấm Copy link để sao chép.");return
            except Exception as e:self.log(f"Bỏ qua mail khi lọc link: {e}")
        self.status("Không tìm thấy link xác thực Shopee trong 30 mail gần nhất.")

    def save_current_mail(self):
        m=self.current_msg
        if not m or not self.address:return self.status("Chưa chọn mail.")
        frm=m.get("from") or {};sender=frm.get("address") or frm.get("name") or "";body=self.body_text.get("1.0","end-1c");h=m.get("html") or "";h="\n".join(h) if isinstance(h,list) else str(h or "")
        row=self.store.save_message(self.provider,self.address,str(m.get("id")),m.get("subject",""),sender,m.get("createdAt", ""),body,h,raw=m)
        typed=self.username_var.get().strip()
        if typed and typed!=row.get("shopee_username"):self.store.update_username(row["id"],typed)
        self.status("Đã lưu mail vào kho.");self.load_saved();self.load_accounts()

    def save_current_username(self):
        m=self.current_msg
        if not m:return self.status("Chưa chọn mail.")
        self.save_current_mail()

    def load_saved(self):
        rows=self.store.list_messages(self.provider,self.saved_search.get() if hasattr(self,"saved_search") else "")
        if hasattr(self,"verify_only_var") and self.verify_only_var.get():
            rows=[r for r in rows if str(r.get("verify_link") or "").strip()]
        self._saved_rows={str(r["id"]):r for r in rows}
        if not hasattr(self,"saved_tree"):return
        self.saved_tree.delete(*self.saved_tree.get_children())
        for r in rows:self.saved_tree.insert("","end",iid=str(r["id"]),values=(r["account_address"],r["shopee_username"],r["shopee_otp"],r["subject"],_fmt_time(r["received_at"] or r["saved_at"])))

    def copy_selected_saved_link(self):
        if not hasattr(self,"saved_tree"):return
        sel=self.saved_tree.selection()
        if not sel:return self.status("Chưa chọn mail có link xác thực.")
        row=self._saved_rows.get(sel[0])
        link=str((row or {}).get("verify_link") or "").strip()
        if not link:return self.status("Mail đã chọn chưa có link xác thực Shopee. Hãy Quét lại hoặc mở mail trước.")
        self.copy(link);self.status("Đã copy link xác thực Shopee.")

    def open_saved(self,_evt=None):
        sel=self.saved_tree.selection()
        if not sel:return
        r=self._saved_rows.get(sel[0]);
        if not r:return
        self.nb.select(self.inbox_tab);self.subject_var.set(r["subject"]);self.username_var.set(r["shopee_username"]);self.otp_var.set(r["shopee_otp"]);self.link_var.set(r["verify_link"]);self.body_text.delete("1.0","end");self.body_text.insert("1.0",r["body"])

    def rescan_saved(self):
        n=self.store.rescan(self.provider);self.load_saved();self.status(f"Đã quét lại Username/OTP: cập nhật {n} mail.")

    def delete_saved(self):
        ids=[int(x) for x in self.saved_tree.selection()]
        if not ids:return self.status("Chưa chọn mail.")
        if messagebox.askyesno("Xóa mail",f"Xóa {len(ids)} mail đã chọn?"):
            n=self.store.delete_messages(self.provider,ids);self.load_saved();self.status(f"Đã xóa {n} mail.")

    def export_saved(self):
        rows=self.store.list_messages(self.provider,self.saved_search.get())
        self.export_csv("mailtm_saved.csv",["email","username_shopee","otp","verify_link","subject","received_at","body"],[[r["account_address"],r["shopee_username"],r["shopee_otp"],r["verify_link"],r["subject"],r["received_at"],r["body"]] for r in rows])

    def bulk_create(self):
        if not self.bulk_domain.get():return self.status("Chưa chọn domain hàng loạt.")
        self.bulk_stop.clear();count=max(1,min(200,int(self.bulk_count.get())));delay=max(.2,float(self.bulk_delay.get()));ok=0
        for i in range(count):
            if self.bulk_stop.is_set() or self.stop_event.is_set():break
            local=(self.bulk_prefix.get().strip() or random_human_username())+f"{random.randint(0,9999):04d}";address=f"{local}@{self.bulk_domain.get()}";password=self.bulk_password.get().strip() or random_password()
            try:
                self.request("POST","/accounts",{"address":address,"password":password});self.store.save_account(self.provider,address,password,self.bulk_domain.get(),{},"THÀNH CÔNG");ok+=1;self.log(f"✅ Tạo {address}")
            except Exception as e:self.log(f"❌ {address}: {e}")
            self.status(f"Tạo hàng loạt {i+1}/{count} | Thành công {ok}");time.sleep(delay)
        self.load_accounts();self.status(f"Hoàn tất tạo hàng loạt: {ok}/{count} thành công.")

    def load_accounts(self):
        rows=self.store.list_accounts(self.provider,self.account_search.get() if hasattr(self,"account_search") else "");rows=[r for r in rows if self._account_filter_ok(r)];self._accounts={r["address"]:r for r in rows}
        if not hasattr(self,"account_tree"):return
        self.account_tree.delete(*self.account_tree.get_children())
        for r in rows:
            active=self._status_is_active(r.get("status"));tag="active" if active else "inactive"
            self.account_tree.insert("","end",iid=r["address"],values=("☑" if r["address"] in self._checked_accounts else "☐",r["address"],r.get("shopee_usernames", ""),r["status"],r.get("last_error", ""),r["mail_count"] if r["mail_count"] is not None else "",_fmt_time(r["last_check"]),_fmt_time(r["created_at"])),tags=(tag,))

    def use_account(self,_evt=None):
        sel=self.account_tree.selection();
        if not sel:return
        r=self._accounts.get(sel[0]);
        if not r:return
        address=r["address"];self.user_var.set(address);self.pass_var.set(r["password"]);self.domain_var.set(address.split("@",1)[1]);self.worker(self.login_account);self.nb.select(self.inbox_tab)

    def load_created_to_check(self):
        rows=self.store.list_accounts(self.provider);self.check_input.delete("1.0","end");self.check_input.insert("1.0","\n".join(f"{r['address']}|{r['password']}" for r in rows))

    def check_accounts(self):
        lines=[x.strip() for x in self.check_input.get("1.0","end").splitlines() if x.strip()];self.check_stop.clear();self.check_results=[]
        for i,line in enumerate(lines,1):
            if self.check_stop.is_set() or self.stop_event.is_set():break
            p=line.split("|",1);address=p[0].strip();password=p[1].strip() if len(p)>1 else ""
            if not password:
                rows=self.store.list_accounts(self.provider,address);password=next((x["password"] for x in rows if x["address"].lower()==address.lower()),"")
            try:
                tok=self.request("POST","/token",{"address":address,"password":password});j=self.request("GET","/messages?page=1",token=tok.get("token"));m=j.get("hydra:member") or [];status="HOẠT ĐỘNG";note="Đăng nhập và Inbox OK";count=len(m);http=200
            except Exception as e:status="KHÔNG HOẠT ĐỘNG";note=str(e);count=0;http=401 if "401" in str(e) else None
            self.store.save_account(self.provider,address,password,address.split("@",1)[1] if "@" in address else "",{},status,note,count,True);self.check_results.append([address,status,http,count,_now_iso(),note]);self.log(f"{'✅' if self._status_is_active(status) else '⚠'} {address} | {status} | {note}");self.status(f"Check {i}/{len(lines)} | {status}")
        self.load_accounts();self.status(f"Check hoàn tất: {len(self.check_results)}/{len(lines)}")

    def export_checks(self):self.export_csv("mailtm_check.csv",["email","status","http","mail_count","checked_at","note"],self.check_results)
    def export_accounts(self):
        rows=self.store.list_accounts(self.provider,self.account_search.get());self.export_csv("mailtm_accounts.csv",["email","password","username_shopee","status","mail_count","last_check","created_at"],[[r["address"],r["password"],r.get("shopee_usernames",""),r["status"],r["mail_count"],r["last_check"],r["created_at"]] for r in rows])


class TinyHostTab(MailTMTab):
    provider="tinyhost"
    def __init__(self,parent,store):
        BaseMailTab.__init__(self,parent,store)
        self.address="";self.domain="";self.user="";self.domains=[];self.inbox=[];self.current_msg=None
        self.bulk_stop=threading.Event();self.check_stop=threading.Event();self._build_tiny();self.worker(self.refresh_domains);self.load_saved();self.load_accounts()

    def request_tiny(self,method,path):
        if httpx is None:raise RuntimeError("Thiếu httpx")
        with httpx.Client(timeout=25,headers={"Accept":"application/json"},follow_redirects=True) as c:r=c.request(method,TINYHOST_API+path)
        try:j=r.json()
        except Exception:j={"message":r.text[:500]}
        if r.status_code>=400:raise RuntimeError(f"HTTP {r.status_code}: {j.get('detail') or j.get('message') or r.reason_phrase}")
        return j

    def _build_tiny(self):
        top=tk.Frame(self,bg="#f6f7f9");top.pack(fill="x",padx=12,pady=(12,6));tk.Label(top,text="TinyHost • Tạo/chọn địa chỉ • Đọc Inbox • Check MX",bg="#f6f7f9",fg="#222",font=("Segoe UI",13,"bold")).pack(side="left");ttk.Button(top,text="💾 Backup JSON",command=self.backup_json,style="Soft.TButton").pack(side="right",padx=3);ttk.Button(top,text="📥 Import JSON",command=self.import_json,style="Soft.TButton").pack(side="right",padx=3)
        login=tk.Frame(self,bg="#fff",highlightthickness=1,highlightbackground="#e4e7eb");login.pack(fill="x",padx=12,pady=6);[login.grid_columnconfigure(i,weight=1) for i in range(7)]
        self.user_var=tk.StringVar();self.domain_var=tk.StringVar();self.custom_domain_var=tk.StringVar();self.suffix_var=tk.BooleanVar();self.suffix_mode=tk.StringVar(value="alnum");self.suffix_len=tk.IntVar(value=4)
        ttk.Entry(login,textvariable=self.user_var).grid(row=0,column=0,sticky="ew",padx=(10,4),pady=8);self.domain_combo=ttk.Combobox(login,textvariable=self.domain_var,state="readonly");self.domain_combo.grid(row=0,column=1,sticky="ew",padx=4);ttk.Entry(login,textvariable=self.custom_domain_var).grid(row=0,column=2,sticky="ew",padx=4);ttk.Button(login,text="↻ Domain",command=lambda:self.worker(self.refresh_domains),style="Soft.TButton").grid(row=0,column=3,sticky="ew",padx=4);ttk.Button(login,text="✓ MX",command=lambda:self.worker(self.check_current_mx),style="Soft.TButton").grid(row=0,column=4,sticky="ew",padx=4);ttk.Button(login,text="🎲 Random",command=lambda:self.user_var.set(random_human_username()),style="Soft.TButton").grid(row=0,column=5,sticky="ew",padx=4);ttk.Button(login,text="➕ Tạo/Chọn",command=lambda:self.worker(self.select_address),style="Accent.TButton").grid(row=0,column=6,sticky="ew",padx=(4,10))
        ttk.Checkbutton(login,text="Random cuối tên",variable=self.suffix_var).grid(row=1,column=0,sticky="w",padx=10,pady=(0,8));ttk.Combobox(login,textvariable=self.suffix_mode,values=["alnum","digits","letters"],state="readonly",width=10).grid(row=1,column=1,sticky="w",padx=4,pady=(0,8));ttk.Spinbox(login,from_=1,to=12,textvariable=self.suffix_len,width=5).grid(row=1,column=2,sticky="w",padx=4,pady=(0,8));ttk.Button(login,text="📬 Mở địa chỉ có sẵn",command=lambda:self.worker(self.select_address),style="Soft.TButton").grid(row=1,column=6,sticky="ew",padx=(4,10),pady=(0,8))
        current=tk.Frame(self,bg="#fff");current.pack(fill="x",padx=12,pady=(0,6));self.current_var=tk.StringVar(value="Chưa chọn địa chỉ");tk.Label(current,textvariable=self.current_var,bg="#fff",fg="#ee4d2d",font=("Segoe UI",9,"bold")).pack(side="left",padx=10,pady=7);ttk.Button(current,text="Copy email",command=lambda:self.copy(self.address),style="Soft.TButton").pack(side="right",padx=4,pady=4);ttk.Button(current,text="↻ Inbox",command=lambda:self.worker(self.refresh_inbox),style="Soft.TButton").pack(side="right",padx=4,pady=4)
        self.nb=ttk.Notebook(self);self.nb.pack(fill="both",expand=True,padx=12,pady=6);self.inbox_tab=tk.Frame(self.nb,bg="#f6f7f9");self.saved_tab=tk.Frame(self.nb,bg="#f6f7f9");self.bulk_tab=tk.Frame(self.nb,bg="#f6f7f9");self.nb.add(self.inbox_tab,text="📥 Inbox / Đọc mail");self.nb.add(self.saved_tab,text="💾 Mail đã lưu / Tìm Shopee");self.nb.add(self.bulk_tab,text="⚡ Tạo hàng loạt / DS địa chỉ");self._build_inbox();self._build_saved();self._build_tiny_bulk()
        self.build_log_panel()

    def _build_tiny_bulk(self):
        tab=self.bulk_tab;tab.grid_columnconfigure(0,weight=1);tab.grid_columnconfigure(1,weight=1);tab.grid_rowconfigure(1,weight=1)
        bulk=tk.LabelFrame(tab,text="Tạo hàng loạt",bg="#fff",fg="#222");bulk.grid(row=0,column=0,sticky="ew",padx=(0,5),pady=5);self.bulk_count=tk.IntVar(value=10);self.bulk_prefix=tk.StringVar();self.bulk_domain=tk.StringVar();self.bulk_custom=tk.StringVar();self.bulk_random_domain=tk.BooleanVar(value=False)
        for i in range(4):bulk.grid_columnconfigure(i,weight=1)
        ttk.Spinbox(bulk,from_=1,to=500,textvariable=self.bulk_count,width=7).grid(row=0,column=0,padx=4,pady=7);ttk.Entry(bulk,textvariable=self.bulk_prefix).grid(row=0,column=1,sticky="ew",padx=4);self.bulk_domain_combo=ttk.Combobox(bulk,textvariable=self.bulk_domain,state="readonly");self.bulk_domain_combo.grid(row=0,column=2,sticky="ew",padx=4);ttk.Entry(bulk,textvariable=self.bulk_custom).grid(row=0,column=3,sticky="ew",padx=4)
        ttk.Checkbutton(bulk,text="Random domain",variable=self.bulk_random_domain).grid(row=1,column=0,sticky="w",padx=4);ttk.Button(bulk,text="▶ Bắt đầu tạo",command=lambda:self.worker(self.bulk_create),style="Accent.TButton").grid(row=1,column=1,columnspan=2,sticky="ew",padx=4,pady=5);ttk.Button(bulk,text="■ Dừng",command=self.bulk_stop.set,style="Danger.TButton").grid(row=1,column=3,sticky="ew",padx=4,pady=5)
        checker=tk.LabelFrame(tab,text="Check địa chỉ TinyHost",bg="#fff",fg="#222");checker.grid(row=0,column=1,sticky="ew",padx=(5,0),pady=5);checker.grid_columnconfigure(0,weight=1);self.check_input=scrolledtext.ScrolledText(checker,height=4,font=("Consolas",8));self.check_input.grid(row=0,column=0,columnspan=4,sticky="ew",padx=5,pady=5);ttk.Button(checker,text="📥 Nạp DS",command=self.load_created_to_check,style="Soft.TButton").grid(row=1,column=0,sticky="ew",padx=3,pady=4);ttk.Button(checker,text="✓ Check",command=lambda:self.worker(self.check_accounts),style="Accent.TButton").grid(row=1,column=1,sticky="ew",padx=3,pady=4);ttk.Button(checker,text="■ Dừng",command=self.check_stop.set,style="Danger.TButton").grid(row=1,column=2,sticky="ew",padx=3,pady=4);ttk.Button(checker,text="📤 CSV",command=self.export_checks,style="Soft.TButton").grid(row=1,column=3,sticky="ew",padx=3,pady=4)
        acc=tk.Frame(tab,bg="#fff");acc.grid(row=1,column=0,columnspan=2,sticky="nsew",pady=5);acc.grid_rowconfigure(1,weight=1);acc.grid_columnconfigure(0,weight=1);abar=tk.Frame(acc,bg="#fff");abar.grid(row=0,column=0,sticky="ew");self.account_search=tk.StringVar();self.account_filter_var=tk.StringVar(value="Tất cả");ttk.Entry(abar,textvariable=self.account_search).pack(side="left",fill="x",expand=True,padx=6,pady=5);ttk.Button(abar,text="🔎 Tìm",command=self.load_accounts,style="Soft.TButton").pack(side="left",padx=2);ttk.Combobox(abar,textvariable=self.account_filter_var,values=["Tất cả","Hoạt động","Mail lỗi / không hoạt động"],state="readonly",width=22).pack(side="left",padx=2);ttk.Button(abar,text="Lọc",command=self.load_accounts,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(abar,text="☑ Chọn tất cả",command=self.check_all_visible_accounts,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(abar,text="☐ Bỏ chọn",command=self.clear_account_checks,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(abar,text="🗑 Xóa tích",command=self.delete_checked_accounts,style="Danger.TButton").pack(side="left",padx=2);ttk.Button(abar,text="📤 CSV",command=self.export_accounts,style="Soft.TButton").pack(side="left",padx=2)
        self.account_tree=ttk.Treeview(acc,columns=("pick","email","user","domain","status","error","count","checked","created"),show="headings",selectmode="extended");
        for c,t,w in (("pick","✓",42),("email","Địa chỉ",220),("user","Username Shopee",150),("domain","Domain",140),("status","Trạng thái",135),("error","Lỗi / ghi chú",250),("count","Số mail",65),("checked","Check lúc",140),("created","Tạo lúc",140)):self.account_tree.heading(c,text=t);self.account_tree.column(c,width=w,anchor="center" if c=="pick" else "w")
        self.account_tree.grid(row=1,column=0,sticky="nsew");self.account_tree.bind("<Double-1>",self.use_account);self.configure_account_tree();self.check_results=[]

    def refresh_domains(self):
        try:
            domains=[]
            for path in ("/api/all-domains/","/api/random-domains/?limit=50"):
                try:
                    j=self.request_tiny("GET",path);raw=j.get("domains") if isinstance(j,dict) else j
                    if isinstance(raw,dict):raw=list(raw.values())
                    if isinstance(raw,list):
                        for x in raw:
                            d=x.get("domain") if isinstance(x,dict) else str(x)
                            if d and d not in domains:domains.append(d)
                    if domains:break
                except Exception:continue
            self.domains=domains;self.ui(lambda:(self.domain_combo.configure(values=domains),self.bulk_domain_combo.configure(values=domains),self.domain_var.set(self.domain_var.get() or (domains[0] if domains else "")),self.bulk_domain.set(self.bulk_domain.get() or (domains[0] if domains else ""))));self.status(f"Đã tải {len(domains)} domain TinyHost.")
        except Exception as e:self.status(f"Lỗi domain: {e}")

    def current_domain(self):return self.custom_domain_var.get().strip() or self.domain_var.get().strip()
    def check_mx(self,domain):
        j=self.request_tiny("GET",f"/api/check-mx/{domain}");return str(j.get("result","")).lower()=="online"
    def check_current_mx(self):
        d=self.current_domain();
        if not d:return self.status("Chưa có domain.")
        try:self.status(f"MX {d}: {'ONLINE' if self.check_mx(d) else 'OFFLINE'}")
        except Exception as e:self.status(f"Check MX lỗi: {e}")

    def select_address(self):
        try:
            user=self.user_var.get().strip() or random_human_username();
            if "@" in user:user,typed_domain=user.split("@",1);domain=typed_domain
            else:domain=self.current_domain()
            user=apply_suffix(user,self.suffix_var.get(),self.suffix_mode.get(),self.suffix_len.get())
            if not re.fullmatch(r"[A-Za-z0-9._-]+",user) or not domain:raise RuntimeError("Địa chỉ không hợp lệ")
            if not self.check_mx(domain):raise RuntimeError("MX domain đang offline/không hợp lệ")
            self.user=user;self.domain=domain;self.address=f"{user}@{domain}";self.store.save_account(self.provider,self.address,"",domain,{},"HOẠT ĐỘNG");self.ui(lambda:self.current_var.set(self.address));self.status(f"Đã chọn {self.address}");self.refresh_inbox();self.load_accounts()
        except Exception as e:self.status(f"Không mở được địa chỉ: {e}")

    def refresh_inbox(self):
        if not self.address:return self.status("Chưa chọn địa chỉ TinyHost.")
        try:
            j=self.request_tiny("GET",f"/api/email/{self.domain}/{self.user}/?page=1&limit=100");self.inbox=j.get("emails") or []
            def render():
                self.inbox_tree.delete(*self.inbox_tree.get_children())
                for i,m in enumerate(self.inbox):self.inbox_tree.insert("","end",iid=f"m{i}",values=(m.get("sender") or "",m.get("subject") or "(Không tiêu đề)",_fmt_time(m.get("date") or "")))
            self.ui(render);self.store.save_account(self.provider,self.address,"",self.domain,{},"HOẠT ĐỘNG",mail_count=int(j.get("total") or len(self.inbox)),checked=True);self.status(f"Inbox có {len(self.inbox)}/{j.get('total',len(self.inbox))} mail.");self.load_accounts()
        except Exception as e:self.status(f"Inbox lỗi: {e}")

    def _load_message(self,mid,row):
        try:
            m=self.request_tiny("GET",f"/api/email/{self.domain}/{self.user}/{mid}");merged={**row,**(m if isinstance(m,dict) else {})};self.current_msg=merged;body=str(merged.get("body") or merged.get("text") or merged.get("content") or "");h=str(merged.get("html_body") or merged.get("html") or "");body=body or _html_to_text(h);u,o,l=detect_shopee_data(merged.get("subject",""),body,h)
            def render():self.subject_var.set(merged.get("subject") or "(Không tiêu đề)");self.otp_var.set(o);self.username_var.set(u);self.link_var.set(l);self.body_text.delete("1.0","end");self.body_text.insert("1.0",body)
            self.ui(render);self.status("Đã đọc mail TinyHost.")
        except Exception as e:self.status(f"Đọc mail lỗi: {e}")

    def find_latest_verify_mail(self):
        if not self.address:return self.status("Chưa chọn địa chỉ TinyHost.")
        if not self.inbox:self.refresh_inbox()
        for i,row in enumerate(list(self.inbox)[:30]):
            try:
                mid=row.get("id");m=self.request_tiny("GET",f"/api/email/{self.domain}/{self.user}/{mid}");merged={**row,**(m if isinstance(m,dict) else {})}
                body=str(merged.get("body") or merged.get("text") or merged.get("content") or "");h=str(merged.get("html_body") or merged.get("html") or "");body=body or _html_to_text(h);u,o,l=detect_shopee_data(merged.get("subject",""),body,h)
                if not l:continue
                self.current_msg=merged
                def render():
                    self.subject_var.set(merged.get("subject") or "(Không tiêu đề)");self.otp_var.set(o);self.username_var.set(u);self.link_var.set(l);self.body_text.delete("1.0","end");self.body_text.insert("1.0",body)
                    iid=f"m{i}"
                    if iid in self.inbox_tree.get_children():self.inbox_tree.selection_set(iid);self.inbox_tree.see(iid)
                self.ui(render);self.status("Đã tìm thấy link xác thực Shopee. Bấm Copy link để sao chép.");return
            except Exception as e:self.log(f"Bỏ qua mail khi lọc link: {e}")
        self.status("Không tìm thấy link xác thực Shopee trong 30 mail gần nhất.")

    def save_current_mail(self):
        m=self.current_msg
        if not m or not self.address:return self.status("Chưa chọn mail.")
        body=self.body_text.get("1.0","end-1c");h=str(m.get("html_body") or m.get("html") or "");row=self.store.save_message(self.provider,self.address,str(m.get("id")),m.get("subject",""),str(m.get("sender") or ""),str(m.get("date") or ""),body,h,raw=m);typed=self.username_var.get().strip();
        if typed and typed!=row.get("shopee_username"):self.store.update_username(row["id"],typed)
        self.status("Đã lưu mail TinyHost.");self.load_saved();self.load_accounts()

    def bulk_create(self):
        self.bulk_stop.clear();count=max(1,min(500,int(self.bulk_count.get())));ok=0
        for i in range(count):
            if self.bulk_stop.is_set() or self.stop_event.is_set():break
            domain=self.bulk_custom.get().strip() or (random.choice(self.domains) if self.bulk_random_domain.get() and self.domains else self.bulk_domain.get().strip())
            local=(self.bulk_prefix.get().strip() or random_human_username())+f"{random.randint(0,9999):04d}"
            try:
                if not domain or not self.check_mx(domain):raise RuntimeError("Domain offline")
                address=f"{local}@{domain}";self.store.save_account(self.provider,address,"",domain,{},"HOẠT ĐỘNG");ok+=1;self.log(f"✅ {address}")
            except Exception as e:self.log(f"❌ {local}@{domain}: {e}")
            self.status(f"Tạo DS {i+1}/{count} | OK {ok}")
        self.load_accounts();self.status(f"Hoàn tất tạo DS: {ok}/{count}")

    def load_created_to_check(self):
        rows=self.store.list_accounts(self.provider);self.check_input.delete("1.0","end");self.check_input.insert("1.0","\n".join(r["address"] for r in rows))

    def check_accounts(self):
        lines=[x.strip() for x in self.check_input.get("1.0","end").splitlines() if x.strip()];self.check_stop.clear();self.check_results=[];mx={}
        for i,address in enumerate(lines,1):
            if self.check_stop.is_set() or self.stop_event.is_set():break
            try:
                user,domain=address.split("@",1)
                if domain not in mx:
                    try:mx[domain]=self.check_mx(domain)
                    except Exception:mx[domain]=None
                if mx[domain] is False:status="DOMAIN OFFLINE";http=None;count=None;note="MX domain offline"
                elif mx[domain] is None:status="CHƯA XÁC ĐỊNH";http=None;count=None;note="Không kiểm tra được MX"
                else:
                    try:j=self.request_tiny("GET",f"/api/email/{domain}/{user}/?page=1&limit=1");status="HOẠT ĐỘNG";http=200;count=int(j.get("total") or len(j.get("emails") or []));note="Inbox phản hồi thành công"
                    except Exception as e:
                        status="KHÔNG HOẠT ĐỘNG" if "404" in str(e) else "CHƯA XÁC ĐỊNH";http=404 if "404" in str(e) else None;count=0 if http==404 else None;note=str(e)
            except Exception as e:status="KHÔNG HỢP LỆ";http=None;count=None;note=str(e)
            domain=address.split("@",1)[1] if "@" in address else "";self.store.save_account(self.provider,address,"",domain,{},status,note,count,True);self.check_results.append([address,status,http,count,_now_iso(),note]);self.log(f"{'✅' if self._status_is_active(status) else '⚠'} {address} | {status} | {note}");self.status(f"Check {i}/{len(lines)} | {status}")
        self.load_accounts();self.status(f"Check TinyHost hoàn tất {len(self.check_results)}/{len(lines)}")

    def use_account(self,_evt=None):
        sel=self.account_tree.selection();
        if not sel:return
        address=sel[0];self.user_var.set(address);self.custom_domain_var.set(address.split("@",1)[1] if "@" in address else "");self.worker(self.select_address);self.nb.select(self.inbox_tab)

    def load_accounts(self):
        rows=self.store.list_accounts(self.provider,self.account_search.get() if hasattr(self,"account_search") else "");rows=[r for r in rows if self._account_filter_ok(r)];self._accounts={r["address"]:r for r in rows}
        if not hasattr(self,"account_tree"):return
        self.account_tree.delete(*self.account_tree.get_children())
        for r in rows:
            active=self._status_is_active(r.get("status"));tag="active" if active else "inactive"
            self.account_tree.insert("","end",iid=r["address"],values=("☑" if r["address"] in self._checked_accounts else "☐",r["address"],r.get("shopee_usernames", ""),r["domain"],r["status"],r.get("last_error", ""),r["mail_count"] if r["mail_count"] is not None else "",_fmt_time(r["last_check"]),_fmt_time(r["created_at"])),tags=(tag,))

    def export_checks(self):self.export_csv("tinyhost_check.csv",["address","status","http","mail_count","checked_at","note"],self.check_results)
    def export_accounts(self):
        rows=self.store.list_accounts(self.provider,self.account_search.get());self.export_csv("tinyhost_accounts.csv",["address","username_shopee","domain","status","mail_count","last_check","created_at"],[[r["address"],r.get("shopee_usernames",""),r["domain"],r["status"],r["mail_count"],r["last_check"],r["created_at"]] for r in rows])


class OutlookTab(BaseMailTab):
    provider="outlook"
    def __init__(self,parent,store):
        super().__init__(parent,store)
        self.session_acct=None;self.current_message=None;self._saved={};self._accounts={}
        self._build();
        if outlook_backend:
            try:outlook_backend.init_db()
            except Exception as e:self.status(f"Lỗi database Outlook: {e}")
        self.load_saved();self.load_accounts()

    def _build(self):
        top=tk.Frame(self,bg="#f6f7f9");top.pack(fill="x",padx=12,pady=(12,6));tk.Label(top,text="Outlook / Hotmail • Đọc trực tiếp bằng refresh token + client ID",bg="#f6f7f9",fg="#222",font=("Segoe UI",13,"bold")).pack(side="left");ttk.Button(top,text="⬇ Backup JSON",command=self.backup_outlook,style="Soft.TButton").pack(side="right",padx=3);ttk.Button(top,text="⬆ Import JSON",command=self.import_outlook,style="Soft.TButton").pack(side="right",padx=3)
        self.nb=ttk.Notebook(self);self.nb.pack(fill="both",expand=True,padx=12,pady=6);self.read_tab=tk.Frame(self.nb,bg="#f6f7f9");self.saved_tab=tk.Frame(self.nb,bg="#f6f7f9");self.accounts_tab=tk.Frame(self.nb,bg="#f6f7f9");self.nb.add(self.read_tab,text="📨 Đọc mail");self.nb.add(self.saved_tab,text="💾 Mail đã lưu");self.nb.add(self.accounts_tab,text="👤 Tài khoản đã lưu");self._build_read();self._build_saved_outlook();self._build_accounts_outlook()
        self.build_log_panel()

    def _build_read(self):
        tab=self.read_tab;tab.grid_columnconfigure(0,weight=1);tab.grid_rowconfigure(3,weight=1)
        tk.Label(tab,text="Full token: email|password|refresh_token|client_id",bg="#f6f7f9",fg="#333",font=("Segoe UI",9,"bold")).grid(row=0,column=0,sticky="w",pady=(6,2))
        self.full_line=scrolledtext.ScrolledText(tab,height=3,font=("Consolas",8));self.full_line.grid(row=1,column=0,sticky="ew")
        bar=tk.Frame(tab,bg="#f6f7f9");bar.grid(row=2,column=0,sticky="ew",pady=5);ttk.Button(bar,text="📨 Đọc mail mới nhất",command=lambda:self.worker(self.read_current,False),style="Accent.TButton").pack(side="left",padx=2);ttk.Button(bar,text="🛍 Đọc mail Shopee",command=lambda:self.worker(self.read_current,True),style="Soft.TButton").pack(side="left",padx=2);ttk.Button(bar,text="🔐 Lọc link xác thực",command=lambda:self.worker(self.read_verify_link),style="Soft.TButton").pack(side="left",padx=2);ttk.Button(bar,text="🔄 Làm mới token",command=lambda:self.worker(self.renew_current),style="Soft.TButton").pack(side="left",padx=2);ttk.Button(bar,text="💾 Lưu tài khoản",command=self.save_account,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(bar,text="↻ Đọc lại",command=lambda:self.worker(self.refresh_current),style="Soft.TButton").pack(side="left",padx=2);ttk.Button(bar,text="Xóa ô",command=lambda:(self.full_line.delete("1.0","end"),self.clear_message()),style="Danger.TButton").pack(side="left",padx=2)
        box=tk.Frame(tab,bg="#fff");box.grid(row=3,column=0,sticky="nsew",pady=5);box.grid_columnconfigure(0,weight=1);box.grid_rowconfigure(4,weight=1);self.subj=tk.StringVar(value="Chưa đọc mail");self.meta=tk.StringVar();self.otp=tk.StringVar();self.username=tk.StringVar();self.link=tk.StringVar();tk.Label(box,textvariable=self.subj,bg="#fff",fg="#222",font=("Segoe UI",10,"bold"),wraplength=1100,justify="left").grid(row=0,column=0,sticky="ew",padx=10,pady=(10,3));tk.Label(box,textvariable=self.meta,bg="#fff",fg="#777").grid(row=1,column=0,sticky="ew",padx=10)
        det=tk.Frame(box,bg="#fff");det.grid(row=2,column=0,sticky="ew",padx=10,pady=5);ttk.Entry(det,textvariable=self.otp,width=12).pack(side="left");ttk.Button(det,text="📋 OTP",command=lambda:self.copy(self.otp.get()),style="Soft.TButton").pack(side="left",padx=2);ttk.Entry(det,textvariable=self.username,width=24).pack(side="left",padx=(8,2));ttk.Button(det,text="💾 Lưu mail",command=self.save_current_message,style="Accent.TButton").pack(side="left",padx=2);ttk.Entry(det,textvariable=self.link).pack(side="left",fill="x",expand=True,padx=(8,2));ttk.Button(det,text="Copy link",command=lambda:self.copy(self.link.get()),style="Soft.TButton").pack(side="left")
        ttk.Button(det, text="Mở link", command=lambda: self.open_email_link(self.link.get()), style="Soft.TButton").pack(side="left", padx=4)
        self.body=scrolledtext.ScrolledText(box,wrap="word",font=("Segoe UI",9));self.body.grid(row=4,column=0,sticky="nsew",padx=10,pady=(5,10))

    def _build_saved_outlook(self):
        t=self.saved_tab;t.grid_columnconfigure(0,weight=1);t.grid_rowconfigure(1,weight=1);b=tk.Frame(t,bg="#fff");b.grid(row=0,column=0,sticky="ew",pady=5);self.saved_search=tk.StringVar();self.verify_only_var=tk.BooleanVar(value=False);ttk.Entry(b,textvariable=self.saved_search).pack(side="left",fill="x",expand=True,padx=6,pady=5);ttk.Button(b,text="↻ Nạp lại",command=self.load_saved,style="Soft.TButton").pack(side="left",padx=2);ttk.Checkbutton(b,text="🔐 Chỉ link xác thực",variable=self.verify_only_var,command=self.load_saved).pack(side="left",padx=3);ttk.Button(b,text="📋 Copy link",command=self.copy_selected_outlook_link,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(b,text="🔎 Quét lại Shopee",command=self.rescan_outlook,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(b,text="📤 CSV",command=self.export_saved_outlook,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(b,text="🗑 Xóa đã chọn",command=self.delete_saved_outlook,style="Danger.TButton").pack(side="left",padx=2)
        self.saved_tree=ttk.Treeview(t,columns=("email","user","otp","subject","received"),show="headings",selectmode="extended");
        for c,tx,w in (("email","Email",230),("user","Username",160),("otp","OTP",80),("subject","Tiêu đề",360),("received","Thời gian",160)):self.saved_tree.heading(c,text=tx);self.saved_tree.column(c,width=w)
        self.saved_tree.grid(row=1,column=0,sticky="nsew");self.saved_tree.bind("<Double-1>",self.open_saved_outlook)

    def _build_accounts_outlook(self):
        t=self.accounts_tab;t.grid_columnconfigure(0,weight=1);t.grid_rowconfigure(1,weight=1);b=tk.Frame(t,bg="#fff");b.grid(row=0,column=0,sticky="ew",pady=5);self.account_search=tk.StringVar();self.account_filter_var=tk.StringVar(value="Tất cả");ttk.Entry(b,textvariable=self.account_search).pack(side="left",fill="x",expand=True,padx=6,pady=5);ttk.Button(b,text="↻ Nạp lại",command=self.load_accounts,style="Soft.TButton").pack(side="left",padx=2);ttk.Combobox(b,textvariable=self.account_filter_var,values=["Tất cả","Hoạt động","Mail lỗi / không hoạt động"],state="readonly",width=22).pack(side="left",padx=2);ttk.Button(b,text="Lọc",command=self.load_accounts,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(b,text="☑ Chọn tất cả",command=self.check_all_visible_accounts,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(b,text="☐ Bỏ chọn",command=self.clear_account_checks,style="Soft.TButton").pack(side="left",padx=2);ttk.Button(b,text="🗑 Xóa tích",command=self.delete_checked_accounts,style="Danger.TButton").pack(side="left",padx=2);ttk.Button(b,text="📤 CSV",command=self.export_accounts_outlook,style="Soft.TButton").pack(side="left",padx=2)
        self.account_tree=ttk.Treeview(t,columns=("pick","email","method","status","error","token","read"),show="headings",selectmode="extended");
        for c,tx,w in (("pick","✓",42),("email","Email",260),("method","Phương thức",110),("status","Trạng thái",150),("error","Lỗi / ghi chú",280),("token","Token cập nhật",155),("read","Đọc lần cuối",155)):self.account_tree.heading(c,text=tx);self.account_tree.column(c,width=w,anchor="center" if c=="pick" else "w")
        self.account_tree.grid(row=1,column=0,sticky="nsew");self.account_tree.bind("<Double-1>",self.use_outlook_account);self.configure_account_tree()

    def _require_backend(self):
        if outlook_backend is None:raise RuntimeError("Thiếu outlook_native_backend.py")

    def _line(self):return self.full_line.get("1.0","end-1c").strip()
    def set_line(self,line):self.ui(lambda:(self.full_line.delete("1.0","end"),self.full_line.insert("1.0",line)))

    def read_current(self,shopee_only=False):
        try:
            self._require_backend();acct=outlook_backend.parse_full_line(self._line());fn=outlook_backend.read_latest_shopee_for_account if shopee_only else outlook_backend.read_latest_for_account;self.status("Đang đọc Outlook/Hotmail...");r=fn(acct)
            if not r.get("ok"):raise RuntimeError(r.get("message") or "Không đọc được mail")
            old=acct["refresh_token"];acct["refresh_token"]=r.get("refresh_token") or old;saved=outlook_backend.upsert_account(acct,method=r.get("method",""),scope=r.get("scope",""),status="OK",error="",token_changed=acct["refresh_token"]!=old,read_ok=True);self.session_acct=acct;self.set_line(outlook_backend.full_line(saved));self.current_message=r.get("message");self.render_outlook_message(self.current_message,r.get("method",""),acct["email"]);self.status("Đã đọc mail." if self.current_message else "Inbox đọc được nhưng không có mail phù hợp.");self.load_accounts()
        except Exception as e:
            try:
                if 'acct' in locals() and acct:outlook_backend.upsert_account(acct,status="ERROR",error=str(e))
            except Exception:pass
            self.status(f"Đọc mail lỗi: {e}");self.log(str(e));self.load_accounts()

    def read_verify_link(self):
        try:
            self._require_backend();acct=outlook_backend.parse_full_line(self._line());self.status("Đang lọc mail xác thực Shopee...")
            fn=getattr(outlook_backend,"read_latest_shopee_verification_for_account",None)
            if fn is None:raise RuntimeError("Backend chưa có bộ lọc link xác thực Shopee")
            r=fn(acct)
            if not r.get("ok"):raise RuntimeError(r.get("message") or "Không đọc được mail xác thực")
            old=acct["refresh_token"];acct["refresh_token"]=r.get("refresh_token") or old;saved=outlook_backend.upsert_account(acct,method=r.get("method",""),scope=r.get("scope",""),status="OK",error="",token_changed=acct["refresh_token"]!=old,read_ok=True);self.session_acct=acct;self.set_line(outlook_backend.full_line(saved));self.current_message=r.get("message");self.render_outlook_message(self.current_message,r.get("method",""),acct["email"]);self.load_accounts()
            m=self.current_message or {};bodyobj=m.get("body") or {};raw=bodyobj.get("content") or m.get("bodyPreview") or "";is_html=str(bodyobj.get("contentType") or "").lower()=="html";text=_html_to_text(raw) if is_html else str(raw);_,_,verify_link=detect_shopee_data(m.get("subject",""),text,raw if is_html else "")
            if self.current_message and verify_link:self.status("Đã lọc được link xác thực Shopee. Bấm Copy link để sao chép.")
            else:self.status("Không tìm thấy mail có link xác thực Shopee trong 50 mail gần nhất.")
        except Exception as e:
            try:
                if 'acct' in locals() and acct:outlook_backend.upsert_account(acct,status="ERROR",error=str(e))
            except Exception:pass
            self.status(f"Lọc link xác thực lỗi: {e}");self.log(str(e));self.load_accounts()

    def refresh_current(self):
        if not self.session_acct:return self.read_current(False)
        try:
            r=outlook_backend.read_latest_for_account(self.session_acct)
            if not r.get("ok"):raise RuntimeError(r.get("message") or "Không đọc được")
            old=self.session_acct["refresh_token"];self.session_acct["refresh_token"]=r.get("refresh_token") or old;saved=outlook_backend.upsert_account(self.session_acct,method=r.get("method",""),scope=r.get("scope",""),status="OK",error="",token_changed=old!=self.session_acct["refresh_token"],read_ok=True);self.set_line(outlook_backend.full_line(saved));self.current_message=r.get("message");self.render_outlook_message(self.current_message,r.get("method",""),self.session_acct["email"]);self.load_accounts();self.status("Đã đọc lại mail mới nhất.")
        except Exception as e:
            try:
                if self.session_acct:outlook_backend.upsert_account(self.session_acct,status="ERROR",error=str(e))
            except Exception:pass
            self.status(f"Đọc lại lỗi: {e}");self.load_accounts()

    def renew_current(self):
        try:
            self._require_backend();acct=self.session_acct or outlook_backend.parse_full_line(self._line());r=outlook_backend.renew_refresh_token(acct,(outlook_backend.get_account(acct["email"]) or {}).get("method",""))
            if not r.get("ok"):raise RuntimeError(r.get("message") or "Không làm mới được token")
            old=acct["refresh_token"];acct["refresh_token"]=r.get("refresh_token") or old;saved=outlook_backend.upsert_account(acct,status="OK",error="",scope=r.get("scope",""),token_changed=acct["refresh_token"]!=old);self.session_acct=acct;self.set_line(outlook_backend.full_line(saved));self.status("Đã làm mới refresh token." if acct["refresh_token"]!=old else "Token hợp lệ; Microsoft giữ nguyên refresh token.");self.load_accounts()
        except Exception as e:
            try:
                if 'acct' in locals() and acct:outlook_backend.upsert_account(acct,status="ERROR",error=str(e))
            except Exception:pass
            self.status(f"Làm mới token lỗi: {e}");self.load_accounts()

    def save_account(self):
        try:self._require_backend();acct=outlook_backend.parse_full_line(self._line());saved=outlook_backend.upsert_account(acct,status="SAVED",error="");self.set_line(outlook_backend.full_line(saved));self.status("Đã lưu tài khoản Outlook/Hotmail.");self.load_accounts()
        except Exception as e:self.status(f"Lưu tài khoản lỗi: {e}")

    def render_outlook_message(self,m,method,email_addr):
        if not m:return self.clear_message()
        bodyobj=m.get("body") or {};raw=bodyobj.get("content") or m.get("bodyPreview") or "";is_html=str(bodyobj.get("contentType") or "").lower()=="html";text=_html_to_text(raw) if is_html else str(raw);u,o,l=detect_shopee_data(m.get("subject",""),text,raw if is_html else "")
        frm=((m.get("from") or {}).get("emailAddress") or {});self.ui(lambda:(self.subj.set(m.get("subject") or "(Không tiêu đề)"),self.meta.set(f"{email_addr} • {method} • {frm.get('address') or frm.get('name') or ''} • {_fmt_time(m.get('receivedDateTime',''))}"),self.otp.set(o),self.username.set(u),self.link.set(l),self.body.delete("1.0","end"),self.body.insert("1.0",text)))

    def clear_message(self):self.current_message=None;self.subj.set("Chưa đọc mail");self.meta.set("");self.otp.set("");self.username.set("");self.link.set("");self.body.delete("1.0","end")

    def save_current_message(self):
        if not self.current_message or not self.session_acct:return self.status("Chưa có mail để lưu.")
        m=self.current_message;u,o,l=detect_shopee_data(m.get("subject",""),self.body.get("1.0","end-1c"),(m.get("body") or {}).get("content") or "");u=self.username.get().strip() or u
        try:row=outlook_backend.save_message({"account_email":self.session_acct["email"],"message":m,"shopee_username":u,"username_manual":bool(self.username.get().strip()),"shopee_otp":o,"verify_link":l});self.status("Đã lưu mail Outlook.");self.load_saved()
        except Exception as e:self.status(f"Lưu mail lỗi: {e}")

    def load_saved(self):
        if not outlook_backend or not hasattr(self,"saved_tree"):return
        q=self.saved_search.get().strip().lower() if hasattr(self,"saved_search") else "";rows=outlook_backend.list_messages();rows=[r for r in rows if not q or q in " ".join(str(r.get(k,"")) for k in ("account_email","subject","shopee_username","shopee_otp","verify_link")).lower()]
        if hasattr(self,"verify_only_var") and self.verify_only_var.get():rows=[r for r in rows if str(r.get("verify_link") or "").strip()]
        self._saved={str(r["id"]):r for r in rows};self.saved_tree.delete(*self.saved_tree.get_children())
        for r in rows:self.saved_tree.insert("","end",iid=str(r["id"]),values=(r["account_email"],r["shopee_username"],r["shopee_otp"],r["subject"],_fmt_time(r["received_at"])))

    def copy_selected_outlook_link(self):
        sel=self.saved_tree.selection() if hasattr(self,"saved_tree") else ()
        if not sel:return self.status("Chưa chọn mail có link xác thực.")
        row=self._saved.get(sel[0]);link=str((row or {}).get("verify_link") or "").strip()
        if not link:return self.status("Mail đã chọn chưa có link xác thực Shopee. Hãy Quét lại trước.")
        self.copy(link);self.status("Đã copy link xác thực Shopee.")

    def open_saved_outlook(self,_evt=None):
        sel=self.saved_tree.selection();
        if not sel:return
        r=self._saved.get(sel[0]);
        if not r:return
        self.nb.select(self.read_tab);self.subj.set(r["subject"]);self.meta.set(r["account_email"]+" • Mail đã lưu");self.username.set(r["shopee_username"]);self.otp.set(r["shopee_otp"]);self.link.set(r["verify_link"]);self.body.delete("1.0","end");self.body.insert("1.0",_html_to_text(r["body"]) if str(r["body_type"]).lower()=="html" else r["body"])

    def rescan_outlook(self):
        if not outlook_backend:return
        changed=0
        for r in outlook_backend.list_messages():
            if r.get("username_manual"):continue
            body=_html_to_text(r["body"]) if str(r.get("body_type","")).lower()=="html" else r.get("body","");u,o,l=detect_shopee_data(r.get("subject",""),body,r.get("body","") if str(r.get("body_type","")).lower()=="html" else "")
            if u!=r.get("shopee_username") or o!=r.get("shopee_otp") or l!=r.get("verify_link"):
                outlook_backend.update_message({"id":r["id"],"shopee_username":u,"username_manual":False,"shopee_otp":o,"verify_link":l});changed+=1
        self.load_saved();self.status(f"Quét lại Shopee: cập nhật {changed} mail.")

    def delete_saved_outlook(self):
        ids=[int(x) for x in self.saved_tree.selection()];
        if ids and messagebox.askyesno("Xóa mail",f"Xóa {len(ids)} mail đã chọn?"):n=outlook_backend.delete_messages(ids);self.load_saved();self.status(f"Đã xóa {n} mail.")

    def export_saved_outlook(self):
        rows=list(self._saved.values());self.export_csv("outlook_saved.csv",["email","username","otp","verify_link","subject","received_at","body"],[[r["account_email"],r["shopee_username"],r["shopee_otp"],r["verify_link"],r["subject"],r["received_at"],r["body"]] for r in rows])

    def load_accounts(self):
        if not outlook_backend or not hasattr(self,"account_tree"):return
        q=self.account_search.get().strip().lower() if hasattr(self,"account_search") else "";rows=outlook_backend.list_accounts();rows=[r for r in rows if (not q or q in " ".join(str(r.get(k,"")) for k in ("email","method","last_status","last_error")).lower()) and self._account_filter_ok(r)];self._accounts={r["email"]:r for r in rows};self.account_tree.delete(*self.account_tree.get_children());
        for r in rows:
            active=self._status_is_active(r.get("last_status"));tag="active" if active else "inactive"
            self.account_tree.insert("","end",iid=r["email"],values=("☑" if r["email"] in self._checked_accounts else "☐",r["email"],r["method"],r["last_status"],r.get("last_error", ""),_fmt_time(r["token_updated_at"]),_fmt_time(r["last_read_at"])),tags=(tag,))

    def use_outlook_account(self,_evt=None):
        sel=self.account_tree.selection();
        if not sel:return
        r=self._accounts.get(sel[0]);
        if not r:return
        self.full_line.delete("1.0","end");self.full_line.insert("1.0",r["full_line"]);self.nb.select(self.read_tab);self.worker(self.read_current,False)

    def delete_accounts_outlook(self):
        emails=list(self.account_tree.selection());
        if emails and messagebox.askyesno("Xóa tài khoản",f"Xóa {len(emails)} tài khoản đã chọn?"):n=outlook_backend.delete_accounts(emails);self.load_accounts();self.status(f"Đã xóa {n} tài khoản.")

    def export_accounts_outlook(self):
        rows=list(self._accounts.values());self.export_csv("outlook_accounts_fulltoken.csv",["email","method","status","token_updated_at","last_read_at","full_token"],[[r["email"],r["method"],r["last_status"],r["token_updated_at"],r["last_read_at"],r["full_line"]] for r in rows])

    def backup_outlook(self):
        if not outlook_backend:return
        path=filedialog.asksaveasfilename(title="Backup Outlook",defaultextension=".json",initialfile=f"outlook_backup_{datetime.now():%Y%m%d}.json",filetypes=[("JSON","*.json")]);
        if not path:return
        obj={"ok":True,"app_version":"native-v3.3","exported_at":_now_iso(),"accounts":outlook_backend.list_accounts(),"messages":outlook_backend.list_messages()};Path(path).write_text(json.dumps(obj,ensure_ascii=False,indent=2),encoding="utf-8");self.status("Đã backup Outlook JSON. File có token/password, hãy giữ kín.")

    def import_outlook(self):
        if not outlook_backend:return
        path=filedialog.askopenfilename(title="Import Outlook JSON",filetypes=[("JSON","*.json")]);
        if not path:return
        try:obj=json.loads(Path(path).read_text(encoding="utf-8-sig"));r=outlook_backend.import_backup(obj);self.load_accounts();self.load_saved();self.status(f"Import xong: {r['accounts']} tài khoản, {r['messages']} mail.")
        except Exception as e:self.status(f"Import lỗi: {e}")


class NativeMailSuiteFrame(tk.Frame):
    """Khung Đọc Mail native Tkinter. Có tìm nhanh email vừa AddMail và mở Inbox trực tiếp."""
    def __init__(self,parent):
        super().__init__(parent,bg="#f6f7f9")
        self.store=NativeMailStore()
        head=tk.Frame(self,bg="#f6f7f9")
        head.pack(fill="x",padx=18,pady=(16,6))
        tk.Label(head,text="Đọc Mail • Ryan Nguyễn",bg="#f6f7f9",fg="#222",font=("Segoe UI",18,"bold")).pack(side="left")
        tk.Label(head,text="Mail AddMail thành công được lưu tự động — tìm để mở/check Inbox",bg="#f6f7f9",fg="#777",font=("Segoe UI",9)).pack(side="left",padx=14,pady=(6,0))

        quick=tk.Frame(self,bg="#ffffff",highlightthickness=1,highlightbackground="#e4e7eb")
        quick.pack(fill="x",padx=18,pady=(2,6))
        tk.Label(quick,text="Email vừa liên kết Shopee",bg="#ffffff",fg="#333",font=("Segoe UI",9,"bold")).pack(side="left",padx=(10,6),pady=8)
        self.global_search_var=tk.StringVar()
        ent=ttk.Entry(quick,textvariable=self.global_search_var)
        ent.pack(side="left",fill="x",expand=True,padx=4,pady=6)
        ent.bind("<Return>",lambda _e:self.search_and_read())
        ttk.Button(quick,text="🔎 Tìm & đọc Inbox",command=self.search_and_read,style="Accent.TButton").pack(side="left",padx=4,pady=6)
        ttk.Button(quick,text="↻ Mail vừa thêm",command=self.load_latest_linked,style="Soft.TButton").pack(side="left",padx=(4,10),pady=6)

        self.global_status_var=tk.StringVar(value="Sẵn sàng. Email AddMail thành công sẽ tự xuất hiện tại đây.")
        tk.Label(self,textvariable=self.global_status_var,bg="#f6f7f9",fg="#666",font=("Segoe UI",8)).pack(fill="x",padx=22,pady=(0,3))

        self.notebook=ttk.Notebook(self)
        self.notebook.pack(fill="both",expand=True,padx=18,pady=(4,18))
        self.mailtm=MailTMTab(self.notebook,self.store)
        self.tinyhost=TinyHostTab(self.notebook,self.store)
        self.outlook=OutlookTab(self.notebook,self.store)
        self.notebook.add(self.mailtm,text="📮 Mail.tm")
        self.notebook.add(self.tinyhost,text="📬 TinyHost")
        self.notebook.add(self.outlook,text="✉ Outlook / Hotmail")
        self.load_latest_linked(silent=True)

    def register_linked_email(self,address: str,shopee_user: str="",source: str="shopee_addmail",meta: dict | None=None):
        try:
            row=self.store.save_linked_email(address,source=source,shopee_user=shopee_user,meta=meta)
            self.global_search_var.set(row["address"])
            self.global_status_var.set(f"✓ Đã tự lưu mail vừa AddMail: {row['address']}")
            return row
        except Exception as e:
            self.global_status_var.set(f"Không lưu được email liên kết: {e}")
            return None

    def load_latest_linked(self,silent: bool=False):
        rows=self.store.list_linked_emails()
        if rows:
            self.global_search_var.set(rows[0]["address"])
            if not silent:self.global_status_var.set(f"Mail AddMail gần nhất: {rows[0]['address']}")
        elif not silent:
            self.global_status_var.set("Chưa có email AddMail thành công được lưu.")

    def search_and_read(self):
        address=self.global_search_var.get().strip().lower()
        if not address:
            rows=self.store.list_linked_emails()
            if rows:
                address=rows[0]["address"]
                self.global_search_var.set(address)
        if not address or "@" not in address:
            self.global_status_var.set("Nhập email cần tìm hoặc bấm 'Mail vừa thêm'.")
            return
        resolved=self.store.find_mail_account(address)
        if not resolved:
            linked=self.store.get_linked_email(address)
            if linked:
                self.global_status_var.set(
                    f"Đã lưu liên kết {address}, nhưng chưa có credential Inbox trong Đọc Mail. "
                    "Hãy lưu/tạo tài khoản mail này ở Mail.tm, TinyHost hoặc Outlook trước."
                )
            else:
                self.global_status_var.set(f"Không tìm thấy {address} trong kho tài khoản Mail.tm / TinyHost / Outlook.")
            return
        provider=resolved["provider"]
        row=resolved["account"]
        self.global_status_var.set(f"Đã tìm thấy {address} ở {provider}. Đang kiểm tra Inbox...")
        if provider=="mailtm":
            password=str(row.get("password") or "")
            if not password:
                self.global_status_var.set(f"Mail.tm {address} đã lưu nhưng thiếu mật khẩu nên chưa thể đọc Inbox.")
                return
            self.notebook.select(self.mailtm)
            self.mailtm.user_var.set(address)
            self.mailtm.pass_var.set(password)
            try:self.mailtm.domain_var.set(address.split("@",1)[1])
            except Exception:pass
            self.mailtm.nb.select(self.mailtm.inbox_tab)
            self.mailtm.worker(self.mailtm.login_account)
        elif provider=="tinyhost":
            self.notebook.select(self.tinyhost)
            self.tinyhost.user_var.set(address)
            try:self.tinyhost.custom_domain_var.set(address.split("@",1)[1])
            except Exception:pass
            self.tinyhost.nb.select(self.tinyhost.inbox_tab)
            self.tinyhost.worker(self.tinyhost.select_address)
        elif provider=="outlook":
            full_line=str(row.get("full_line") or "")
            if not full_line:
                self.global_status_var.set(f"Outlook {address} đã lưu nhưng thiếu full token nên chưa thể đọc Inbox.")
                return
            self.notebook.select(self.outlook)
            self.outlook.full_line.delete("1.0","end")
            self.outlook.full_line.insert("1.0",full_line)
            self.outlook.nb.select(self.outlook.read_tab)
            self.outlook.worker(self.outlook.read_current,False)

    def stop_all(self):
        for tab in (self.mailtm,self.tinyhost,self.outlook):
            try:tab.stop_all()
            except Exception:pass
