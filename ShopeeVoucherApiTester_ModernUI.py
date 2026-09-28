"""Standalone GUI for multi-lane Shopee voucher-code API tests.
Hỗ trợ đa luồng Voucher, hàng loạt cookie/mã, proxy riêng từng luồng, kiểm tra Live/Die,
hẹn giờ chạy chính xác, cảnh báo chuông trước mốc 5 giây, TỰ ĐỘNG LƯU cấu hình và TÍCH HỢP VUBEL API vượt IVS.
"""

from __future__ import annotations

import json
import math
import os
import base64
import html as html_lib
import hashlib
import hmac
import platform
import uuid
import subprocess
import sys
import re
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Any, Callable
from datetime import datetime, timedelta, timezone
from urllib.parse import quote, parse_qs, urlparse

try:
    import httpx
except ModuleNotFoundError:
    httpx = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter
except ModuleNotFoundError:
    Workbook = None
    Font = None
    Alignment = None
    get_column_letter = None

try:
    from mail_native_suite import (
        NativeMailSuiteFrame,
        register_linked_email_global,
        detect_shopee_verification_link,
    )
except Exception:
    NativeMailSuiteFrame = None
    register_linked_email_global = None
    detect_shopee_verification_link = None

SAVE_BY_CODE_URL = "https://shopee.vn/api/v2/voucher_wallet/save_platform_voucher_by_voucher_code"
SAVE_VOUCHER_BY_LINK_URL = "https://mall.shopee.vn/api/v2/voucher_wallet/save_voucher"
ACCOUNT_INFO_URL = "https://mall.shopee.vn/api/v4/account/basic/get_account_info"
LOGIN_BY_PASSWORD_URL = "https://shopee.vn/api/v4/account/login_by_password"
VUBEL_LOGIN_URL = "https://api.vubel.store/v1/shopee/login"
VUBEL_ADDMAIL_URL = "https://api.vubel.store/v1/shopee/addmail"
VUBEL_EMAIL_GET_URL = "https://api.vubel.store/v1/email/get"
VUBEL_EMAIL_HISTORY_URL = "https://api.vubel.store/v1/email/history"
KIOTPROXY_CURRENT_URL = "https://api.kiotproxy.com/api/v1/proxies/current"
KIOTPROXY_NEW_URL = "https://api.kiotproxy.com/api/v1/proxies/new"
CONFIG_FILE = "config_minsu.json"
OWNER_NAME = "Ryan Nguyễn"
EDITION_NAME = "Ryan Nguyễn Native Mail 3 Tab Edition v3.20 SERVER LICENSE"
MAX_BATCH_CODES = 100
MAX_BATCH_COOKIES = 1000
# AddMail/MailFree không giới hạn độ trễ tối đa. Voucher có cấu hình riêng bên dưới.
MIN_DELAY_SECONDS = 0.5
DEFAULT_DELAY_SECONDS = 1.0
MIN_VOUCHER_DELAY_SECONDS = 1.0
MAX_VOUCHER_DELAY_SECONDS = 120.0
DEFAULT_VOUCHER_DELAY_SECONDS = 1.0
MIN_PROXY_ROTATE_SECONDS = 1.0
MAX_PROXY_ROTATE_SECONDS = 86400.0
DEFAULT_PROXY_ROTATE_SECONDS = 120.0
PROXY_PROTOCOL_HTTP = "HTTP"
PROXY_PROTOCOL_SOCKS5 = "SOCKS5"
PROXY_PROTOCOL_OPTIONS = (PROXY_PROTOCOL_HTTP, PROXY_PROTOCOL_SOCKS5)
MAX_VOUCHER_THREADS = 100
MAX_PROXY_ROTATE_ATTEMPTS = 3
PROXY_RETRY_DELAY_SECONDS = 5.0
EMAIL_ADDRESS_PATTERN = (
    r"[A-Za-z0-9][A-Za-z0-9._+-]*@"
    r"[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?"
    r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?)+"
)
EMAIL_ADDRESS_RE = re.compile(EMAIL_ADDRESS_PATTERN)

# ======================================================================
# BẢN QUYỀN / KÍCH HOẠT - xác minh qua server (Vercel)
# ======================================================================
# SECRET_SALT đã được xóa — không còn xác minh local.
# Mọi license check đều gọi LICENSE_SERVER_URL.
LICENSE_SERVER_URL = "https://ryan-voucher-license.vercel.app"
API_SERVER_URL     = "https://ryan-voucher-api.fly.dev"           # Thay bằng URL Fly.io thật
LICENSE_DIR_NAME   = "RyanNguyen_ShopeeVoucher"
LICENSE_FILE_NAME  = "license.json"
CURRENT_LICENSE_INFO: dict[str, Any] = {}
# RAM-only: không bao giờ ghi vào file config
_SESSION_TOKEN_RAM: str = ""
_VUBEL_KEY_RAM:     str = ""


def resource_path(filename: str) -> str:
    """Tìm file đi kèm khi chạy source hoặc Nuitka onefile."""
    candidates = []
    try:
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), filename))
    except Exception:
        pass
    try:
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), filename))
    except Exception:
        pass
    candidates.append(os.path.abspath(filename))
    for path in candidates:
        if path and os.path.exists(path):
            return path
    return candidates[0] if candidates else filename


def apply_ryan_window_icon(root: tk.Tk) -> None:
    """Áp dụng RYAN_VOUCHER.ico cho cửa sổ/taskbar nếu file có sẵn."""
    try:
        icon_path = resource_path("RYAN_VOUCHER.ico")
        if os.path.exists(icon_path):
            root.iconbitmap(default=icon_path)
    except Exception:
        pass


def get_license_file_path() -> str:
    """Trả về đường dẫn lưu license ở thư mục người dùng, không phụ thuộc thư mục chạy EXE."""
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or os.path.expanduser("~")
    folder = os.path.join(base, LICENSE_DIR_NAME)
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, LICENSE_FILE_NAME)


def _windows_machine_guid() -> str:
    if os.name != "nt":
        return ""
    try:
        import winreg
        access = winreg.KEY_READ
        # KEY_WOW64_64KEY giúp đọc đúng MachineGuid trên Windows 64-bit.
        if hasattr(winreg, "KEY_WOW64_64KEY"):
            access |= winreg.KEY_WOW64_64KEY
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Cryptography",
            0,
            access,
        ) as key:
            value, _ = winreg.QueryValueEx(key, "MachineGuid")
            return str(value or "").strip()
    except Exception:
        return ""


def get_machine_hwid() -> str:
    """Tạo HWID ổn định để khách gửi cho admin tạo key."""
    machine_guid = _windows_machine_guid()
    if machine_guid:
        raw = f"WIN|{machine_guid}"
    else:
        raw = "|".join([
            platform.system(),
            platform.node(),
            platform.machine(),
            str(uuid.getnode()),
        ])
    digest = hashlib.sha256(raw.encode("utf-8", errors="ignore")).hexdigest().upper()[:24]
    return "RYAN-" + "-".join(digest[i:i+4] for i in range(0, len(digest), 4))


def verify_license_server(key_text: str, hwid: str | None = None, action: str = "activate") -> dict[str, Any]:
    """Xác minh license qua API server (Vercel). action = 'activate' | 'verify'."""
    global _SESSION_TOKEN_RAM
    key_text = str(key_text or "").strip()
    hwid = (hwid or get_machine_hwid()).strip().upper()
    result: dict[str, Any] = {
        "valid": False,
        "message": "Không thể kết nối server license",
        "hwid": hwid,
        "expiry_timestamp": None,
        "expiry_text": "",
    }
    if not key_text:
        result["message"] = "Chưa có key kích hoạt"
        return result
    try:
        import urllib.request as _req
        import urllib.error as _uerr
        payload = json.dumps({
            "action": action,
            "key": key_text,
            "hwid": hwid,
            "bind": 1,
            "device_name": os.environ.get("COMPUTERNAME", "PC"),
        }).encode("utf-8")
        req = _req.Request(
            f"{LICENSE_SERVER_URL}/api/license",
            data=payload,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with _req.urlopen(req, timeout=15) as resp:
                data: dict = json.loads(resp.read().decode("utf-8"))
        except _uerr.HTTPError as http_exc:
            # Server trả 4xx/5xx kèm JSON {"ok":false,"message":...} — đọc message thật
            try:
                data = json.loads(http_exc.read().decode("utf-8"))
            except Exception:
                result["message"] = f"Lỗi server: HTTP {http_exc.code}"
                return result
    except Exception as exc:
        result["message"] = f"Lỗi kết nối server: {exc}"
        return result

    if not data.get("ok"):
        result["message"] = str(data.get("message") or "Server từ chối license")
        return result

    session_token = str(data.get("session_token") or "")
    if session_token:
        _SESSION_TOKEN_RAM = session_token

    record = data.get("record") or {}
    expiry_raw = data.get("expires") or record.get("expires") or ""
    expiry_timestamp = None
    exp_text = "Vĩnh viễn"
    if expiry_raw:
        try:
            if str(expiry_raw).isdigit():
                dt = datetime.fromtimestamp(int(expiry_raw))
            else:
                dt = datetime.fromisoformat(str(expiry_raw).replace("Z", "+00:00"))
            expiry_timestamp = int(dt.timestamp())
            exp_text = dt.strftime("%d/%m/%Y %H:%M:%S")
        except Exception:
            exp_text = str(expiry_raw)

    result.update({
        "valid": True,
        "message": str(data.get("message") or "Kích hoạt hợp lệ"),
        "expiry_timestamp": expiry_timestamp,
        "expiry_text": exp_text,
        "remaining_days": data.get("days_left", 0),
        "key": key_text,
        "session_token": session_token,
        "plan": data.get("plan", ""),
    })
    return result


def load_saved_license_key() -> str:
    try:
        path = get_license_file_path()
        if not os.path.exists(path):
            return ""
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return str(data.get("license_key") or "").strip()
    except Exception:
        return ""


def save_license_key(key_text: str, info: dict[str, Any]) -> None:
    path = get_license_file_path()
    payload = {
        "license_key": str(key_text or "").strip(),
        "hwid": str(info.get("hwid") or ""),
        "expiry_timestamp": info.get("expiry_timestamp"),
        "activated_at": int(time.time()),
        "owner": OWNER_NAME,
        "edition": EDITION_NAME,
        # session_token KHÔNG lưu file — chỉ giữ trong RAM (_SESSION_TOKEN_RAM)
    }
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _fetch_vubel_key_from_server() -> bool:
    """Lấy Vubel key từ API server sau khi đã có session_token, lưu vào RAM."""
    global _VUBEL_KEY_RAM
    if not _SESSION_TOKEN_RAM:
        return False
    try:
        import urllib.request as _req
        req = _req.Request(
            f"{API_SERVER_URL}/v1/mail/vubel-key",
            headers={"Authorization": f"Bearer {_SESSION_TOKEN_RAM}"},
            method="GET",
        )
        with _req.urlopen(req, timeout=10) as resp:
            data: dict = json.loads(resp.read().decode("utf-8"))
        key = str(data.get("key") or "").strip()
        if key:
            _VUBEL_KEY_RAM = key
            return True
    except Exception:
        pass
    return False


class LicenseActivationApp:
    """Cửa sổ kích hoạt trước khi mở tool chính."""
    def __init__(self, root: tk.Tk, initial_key: str = "", initial_message: str = ""):
        self.root = root
        self.hwid = get_machine_hwid()
        self.activated = False
        self.colors = {
            "bg": "#14151a",
            "panel": "#1d1f27",
            "input": "#282b36",
            "text": "#f5f6f7",
            "muted": "#9da4b0",
            "accent": "#ee4d2d",
            "success": "#35c46a",
            "danger": "#ff5f57",
        }

        root.title(f"Kích hoạt bản quyền • {OWNER_NAME}")
        apply_ryan_window_icon(root)
        root.geometry("650x470")
        root.resizable(False, False)
        root.configure(bg=self.colors["bg"])
        root.protocol("WM_DELETE_WINDOW", self._close)

        # Căn giữa màn hình.
        root.update_idletasks()
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        x = max(0, (sw - 650) // 2)
        y = max(0, (sh - 470) // 2)
        root.geometry(f"650x470+{x}+{y}")

        header = tk.Frame(root, bg=self.colors["panel"], height=78)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="RYAN NGUYỄN • LICENSE",
            bg=self.colors["panel"],
            fg=self.colors["accent"],
            font=("Segoe UI", 18, "bold"),
        ).pack(anchor="w", padx=28, pady=(15, 0))
        tk.Label(
            header,
            text="Shopee Voucher & SPC_ST Tool • v3.20",
            bg=self.colors["panel"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=29, pady=(1, 0))

        body = tk.Frame(root, bg=self.colors["bg"], padx=28, pady=22)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="Mã máy (HWID)", bg=self.colors["bg"], fg=self.colors["text"], font=("Segoe UI", 10, "bold")).pack(anchor="w")
        hwid_row = tk.Frame(body, bg=self.colors["bg"])
        hwid_row.pack(fill="x", pady=(6, 16))
        self.hwid_var = tk.StringVar(value=self.hwid)
        hwid_entry = tk.Entry(
            hwid_row, textvariable=self.hwid_var, state="readonly",
            readonlybackground=self.colors["input"], fg="#89b4fa",
            relief="flat", font=("Consolas", 11), bd=8,
        )
        hwid_entry.pack(side="left", fill="x", expand=True)
        tk.Button(
            hwid_row, text="COPY HWID", command=self.copy_hwid,
            bg="#343844", fg=self.colors["text"], activebackground="#414653",
            activeforeground="white", relief="flat", bd=0, padx=14, cursor="hand2",
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left", padx=(8, 0), ipady=7)

        tk.Label(body, text="Key kích hoạt", bg=self.colors["bg"], fg=self.colors["text"], font=("Segoe UI", 10, "bold")).pack(anchor="w")
        key_row = tk.Frame(body, bg=self.colors["bg"])
        key_row.pack(fill="x", pady=(6, 8))
        self.key_var = tk.StringVar(value=initial_key)
        self.key_entry = tk.Entry(
            key_row, textvariable=self.key_var,
            bg=self.colors["input"], fg="white", insertbackground="white",
            relief="flat", font=("Consolas", 10), bd=8,
        )
        self.key_entry.pack(side="left", fill="x", expand=True)
        tk.Button(
            key_row, text="PASTE", command=self.paste_key,
            bg="#343844", fg=self.colors["text"], activebackground="#414653",
            activeforeground="white", relief="flat", bd=0, padx=14, cursor="hand2",
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left", padx=(8, 0), ipady=7)

        self.status_var = tk.StringVar(value=initial_message or "Gửi HWID cho admin để nhận key, sau đó dán key vào đây.")
        self.status_label = tk.Label(
            body, textvariable=self.status_var, bg=self.colors["bg"], fg=self.colors["muted"],
            font=("Segoe UI", 9), anchor="w", justify="left", wraplength=580,
        )
        self.status_label.pack(fill="x", pady=(8, 14))

        self.activate_btn = tk.Button(
            body, text="KÍCH HOẠT & MỞ TOOL", command=self.activate,
            bg=self.colors["accent"], fg="white", activebackground="#d84325",
            activeforeground="white", relief="flat", bd=0, cursor="hand2",
            font=("Segoe UI", 11, "bold"), pady=11,
        )
        self.activate_btn.pack(fill="x")

        tk.Label(
            body,
            text="Key được khóa theo HWID và thời hạn. File license được lưu trong thư mục người dùng Windows.",
            bg=self.colors["bg"], fg="#6f7682", font=("Segoe UI", 8), wraplength=580, justify="left",
        ).pack(anchor="w", pady=(16, 0))

        self.key_entry.focus_set()
        root.bind("<Return>", lambda _e: self.activate())

    def copy_hwid(self):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.hwid)
            self.status_var.set("Đã copy HWID. Gửi mã này cho admin để tạo key.")
            self.status_label.config(fg=self.colors["success"])
        except Exception as exc:
            self.status_var.set(f"Không thể copy HWID: {exc}")
            self.status_label.config(fg=self.colors["danger"])

    def paste_key(self):
        try:
            value = self.root.clipboard_get().strip()
            self.key_var.set(value)
        except Exception:
            pass

    def activate(self):
        key_text = self.key_var.get().strip()
        self.status_var.set("Đang xác minh với server...")
        self.status_label.config(fg=self.colors["muted"])
        self.root.update_idletasks()
        info = verify_license_server(key_text, self.hwid, action="activate")
        if not info.get("valid"):
            self.status_var.set(str(info.get("message") or "Key không hợp lệ"))
            self.status_label.config(fg=self.colors["danger"])
            return
        try:
            save_license_key(key_text, info)
        except Exception as exc:
            self.status_var.set(f"Key hợp lệ nhưng không lưu được license: {exc}")
            self.status_label.config(fg=self.colors["danger"])
            return

        CURRENT_LICENSE_INFO.clear()
        CURRENT_LICENSE_INFO.update(info)
        _fetch_vubel_key_from_server()
        self.status_var.set(f"Kích hoạt thành công • Hết hạn: {info.get('expiry_text')}")
        self.status_label.config(fg=self.colors["success"])
        self.activated = True
        messagebox.showinfo(
            "Kích hoạt thành công",
            f"Bản quyền hợp lệ.\n\nHết hạn: {info.get('expiry_text')}",
            parent=self.root,
        )
        self.root.destroy()

    def _close(self):
        self.activated = False
        self.root.destroy()


def run_license_gate() -> bool:
    """Nếu license đã hợp lệ thì mở thẳng; nếu chưa thì hiện màn hình kích hoạt."""
    saved_key = load_saved_license_key()
    if saved_key:
        info = verify_license_server(saved_key, action="verify")
        if info.get("valid"):
            CURRENT_LICENSE_INFO.clear()
            CURRENT_LICENSE_INFO.update(info)
            _fetch_vubel_key_from_server()
            return True
        initial_message = str(info.get("message") or "License hiện tại không hợp lệ")
    else:
        initial_message = "Chưa kích hoạt. Copy HWID và gửi cho admin để tạo key."

    gate_root = tk.Tk()
    app = LicenseActivationApp(gate_root, initial_key=saved_key, initial_message=initial_message)
    gate_root.mainloop()
    return bool(app.activated)


def open_file_in_default_app(path: str) -> None:
    """Mở file bằng ứng dụng mặc định của hệ điều hành (Excel trên Windows nếu đã gán .xlsx)."""
    path = os.path.abspath(path)
    try:
        if os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as exc:
        raise RuntimeError(f"Đã lưu file nhưng không thể tự mở: {exc}") from exc

INVALID_MESSAGE_EXPLANATIONS = {
    1: "Voucher không hợp lệ",
    2: "Mã voucher không đúng",
    3: "Voucher đã hết hạn",
    4: "Voucher đã hết lượt",
    5: "Voucher không áp dụng cho tài khoản hoặc điều kiện hiện tại",
    6: "Voucher đã được sử dụng",
    7: "Chưa đạt giá trị đơn hàng tối thiểu",
    8: "Voucher chưa bắt đầu",
    9: "Voucher không áp dụng tại quốc gia này",
    10: "Chưa hoàn thành nhiệm vụ",
    14: "Thiết bị/App không đáp ứng điều kiện",
    16: "Không thể nhận voucher từ chính shop",
    17: "Phương thức thanh toán không phù hợp",
    18: "Chưa đến thời gian nhận voucher",
    19: "Kênh vận chuyển không phù hợp",
    21: "Không có phương thức thanh toán đáp ứng điều kiện",
    22: "Voucher chưa thể sử dụng ở thời điểm hiện tại",
    23: "Thanh toán đang bị hạn chế",
    24: "Nền tảng hiện tại không được hỗ trợ",
    25: "Voucher chỉ áp dụng trên Shopee Live",
    26: "Voucher chỉ áp dụng trên Shopee Video",
}

# Nhãn ngắn hiển thị ở cột "Kết quả" cho từng invalid_message_code cụ thể,
# thay vì gộp chung một nhãn "KHÔNG LƯU" mơ hồ. Quan trọng nhất là mã 6:
# "Voucher đã được sử dụng" — nghĩa là voucher đã bị TIÊU/DÙNG rồi, khác hẳn
# với trường hợp "đã có trong kho" (is_claimed_before, xem interpret_shopee_response)
# vốn là THÀNH CÔNG vì voucher vẫn còn nguyên trong ví, chỉ là đã lưu từ trước.
INVALID_MESSAGE_STATUS_LABELS = {
    1: "KHÔNG HỢP LỆ",
    2: "SAI MÃ",
    3: "HẾT HẠN",
    4: "HẾT LƯỢT",
    5: "KHÔNG ĐỦ ĐK",
    6: "ĐÃ SỬ DỤNG",
    7: "CHƯA ĐỦ ĐƠN",
    8: "CHƯA BẮT ĐẦU",
    9: "SAI QUỐC GIA",
    10: "CHƯA XONG NV",
    14: "SAI THIẾT BỊ",
    16: "TỪ CHÍNH SHOP",
    17: "SAI PTTT",
    18: "CHƯA ĐẾN GIỜ",
    19: "SAI VẬN CHUYỂN",
    21: "THIẾU PTTT",
    22: "CHƯA THỂ DÙNG",
    23: "TT BỊ HẠN CHẾ",
    24: "SAI NỀN TẢNG",
    25: "CHỈ TRÊN LIVE",
    26: "CHỈ TRÊN VIDEO",
}


def extract_invalid_message_code(response_data: Any) -> int | None:
    """Lấy invalid_message_code từ response JSON gốc của Shopee (nếu có)."""
    if not isinstance(response_data, dict):
        return None
    inner = response_data.get("data")
    inner = inner if isinstance(inner, dict) else {}
    code = inner.get("invalid_message_code")
    if code in (None, 0, "0"):
        return None
    try:
        return int(code)
    except (TypeError, ValueError):
        return None

@dataclass(frozen=True)
class TestResult:
    success: bool
    message: str
    status_code: int | None
    response_data: Any = None
    resolved_cookie: str = ""
    skipped: bool = False


@dataclass(frozen=True)
class VoucherThreadConfig:
    """Một luồng Voucher sở hữu một proxy/key và chu kỳ đổi IP riêng."""

    proxy_key: str
    rotate_seconds: float = DEFAULT_PROXY_ROTATE_SECONDS


def _parse_seconds(value: Any, label: str, minimum: float, maximum: float) -> float:
    try:
        seconds = float(str(value).strip().replace(",", "."))
    except (TypeError, ValueError):
        raise ValueError(f"{label} phải là một số") from None
    if not minimum <= seconds <= maximum:
        raise ValueError(f"{label} từ {minimum:g} đến {maximum:g} giây")
    return seconds


def parse_voucher_request_delay(value: Any) -> float:
    """Đọc độ trễ chỉ dành cho trang Voucher (1-120 giây)."""

    return _parse_seconds(
        value,
        "Độ trễ Voucher",
        MIN_VOUCHER_DELAY_SECONDS,
        MAX_VOUCHER_DELAY_SECONDS,
    )


def parse_proxy_rotate_seconds(value: Any) -> float | None:
    """Trống: dùng TTC nhà cung cấp; số dương: không giới hạn tối đa."""
    if value is None or not str(value).strip():
        return None
    try:
        seconds = float(str(value).strip().replace(",", "."))
    except (TypeError, ValueError, OverflowError):
        raise ValueError("Thời gian đổi IP phải là số giây lớn hơn 0 hoặc để trống") from None
    if not math.isfinite(seconds) or seconds <= 0:
        raise ValueError("Thời gian đổi IP phải là số hữu hạn lớn hơn 0")
    return seconds


def normalize_voucher_thread_configs(rows: Any) -> list[VoucherThreadConfig]:
    """Chuẩn hóa cấu hình lưu/UI; mỗi phần tử tương ứng đúng một worker."""

    if rows is None:
        rows = []
    if not isinstance(rows, (list, tuple)):
        raise ValueError("Cấu hình luồng Voucher phải là một danh sách")
    if len(rows) > MAX_VOUCHER_THREADS:
        raise ValueError(f"Tối đa {MAX_VOUCHER_THREADS} luồng Voucher")

    configs: list[VoucherThreadConfig] = []
    for index, row in enumerate(rows, 1):
        if isinstance(row, VoucherThreadConfig):
            proxy_key = row.proxy_key
            rotate_value = row.rotate_seconds
        elif isinstance(row, dict):
            proxy_key = row.get("proxy", row.get("proxy_key", ""))
            rotate_value = row.get("rotate_seconds", DEFAULT_PROXY_ROTATE_SECONDS)
        else:
            raise ValueError(f"Luồng {index}: cấu hình không hợp lệ")

        proxy_key = str(proxy_key or "").strip()
        if proxy_key.casefold() in {"direct", "trực tiếp", "tructiep"}:
            proxy_key = ""
        rotate_seconds = _parse_seconds(
            rotate_value,
            f"Luồng {index} - thời gian đổi IP API",
            MIN_PROXY_ROTATE_SECONDS,
            MAX_PROXY_ROTATE_SECONDS,
        )
        configs.append(VoucherThreadConfig(proxy_key, rotate_seconds))

    return configs or [VoucherThreadConfig("", DEFAULT_PROXY_ROTATE_SECONDS)]


def voucher_thread_configs_from_saved(config: Any) -> list[VoucherThreadConfig]:
    """Đọc cấu hình mới và nâng cấp proxy đơn của các bản cũ thành luồng 1."""

    config = config if isinstance(config, dict) else {}
    if "voucher_threads" in config:
        saved_rows = config.get("voucher_threads")
        if saved_rows == []:
            return [VoucherThreadConfig("", DEFAULT_PROXY_ROTATE_SECONDS)]
        if isinstance(saved_rows, (list, tuple)):
            valid_rows: list[VoucherThreadConfig] = []
            for row in saved_rows:
                try:
                    valid_rows.append(normalize_voucher_thread_configs([row])[0])
                except ValueError:
                    # Không để một dòng JSON hỏng làm mất các proxy/key hợp lệ còn lại.
                    continue
            if valid_rows:
                return valid_rows
    legacy_proxy = str(config.get("proxy") or "").strip()
    return normalize_voucher_thread_configs([{
        "proxy": legacy_proxy,
        "rotate_seconds": DEFAULT_PROXY_ROTATE_SECONDS,
    }])


def voucher_delay_from_saved(config: Any) -> float:
    """Đọc delay Voucher; giá trị legacy ngoài 1-120 được đưa về mặc định."""

    config = config if isinstance(config, dict) else {}
    raw_value = config.get("voucher_delay", config.get("delay", DEFAULT_VOUCHER_DELAY_SECONDS))
    try:
        return parse_voucher_request_delay(raw_value)
    except ValueError:
        return DEFAULT_VOUCHER_DELAY_SECONDS

def parse_text_list(raw_text: str, max_items: int, item_name: str) -> list[str]:
    items = []
    seen = set()
    for token in re.split(r"[\r\n,;]+", str(raw_text or "")):
        item = token.strip()
        if not item:
            continue
        key = item.casefold()
        if key in seen:
            continue
        seen.add(key)
        items.append(item)
        if len(items) > max_items:
            raise ValueError(f"Giới hạn tối đa {max_items} {item_name} mỗi lần chạy.")
    return items


def parse_line_list(raw_text: str, max_items: int, item_name: str) -> list[str]:
    """Mỗi dòng là một mục; không tách dấu ; vì Cookie Shopee dùng dấu chấm phẩy."""
    items: list[str] = []
    seen: set[str] = set()
    for token in str(raw_text or "").replace("\r", "").split("\n"):
        item = token.strip()
        if not item:
            continue
        key = item.casefold()
        if key in seen:
            continue
        seen.add(key)
        items.append(item)
        if len(items) > max_items:
            raise ValueError(f"Giới hạn tối đa {max_items} {item_name} mỗi lần chạy.")
    return items

def build_direct_code_payload(code: str) -> dict[str, str]:
    return {"voucher_code": code}


def parse_voucher_entry(raw_str: str) -> dict[str, Any]:
    """Phân tích 1 dòng nhập: có thể là mã voucher thường hoặc link voucher Shopee.
    Link được bóc tách promotionId/signature/evcode (evcode giải mã Base64 ra mã thực)."""
    raw = str(raw_str or "").strip()
    lower = raw.casefold()
    is_link = raw.startswith(("http://", "https://")) or any(
        marker in lower for marker in ("evcode=", "promotionid=", "promotion_id=", "signature=")
    )

    if not is_link:
        code = re.sub(r"[^A-Za-z0-9]", "", raw).upper()
        return {
            "raw": raw,
            "is_link": False,
            "code": code,
            "promotion_id": "",
            "signature": "",
            "is_valid": bool(code),
        }

    parsed_url = urlparse(raw)
    query_params = parse_qs(parsed_url.query or raw)
    lower_map: dict[str, list[str]] = {k.casefold(): v for k, v in query_params.items()}

    def get_param(*names: str) -> str:
        for name in names:
            values = lower_map.get(name.casefold())
            if values and values[0].strip():
                return values[0].strip()
        return ""

    evcode = get_param("evcode", "voucher_code", "code")
    code = ""
    if evcode:
        try:
            padding = len(evcode) % 4
            padded = evcode + ("=" * (4 - padding) if padding else "")
            decoded = base64.b64decode(padded).decode("utf-8")
            code = decoded.upper() if decoded.isalnum() else evcode.upper()
        except Exception:
            code = evcode.upper()

    promo_raw = get_param("promotionId", "promotionid", "promotion_id", "voucher_promotionid")
    promotion_id: Any = int(promo_raw) if promo_raw.isdigit() else promo_raw
    signature = get_param("signature", "sig")

    return {
        "raw": raw,
        "is_link": True,
        "code": code,
        "promotion_id": promotion_id,
        "signature": signature,
        "is_valid": bool(promotion_id and signature),
    }


def build_voucher_request(entry: dict[str, Any]) -> tuple[str, dict[str, Any]]:
    """Chọn URL + payload đúng theo loại voucher: theo link (promotionId/signature) hay theo mã."""
    if entry.get("is_link") and entry.get("promotion_id") and entry.get("signature"):
        return SAVE_VOUCHER_BY_LINK_URL, {
            "voucher_promotionid": entry["promotion_id"],
            "signature": str(entry["signature"]),
            "signature_source": "0",
        }
    return SAVE_BY_CODE_URL, build_direct_code_payload(str(entry.get("code") or ""))


def _voucher_entry_dedupe_key(entry: dict[str, Any], raw: str) -> str:
    if entry.get("is_link") and entry.get("promotion_id") and entry.get("signature"):
        return f"link:{entry['promotion_id']}:{entry['signature']}"
    if entry.get("code"):
        return f"code:{entry['code']}"
    return f"raw:{raw.casefold()}"


def dedupe_voucher_entries(raw_lines: list[str]) -> tuple[list[str], int]:
    """Loại các dòng trùng theo mã/promotionId+signature thực tế (không chỉ trùng y hệt chuỗi),
    để voucher đã được thêm vào danh sách rồi (dù gõ khác định dạng) không bị lặp lại."""
    kept: list[str] = []
    seen: set[str] = set()
    duplicate_count = 0
    for raw in raw_lines:
        entry = parse_voucher_entry(raw)
        key = _voucher_entry_dedupe_key(entry, raw)
        if key in seen:
            duplicate_count += 1
            continue
        seen.add(key)
        kept.append(raw)
    return kept, duplicate_count

def cookie_value(cookie: str, name: str) -> str:
    wanted = name.casefold()
    for part in str(cookie or "").split(";"):
        key, separator, value = part.partition("=")
        if separator and key.strip().casefold() == wanted:
            return value.strip()
    return ""

def build_shopee_headers(cookie: str) -> dict[str, str]:
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json;charset=UTF-8",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Origin": "https://shopee.vn",
        "Referer": "https://shopee.vn/",
        "X-API-Source": "pc",
        "X-Requested-With": "XMLHttpRequest",
        "Cookie": cookie,
    }
    csrf_token = cookie_value(cookie, "csrftoken")
    if csrf_token:
        headers["X-CSRFToken"] = csrf_token
    return headers

def parse_combo(combo_str: str) -> tuple[str, str, str]:
    """Phân tích chuỗi User|Pass|SPC_F (Hỗ trợ linh hoạt vị trí)"""
    parts = [p.strip() for p in combo_str.split('|')]
    spcf = username = password = ""
    for p in parts:
        if p.startswith('SPC_F=') or len(p) > 30:
            spcf = p.replace('SPC_F=', '')
        elif not username:
            username = p
        else:
            password = p
    return username, password, spcf

def extract_spcst_from_httpx(response: Any) -> str | None:
    """Trích xuất SPC_ST từ phản hồi của Vubel (Quét sâu)"""
    spc_st = response.cookies.get("SPC_ST")
    if spc_st: return spc_st
    
    for name, value in response.headers.multi_items():
        if name.lower() == 'set-cookie':
            match = re.search(r'SPC_ST=([^;]+)', value)
            if match: return match.group(1)
            
    try:
        data = response.json()
        raw = json.dumps(data)
        match = re.search(r'SPC_ST=([^;\"\'\s\,\}\]]+)', raw)
        if match: return match.group(1)
    except: pass
    
    return None


def normalize_proxy_protocol(protocol: Any) -> str:
    """Chuẩn hóa lựa chọn giao thức cho proxy dạng ``host:port``."""

    value = str(protocol or "").strip().casefold()
    value = value.replace("://", "").replace("-", "").replace("_", "")
    if value in {"socks", "socks5", "socks5h"}:
        return PROXY_PROTOCOL_SOCKS5
    return PROXY_PROTOCOL_HTTP


def normalize_proxy_url(
    proxy_url: str | None,
    default_protocol: Any = PROXY_PROTOCOL_HTTP,
) -> str:
    """Thêm scheme mặc định nhưng giữ nguyên scheme người dùng đã dán.

    KiotProxy thường trả ``host:port`` không kèm scheme. AddMail/MailFree cho
    phép chọn HTTP hoặc SOCKS5, vì vậy không thể mặc định mọi địa chỉ trần là
    HTTP như các bản trước.
    """

    proxy = str(proxy_url or "").strip()
    if not proxy:
        return ""
    if re.match(r"^(?:http|https|socks5|socks5h)://", proxy, re.IGNORECASE):
        return proxy
    scheme = "socks5" if normalize_proxy_protocol(default_protocol) == PROXY_PROTOCOL_SOCKS5 else "http"
    return f"{scheme}://{proxy}"


def infer_manual_proxy_protocol(proxy_value: str | None, default_protocol: Any = PROXY_PROTOCOL_HTTP) -> str:
    """Suy ra giao thức proxy thủ công từ scheme; nếu không có thì dùng lựa chọn UI."""

    raw = str(proxy_value or "").strip().lower()
    if raw.startswith(("socks://", "sock://", "socks5://", "socks5h://")):
        return PROXY_PROTOCOL_SOCKS5
    if raw.startswith(("http://", "https://")):
        return PROXY_PROTOCOL_HTTP
    return normalize_proxy_protocol(default_protocol)


def normalize_manual_proxy(
    proxy_value: str | None,
    default_protocol: Any = PROXY_PROTOCOL_HTTP,
) -> str:
    """Chuẩn hóa proxy thủ công cho AddMail/MailFree.

    Hỗ trợ host:port, user:pass@host:port, URL HTTP/SOCKS5 và dạng
    host:port:user:pass. Scheme sock:// hoặc socks:// được đổi thành socks5://.
    """

    raw = str(proxy_value or "").strip()
    if not raw:
        return ""
    if "\n" in raw or "\r" in raw:
        raw = next((line.strip() for line in raw.splitlines() if line.strip()), "")
    raw = re.sub(r"^(?:sock|socks)://", "socks5://", raw, flags=re.IGNORECASE)

    if "://" not in raw and "@" not in raw:
        m = re.fullmatch(r"([^:\s]+):(\d{2,5}):([^:\s]+):(.+)", raw)
        if m:
            host, port, user, password = m.groups()
            raw = f"{user}:{password}@{host}:{port}"

    protocol = infer_manual_proxy_protocol(raw, default_protocol)
    normalized = str(normalize_proxy_url(raw, protocol) or "").strip()
    no_scheme = re.sub(r"^[A-Za-z0-9+.-]+://", "", normalized)
    authority, sep, path = no_scheme.partition("/")
    hostpart = authority.rsplit("@", 1)[-1]
    is_static = bool(
        (not sep or not path.strip())
        and re.fullmatch(r"(?:\[[0-9A-Fa-f:]+\]|[^:]+):\d{2,5}", hostpart)
    )
    if not normalized or not is_static:
        raise ValueError(
            "Proxy thủ công không hợp lệ. Dùng IP:PORT, user:pass@IP:PORT, "
            "http://... hoặc socks5://..."
        )
    return normalized


def format_remaining_duration(seconds: int | float | None) -> str:
    """Định dạng bộ đếm hạn dùng ngắn gọn cho giao diện."""

    if seconds is None:
        return "chưa có dữ liệu"
    try:
        total = max(0, int(float(seconds)))
    except (TypeError, ValueError, OverflowError):
        return "chưa có dữ liệu"
    days, remainder = divmod(total, 86400)
    hours, remainder = divmod(remainder, 3600)
    minutes, secs = divmod(remainder, 60)
    if days:
        return f"{days} ngày {hours:02d} giờ {minutes:02d} phút {secs:02d} giây"
    if hours:
        return f"{hours} giờ {minutes:02d} phút {secs:02d} giây"
    if minutes:
        return f"{minutes} phút {secs:02d} giây"
    return f"{secs} giây"


def _normalized_field_name(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").casefold())


def _iter_response_dicts(value: Any, depth: int = 0):
    """Duyệt các dict lồng trong response KiotProxy mà không phụ thuộc schema."""

    if depth > 8:
        return
    if isinstance(value, dict):
        yield value
        for child in value.values():
            if isinstance(child, (dict, list)):
                yield from _iter_response_dicts(child, depth + 1)
    elif isinstance(value, list):
        for child in value[:100]:
            if isinstance(child, (dict, list)):
                yield from _iter_response_dicts(child, depth + 1)


def _coerce_timestamp(value: Any) -> int | None:
    """Đọc Unix timestamp (giây/ms) hoặc ISO-8601 từ response."""

    if isinstance(value, bool) or value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        try:
            number = float(value)
        except (TypeError, ValueError, OverflowError):
            return None
        if number <= 0:
            return None
        if number > 10**12:
            number /= 1000.0
        if number < 10**8:
            return None
        return int(number)

    text = str(value).strip()
    if not text:
        return None
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        return _coerce_timestamp(float(text))
    iso_text = text.replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(iso_text)
    except ValueError:
        parsed = None
    if parsed is None:
        for pattern in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S"):
            try:
                parsed = datetime.strptime(text, pattern)
                break
            except ValueError:
                continue
    if parsed is None:
        return None
    if parsed.tzinfo is None:
        return int(parsed.timestamp())
    return int(parsed.timestamp())


def _coerce_seconds(value: Any) -> float | None:
    if isinstance(value, bool) or value in (None, ""):
        return None
    try:
        return float(str(value).strip().replace(",", "."))
    except (TypeError, ValueError, OverflowError):
        return None


def extract_kiotproxy_key_expiry_timestamp(data: Any) -> int | None:
    """Tìm hạn key nếu KiotProxy trả trường tuyệt đối hoặc TTL của key.

    ``ttc``/``rotate_seconds`` là thời gian đổi IP, không phải hạn key nên
    được xử lý riêng và không bị dùng nhầm làm ngày hết hạn tài khoản.
    """

    absolute_keys = {
        "expireat", "expiresat", "expirydatetime", "expirationdatetime",
        "expiretime", "expireson", "expirytime", "expirationtime",
        "expire", "expiry", "expiration", "expiredate", "expirationdate",
        "expiredat", "expiredtime", "validuntil", "validto", "endtime",
        "endat", "deadline", "keyexpireat", "keyexpiresat", "keyexpiry",
        "licenseexpireat", "licenseexpiresat", "subscriptionexpireat",
    }
    duration_keys = {
        "keyttl", "keyremaining", "keyremainingseconds", "remainingkey",
        "keyvalidfor", "keyduration", "license_ttl", "license_remaining",
    }
    now = int(time.time())
    candidates: list[tuple[int, int]] = []
    for obj in _iter_response_dicts(data):
        for key, value in obj.items():
            normalized = _normalized_field_name(key)
            priority = 2 if any(marker in normalized for marker in ("key", "license", "subscription")) else 1
            if normalized in absolute_keys:
                timestamp = _coerce_timestamp(value)
                if timestamp is not None:
                    candidates.append((priority, timestamp))
            elif normalized in {_normalized_field_name(item) for item in duration_keys}:
                duration = _coerce_seconds(value)
                if duration is not None and duration >= 0:
                    candidates.append((priority, now + int(duration)))
    if not candidates:
        return None
    # Ưu tiên trường dành riêng cho key; nếu API chỉ có expire_at thì lấy giá trị
    # có thời hạn xa nhất để không nhầm TTL đổi IP ngắn với hạn tài khoản.
    best_priority = max(priority for priority, _ in candidates)
    return max(timestamp for priority, timestamp in candidates if priority == best_priority)


def extract_kiotproxy_proxy_info(data: Any, preferred_protocol: Any = PROXY_PROTOCOL_HTTP) -> dict[str, Any]:
    """Trích endpoint HTTP/SOCKS5 và TTL đổi IP từ response KiotProxy."""

    protocol = normalize_proxy_protocol(preferred_protocol)
    http_value = ""
    socks5_value = ""
    generic_value = ""
    ttc_seconds: float | None = None

    if isinstance(data, str):
        generic_value = data.strip()
    for obj in _iter_response_dicts(data):
        for key, value in obj.items():
            if not isinstance(value, (str, int, float)) or isinstance(value, bool):
                continue
            normalized = _normalized_field_name(key)
            text_value = str(value).strip()
            if not text_value:
                continue
            if normalized in {"http", "https", "httpproxy", "httpsproxy", "httpaddress", "httpendpoint"}:
                if not http_value:
                    http_value = text_value
            elif normalized in {"socks", "socks5", "socks5proxy", "socks5address", "socks5endpoint"}:
                if not socks5_value:
                    socks5_value = text_value
            elif normalized in {"proxy", "ipport", "ipaddressport", "hostport", "proxyaddress", "endpoint"}:
                if not generic_value:
                    generic_value = text_value
            elif normalized in {"ttc", "timetochanged", "timetorotate", "rotateseconds", "rotationseconds"}:
                if ttc_seconds is None:
                    candidate = _coerce_seconds(value)
                    if candidate is not None and candidate >= 0:
                        ttc_seconds = candidate

    http_value = http_value or generic_value
    socks5_value = socks5_value or generic_value
    selected_raw = socks5_value if protocol == PROXY_PROTOCOL_SOCKS5 else http_value
    if not selected_raw:
        selected_raw = http_value or socks5_value

    return {
        "http_proxy": normalize_proxy_url(http_value, PROXY_PROTOCOL_HTTP) if http_value else "",
        "socks5_proxy": normalize_proxy_url(socks5_value, PROXY_PROTOCOL_SOCKS5) if socks5_value else "",
        "selected_proxy": normalize_proxy_url(selected_raw, protocol) if selected_raw else "",
        "ttc_seconds": ttc_seconds,
        "key_expiry_timestamp": extract_kiotproxy_key_expiry_timestamp(data),
    }


def proxy_display_name(proxy_url: str | None) -> str:
    """Hiện host:port nhưng không lộ user/password proxy trong log/status."""

    raw = str(proxy_url or "").strip()
    if not raw:
        return "Direct"
    no_scheme = re.sub(r"^[A-Za-z0-9+.-]+://", "", raw)
    authority = no_scheme.split("/", 1)[0]
    return authority.rsplit("@", 1)[-1] or "Proxy"


def extract_emails_from_data(value: Any) -> list[str]:
    """Lấy mọi email mà response khai báo để phát hiện backend đổi mail."""
    email_keys = {
        "email", "mail", "address", "mailfree", "generatedemail",
        "addedemail", "actualemail",
    }
    found: list[str] = []
    seen: set[str] = set()

    def walk(obj: Any, depth: int = 0):
        if depth > 6:
            return
        if isinstance(obj, dict):
            for key, candidate in obj.items():
                normalized_key = re.sub(r"[^a-z0-9]", "", str(key).lower())
                if normalized_key in email_keys and isinstance(candidate, str) and "@" in candidate:
                    address = _extract_valid_email_address(candidate)
                    marker = address.casefold()
                    if address and marker not in seen:
                        seen.add(marker)
                        found.append(address)
                if isinstance(candidate, (dict, list)):
                    walk(candidate, depth + 1)
        elif isinstance(obj, list):
            for child in obj[:100]:
                walk(child, depth + 1)

    walk(value)
    return found


def extract_email_from_data(value: Any) -> str:
    emails = extract_emails_from_data(value)
    return emails[0] if emails else ""


def extract_mailfree_email_from_data(value: Any) -> str:
    """Lấy mail vừa tạo, không nhầm ``account.email`` của tài khoản Shopee."""
    priority = (
        "generatedemail", "generatedmail", "createdemail", "createdmail",
        "newemail", "newmail", "mailfree", "mailfreeemail",
        "addedemail", "actualemail", "resultemail",
    )
    found: dict[str, list[str]] = {key: [] for key in priority}
    generic: list[str] = []
    blocked_contexts = {
        "account", "shopeeaccount", "shopeeuser", "user", "profile",
        "request", "input", "credentials", "credential",
    }

    def add_unique(bucket: list[str], candidate: Any):
        if not isinstance(candidate, str) or "@" not in candidate:
            return
        address = _extract_valid_email_address(candidate)
        if address and address.casefold() not in {item.casefold() for item in bucket}:
            bucket.append(address)

    def walk(obj: Any, path: tuple[str, ...] = (), depth: int = 0):
        if depth > 6:
            return
        if isinstance(obj, dict):
            for key, candidate in obj.items():
                normalized_key = re.sub(r"[^a-z0-9]", "", str(key).lower())
                if normalized_key in found:
                    add_unique(found[normalized_key], candidate)
                elif normalized_key in {"email", "mail", "address"}:
                    if not any(part in blocked_contexts for part in path):
                        add_unique(generic, candidate)
                elif (
                    normalized_key in {"value", "result", "output"}
                    and any(part in {"mailfree", "mailbox", "inbox", "generatedmail"} for part in path)
                ):
                    add_unique(generic, candidate)
                if isinstance(candidate, (dict, list)):
                    walk(candidate, path + (normalized_key,), depth + 1)
        elif isinstance(obj, list):
            for child in obj[:100]:
                walk(child, path, depth + 1)

    walk(value)
    for key in priority:
        if found[key]:
            return found[key][0]
    return generic[0] if generic else ""


def extract_mailfree_password_from_data(value: Any, email: str = "") -> str:
    """Lấy mật khẩu inbox đi kèm mail vừa tạo, không lấy password Shopee."""
    target = _normalize_email_for_addmail(email).casefold()
    explicit_keys = {
        "mailpassword", "emailpassword", "mailboxpassword", "inboxpassword",
        "generatedpassword", "mailpass", "emailpass", "mailboxpass", "inboxpass",
    }
    blocked_contexts = {
        "account", "shopeeaccount", "shopeeuser", "user", "profile",
        "request", "input", "credentials", "credential",
    }
    explicit: list[str] = []
    paired: list[str] = []
    generic: list[str] = []

    def clean_secret(candidate: Any) -> str:
        return str(candidate or "").strip() if isinstance(candidate, (str, int, float)) else ""

    def dict_has_target(obj: dict[str, Any]) -> bool:
        for key, candidate in obj.items():
            normalized_key = re.sub(r"[^a-z0-9]", "", str(key).casefold())
            if normalized_key not in {
                "generatedemail", "generatedmail", "createdemail", "createdmail",
                "newemail", "newmail", "mailfree", "mailfreeemail",
                "addedemail", "actualemail", "resultemail", "email", "mail", "address",
            }:
                continue
            address = _extract_valid_email_address(candidate)
            if address and (not target or address.casefold() == target):
                return True
        return False

    def walk(obj: Any, path: tuple[str, ...] = (), depth: int = 0):
        if depth > 6:
            return
        if isinstance(obj, dict):
            has_mail = dict_has_target(obj)
            for key, candidate in obj.items():
                normalized_key = re.sub(r"[^a-z0-9]", "", str(key).casefold())
                secret = clean_secret(candidate)
                if normalized_key in explicit_keys and secret:
                    explicit.append(secret)
                elif (
                    normalized_key in {"password", "pass", "pwd"}
                    and secret
                    and not any(part in blocked_contexts for part in path)
                ):
                    if has_mail:
                        paired.append(secret)
                    else:
                        generic.append(secret)
                if isinstance(candidate, (dict, list)):
                    walk(candidate, path + (normalized_key,), depth + 1)
        elif isinstance(obj, list):
            for child in obj[:100]:
                walk(child, path, depth + 1)

    walk(value)
    if explicit:
        return explicit[0]
    if paired:
        return paired[0]
    generated = extract_mailfree_email_from_data(value)
    if generic and generated and (not target or generated.casefold() == target):
        return generic[0]
    return ""


def merge_mailfree_pairs_into_addmail(
    existing_accounts_text: str,
    existing_emails_text: str,
    new_pairs: list[tuple[str, str]],
) -> tuple[str, str, int]:
    """Nối cặp MailFree vào AddMail, giữ thẳng hàng và chống trùng cặp."""
    accounts = [line.strip() for line in str(existing_accounts_text or "").splitlines() if line.strip()]
    emails = [line.strip() for line in str(existing_emails_text or "").splitlines() if line.strip()]
    if len(accounts) != len(emails):
        raise ValueError(
            f"Hai ô AddMail đang lệch số dòng ({len(accounts)} tài khoản / {len(emails)} email). "
            "Hãy chỉnh thẳng hàng trước khi chuyển MailFree."
        )

    seen = {(account.casefold(), email.casefold()) for account, email in zip(accounts, emails)}
    added = 0
    for raw_account, raw_email in new_pairs:
        account = str(raw_account or "").strip()
        email = _normalize_email_for_addmail(raw_email)
        if not account:
            continue
        if not EMAIL_ADDRESS_RE.fullmatch(email):
            raise ValueError(f"MailFree không hợp lệ nên chưa chuyển: {raw_email!r}")
        marker = (account.casefold(), email.casefold())
        if marker in seen:
            continue
        accounts.append(account)
        emails.append(email)
        seen.add(marker)
        added += 1

    if len(accounts) > MAX_BATCH_COOKIES:
        raise ValueError(f"AddMail vượt giới hạn {MAX_BATCH_COOKIES} dòng sau khi chuyển.")
    return "\n".join(accounts), "\n".join(emails), added


def vubel_message(data: Any, default: str = "") -> str:
    if not isinstance(data, dict):
        return default
    message_keys = {"message", "msg", "error", "errormessage", "errormsg", "errordescription", "detail"}

    def message_from(obj: dict[str, Any]) -> str:
        for key, value in obj.items():
            normalized_key = re.sub(r"[^a-z0-9]", "", str(key).casefold())
            if normalized_key in message_keys and isinstance(value, str) and value.strip():
                return value.strip()
        return ""

    message = message_from(data)
    if message:
        return message
    inner = data.get("data")
    if isinstance(inner, dict):
        message = message_from(inner)
        if message:
            return message
    return default


def hash_shopee_password(password: str) -> str:
    """Băm mật khẩu chuẩn Shopee: SHA256(MD5(password))."""
    md5_hex = hashlib.md5(password.encode("utf-8")).hexdigest()
    return hashlib.sha256(md5_hex.encode("utf-8")).hexdigest()


def detect_shopee_account_locked(data: Any) -> tuple[bool, str]:
    """Nhận diện tài khoản bị khoá/bất thường từ response login Shopee (điển hình lỗi F02).

    Trả về (locked, message). Khi khoá thì không cần thử tiếp payload dự phòng và
    cũng không check được link voucher, nên caller dừng sớm và báo rõ cho người dùng.
    """
    if not isinstance(data, dict):
        return False, ""
    error_code = str(data.get("error") or data.get("error_code") or "").strip()
    # Ưu tiên câu mô tả đầy đủ (error_msg/message), tránh chỉ lấy mã lỗi "F02".
    message = ""
    for key in ("error_msg", "errormsg", "message", "msg", "detail", "error_description"):
        value = data.get(key)
        if isinstance(value, str) and value.strip():
            message = value.strip()
            break
    if not message:
        message = vubel_message(data, "")
    haystack = f"{error_code} {message}".casefold()
    locked_markers = (
        "f02",
        "bất thường", "bat thuong",
        "vi phạm", "vi pham",
        "không thể đăng nhập", "khong the dang nhap",
        "bị khoá", "bi khoa", "bị khóa", "bi khóa",
        "banned", "locked", "suspend", "restricted",
    )
    if any(marker in haystack for marker in locked_markers):
        return True, message
    return False, ""


def shopee_login_spcst(
    client: Any,
    combo: str,
    proxy_url: str | None = None,
    on_ivs_needed: Callable[[str, list, str | None], str | None] | None = None,
) -> dict[str, Any]:
    """Làm mới SPC_ST bằng cách gọi TRỰC TIẾP API login_by_password của Shopee.

    Không đi qua Vubel. Trả về cùng cấu trúc dict như ``vubel_login_spcst`` để
    thay thế trực tiếp tại các call-site. ``client`` là httpx client đã gắn proxy
    của lane nên login và lưu voucher đi chung một IP. ``on_ivs_needed`` giữ trong
    chữ ký cho tương thích; login trực tiếp chưa hỗ trợ luồng OTP/Captcha.
    """
    username, password, spcf = parse_combo(combo)
    if not username or not password or not spcf:
        return {"success": False, "message": "Sai định dạng. Cần User|Pass|SPC_F", "data": None}

    csrf = uuid.uuid4().hex
    password_hash = hash_shopee_password(password)
    headers = {
        "User-Agent": "Android app Shopee appver=30010 app_type=1",
        "Cookie": f"SPC_F={spcf}; csrftoken={csrf}",
        "Referer": "https://shopee.vn/",
        "x-csrftoken": csrf,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    base_payload = {
        "username": username,
        "password": password_hash,
        "client_identifier": {},
    }
    # Payload chính, sau đó 2 payload dự phòng theo cờ support_ivs.
    payloads = [
        dict(base_payload),
        {**base_payload, "support_ivs": False},
        {**base_payload, "support_ivs": True},
    ]

    last_data: Any = None
    last_status: int | None = None
    try:
        for payload in payloads:
            response = client.post(LOGIN_BY_PASSWORD_URL, json=payload, headers=headers)
            last_status = response.status_code
            try:
                last_data = response.json()
            except Exception:
                last_data = {"raw_text": str(response.text or "")[:3000]}

            spc_st = extract_spcst_from_httpx(response)
            if spc_st:
                return {
                    "success": True,
                    "message": "Đã lấy SPC_ST thành công (Shopee trực tiếp)",
                    "cookie": f"SPC_ST={spc_st}; SPC_F={spcf};",
                    "spc_st": spc_st,
                    "spc_f": spcf,
                    "username": username,
                    "data": last_data,
                    "status_code": last_status,
                }

            # Tài khoản bị khoá/bất thường (F02): dừng ngay, không thử payload khác.
            locked, lock_msg = detect_shopee_account_locked(last_data)
            if locked:
                detail = lock_msg or "Tài khoản có dấu hiệu bất thường / vi phạm chính sách Shopee"
                return {
                    "success": False,
                    "locked": True,
                    "message": f"Tài khoản bị khoá: {detail}",
                    "data": last_data,
                    "status_code": last_status,
                }

        return {
            "success": False,
            "message": f"Shopee không trả SPC_ST: {vubel_message(last_data, 'Đăng nhập thất bại')}",
            "data": last_data,
            "status_code": last_status,
        }
    except Exception as exc:
        return {"success": False, "message": f"Lỗi gọi API Shopee login: {exc}", "data": None}


def vubel_login_spcst(
    client: Any,
    combo: str,
    vubel_key: str,
    proxy_url: str | None,
    on_ivs_needed: Callable[[str, list, str | None], str | None] | None = None,
) -> dict[str, Any]:
    """Đăng nhập Vubel từ User|Pass|SPC_F và trả SPC_ST/full cookie.

    v3.14: nhận diện flow xác minh theo *nội dung response*, không chỉ dựa vào
    hai cờ ``need_ivs`` / ``need_verify``. Một số response thực tế trả
    ``status=input_required`` + ``sessionId`` + ``methods`` kèm thông báo
    OTP/Captcha, nên bản cũ bỏ sót và dừng trước khi mở cửa sổ xác minh.
    """
    username, password, spcf = parse_combo(combo)
    if not username or not password or not spcf:
        return {"success": False, "message": "Sai định dạng. Cần User|Pass|SPC_F", "data": None}
    if not vubel_key:
        return {"success": False, "message": "Thiếu Vubel API Key", "data": None}

    proxy = normalize_proxy_url(proxy_url)
    payload = {"username": username, "password": password, "spc_f": spcf, "proxy": proxy}
    headers = {"x-api-key": vubel_key, "Content-Type": "application/json", "Accept": "application/json"}

    def _obj_data(obj: Any) -> dict[str, Any]:
        return obj if isinstance(obj, dict) else {}

    def _verify_context(data: Any) -> tuple[str | None, list, bool, str]:
        top = _obj_data(data)
        inner = _obj_data(top.get("data"))
        sid = (
            inner.get("sessionId") or inner.get("session_id")
            or top.get("sessionId") or top.get("session_id")
        )
        methods = inner.get("methods") or top.get("methods") or []
        if not isinstance(methods, list):
            methods = []

        status = str(inner.get("status") or top.get("status") or "").strip().lower()
        msg = " ".join(
            str(x) for x in (
                top.get("message"), top.get("msg"), top.get("detail"), top.get("error_msg"),
                inner.get("message"), inner.get("msg"), inner.get("detail"), inner.get("error_msg"),
            ) if x not in (None, "")
        ).strip()
        hint_text = f"{status} {msg}".lower()
        explicit_flag = bool(
            top.get("need_ivs") or top.get("need_verify") or top.get("need_verification")
            or inner.get("need_ivs") or inner.get("need_verify") or inner.get("need_verification")
        )
        verify_words = (
            "input_required", "verify", "verification", "ivs", "otp", "captcha",
            "xác minh", "xac minh", "duyệt", "duyet", "security verification",
        )
        needs_verify = explicit_flag or (bool(sid) and any(w in hint_text for w in verify_words))
        return (str(sid) if sid else None), methods, needs_verify, msg

    try:
        response = client.post(VUBEL_LOGIN_URL, json=payload, headers=headers)
        try:
            data = response.json()
        except Exception:
            data = {"raw_text": str(response.text or "")[:3000]}

        spc_st = extract_spcst_from_httpx(response)
        sid, methods, needs_verify, verify_msg = _verify_context(data)

        if not spc_st and needs_verify:
            if not sid:
                return {
                    "success": False,
                    "message": "Vubel yêu cầu OTP/Captcha nhưng response không có sessionId. "
                               + (verify_msg or vubel_message(data, "")),
                    "data": data,
                    "status_code": response.status_code,
                }
            if not on_ivs_needed:
                return {
                    "success": False,
                    "message": "Vubel yêu cầu xác minh OTP/Captcha nhưng tool chưa có callback xác minh.",
                    "data": data,
                    "status_code": response.status_code,
                }

            spc_st = on_ivs_needed(sid, methods, proxy_url)
            if not spc_st:
                return {
                    "success": False,
                    "message": "Đã mở xác minh nhưng chưa lấy được SPC_ST mới (có thể chưa duyệt hoặc đã đóng cửa sổ).",
                    "data": data,
                    "status_code": response.status_code,
                }

        if spc_st:
            return {
                "success": True,
                "message": "Đã lấy SPC_ST thành công",
                "cookie": f"SPC_ST={spc_st}; SPC_F={spcf};",
                "spc_st": spc_st,
                "spc_f": spcf,
                "username": username,
                "data": data,
                "status_code": response.status_code,
            }

        return {
            "success": False,
            "message": f"Vubel báo lỗi: {vubel_message(data, 'Không có SPC_ST trả về')}",
            "data": data,
            "status_code": response.status_code,
        }
    except Exception as exc:
        return {"success": False, "message": f"Lỗi gọi Vubel API: {exc}", "data": None}

def _strip_command_prefix(value: str) -> str:
    """Bỏ /addmail hoặc /mailfree2 nếu người dùng dán cả lệnh vào ô tài khoản."""
    text = str(value or "").strip()
    text = re.sub(r"^\s*/(?:addmail|mailfree2)\s+", "", text, flags=re.I)
    return text.strip()


def _extract_named_token(text: str, *names: str) -> str:
    """Lấy giá trị NAME=... từ cookie/combo, chấp nhận cả dấu ; và |."""
    source = str(text or "")
    for name in names:
        m = re.search(rf"(?i)(?:^|[;|,\s]){re.escape(name)}\s*=\s*([^;|,\s]+)", source)
        if m:
            return m.group(1).strip()
    return ""


def _parse_addmail_account(account: str) -> dict[str, str]:
    """Tự nhận dạng dữ liệu dán vào ô Tài khoản/Cookie.

    Hỗ trợ:
      - user|pass|SPC_F=xxx
      - SPC_F=xxx|user|pass
      - SPC_ST=xxx; SPC_F=yyy
      - SPC_ST=xxx
      - SPCST=xxx (alias người dùng hay dán)
      - cả chuỗi đã có tiền tố /addmail hoặc /mailfree2
    """
    raw = _strip_command_prefix(account)
    spc_st = _extract_named_token(raw, "SPC_ST", "SPCST")
    spc_f = _extract_named_token(raw, "SPC_F", "SPCF")

    # Cookie/session ưu tiên hơn combo vì không cần user/pass để xác thực request.
    if spc_st:
        cookie_header = raw
        # Nếu người dùng chỉ dán SPCST=..., chuẩn hóa thành tên cookie chuẩn.
        if ";" not in cookie_header and "|" not in cookie_header:
            cookie_header = f"SPC_ST={spc_st}"
            if spc_f:
                cookie_header += f"; SPC_F={spc_f}"
        else:
            cookie_header = re.sub(r"(?i)\bSPCST\s*=", "SPC_ST=", cookie_header)
            cookie_header = re.sub(r"(?i)\bSPCF\s*=", "SPC_F=", cookie_header)
        return {
            "kind": "spc_st",
            "raw": raw,
            "username": "",
            "password": "",
            "spc_f": spc_f,
            "spc_st": spc_st,
            "cookie": cookie_header,
        }

    # Combo SPC_F: tìm token SPC_F rồi giữ 2 cột còn lại làm user/pass.
    parts = [x.strip() for x in raw.split("|")]
    cleaned: list[str] = []
    for part in parts:
        if not part:
            cleaned.append("")
            continue
        if re.match(r"(?i)^SPC_?F\s*=", part):
            if not spc_f:
                spc_f = part.split("=", 1)[1].strip() if "=" in part else ""
            continue
        if re.match(r"(?i)^SPC_?ST\s*=", part) or re.match(r"(?i)^SPCST\s*=", part):
            continue
        cleaned.append(part)

    nonempty = [x for x in cleaned if x]
    username = nonempty[0] if len(nonempty) >= 1 else ""
    password = nonempty[1] if len(nonempty) >= 2 else ""

    # Tương thích parse_combo cũ nếu chuỗi lạ nhưng vẫn có đủ dữ liệu.
    if spc_f and (not username or not password):
        u2, p2, f2 = parse_combo(raw)
        username = username or u2
        password = password or p2
        spc_f = spc_f or f2

    if spc_f:
        return {
            "kind": "spc_f",
            "raw": raw,
            "username": username,
            "password": password,
            "spc_f": spc_f,
            "spc_st": "",
            "cookie": "",
        }

    return {
        "kind": "unknown",
        "raw": raw,
        "username": "",
        "password": "",
        "spc_f": "",
        "spc_st": "",
        "cookie": "",
    }


def _normalize_email_for_addmail(email: str) -> str:
    value = str(email or "")
    for hidden in ("\ufeff", "\u200b", "\u200c", "\u200d"):
        value = value.replace(hidden, "")
    value = value.strip()
    if value.lower().startswith("mailto:"):
        value = value[7:].strip()
    value = value.strip("<> \t\r\n")
    return value


def _extract_valid_email_address(value: Any) -> str:
    """Tách đúng một địa chỉ email hợp lệ khỏi giá trị response có chú thích."""
    cleaned = _normalize_email_for_addmail(str(value or ""))
    match = EMAIL_ADDRESS_RE.search(cleaned)
    return match.group(0) if match else ""


def _build_addmail_argument(email: str, parsed: dict[str, str]) -> str:
    """Tạo đúng phần sau lệnh /addmail theo cú pháp của tool converter cũ."""
    email = _normalize_email_for_addmail(email)
    if not email or "@" not in email:
        return ""
    if parsed.get("kind") == "spc_st":
        return f"{email}|SPC_ST={parsed.get('spc_st', '')}"
    if parsed.get("kind") == "spc_f":
        return (
            f"{email}|{parsed.get('username', '')}|{parsed.get('password', '')}"
            f"|SPC_F={parsed.get('spc_f', '')}"
        )
    return ""


def _decode_vubel_response(response: Any) -> Any:
    try:
        return response.json()
    except Exception:
        return {"raw_text": str(getattr(response, "text", "") or "")[:4000]}


def _vubel_response_text(data: Any) -> str:
    values: list[str] = []
    interesting = {
        "message", "msg", "error", "errors", "errormessage", "errormsg",
        "errordescription", "detail", "status", "code", "errorcode", "rawtext",
    }

    def walk(obj: Any, depth: int = 0):
        if depth > 4:
            return
        if isinstance(obj, dict):
            for key, candidate in obj.items():
                normalized = re.sub(r"[^a-z0-9]", "", str(key).casefold())
                if normalized in interesting and candidate not in (None, "", [], {}):
                    values.append(str(candidate))
                elif isinstance(candidate, (dict, list)):
                    walk(candidate, depth + 1)
        elif isinstance(obj, list):
            for candidate in obj[:30]:
                walk(candidate, depth + 1)

    walk(data)
    return " ".join(values).strip()


def _vubel_required_input(data: Any) -> str:
    text = _vubel_response_text(data)
    match = re.search(r"(?:missing|invalid)\s+['\"]?([A-Za-z0-9_-]+)['\"]?\s+input", text, re.I)
    if match:
        return match.group(1).strip().casefold()

    def walk(obj: Any, depth: int = 0) -> str:
        if depth > 5:
            return ""
        if isinstance(obj, dict):
            for key in ("required_input", "requiredInput", "field"):
                candidate = obj.get(key)
                if isinstance(candidate, str) and candidate.strip():
                    return candidate.strip().casefold()
            for candidate in obj.values():
                if isinstance(candidate, (dict, list)):
                    found = walk(candidate, depth + 1)
                    if found:
                        return found
        elif isinstance(obj, list):
            for candidate in obj[:30]:
                found = walk(candidate, depth + 1)
                if found:
                    return found
        return ""

    return walk(data)


def _vubel_is_input_required(data: Any) -> bool:
    text = _vubel_response_text(data).casefold()
    if "input_required" in text or "input required" in text:
        return True
    if "missing " in text and " input" in text:
        return True

    def walk(obj: Any, depth: int = 0) -> bool:
        if depth > 5:
            return False
        if isinstance(obj, dict):
            if str(obj.get("status") or "").strip().casefold() == "input_required":
                return True
            return any(walk(candidate, depth + 1) for candidate in obj.values() if isinstance(candidate, (dict, list)))
        if isinstance(obj, list):
            return any(walk(candidate, depth + 1) for candidate in obj[:30])
        return False

    return walk(data)


def _vubel_continuation_state(data: Any) -> dict[str, Any]:
    wanted = {
        "sessionId", "session_id", "requestId", "request_id", "taskId", "task_id",
        "requestState", "request_state", "flowId", "flow_id", "stateId", "state_id",
        "tokenId", "token_id",
    }
    found: dict[str, Any] = {}

    def walk(obj: Any, depth: int = 0):
        if depth > 5:
            return
        if isinstance(obj, dict):
            for key, candidate in obj.items():
                if key in wanted and candidate not in (None, "", [], {}):
                    found[key] = candidate
                elif isinstance(candidate, (dict, list)):
                    walk(candidate, depth + 1)
        elif isinstance(obj, list):
            for candidate in obj[:30]:
                walk(candidate, depth + 1)

    walk(data)
    return found


def _mailfree_result_data(
    data: Any,
    *,
    flow: list[str],
    email: str,
    proxy: str,
) -> dict[str, Any]:
    output = dict(data) if isinstance(data, dict) else {"response": data}
    debug = output.setdefault("_client_debug", {})
    if not isinstance(debug, dict):
        debug = {}
        output["_client_debug"] = debug
    debug.update({
        "mode": "mailfree_generate_only",
        "flow": list(flow),
        "actual_email": email,
        "proxy": proxy,
        "request_contract": "mode=random+mailfree=random",
        "linked_to_shopee": False,
        "mail_password_available": bool(extract_mailfree_password_from_data(data, email)),
    })
    return output


def vubel_generate_mailfree(
    client: Any,
    vubel_key: str,
    proxy_url: str | None,
) -> tuple[TestResult, str]:
    """Chỉ tạo MailFree; tuyệt đối không đăng nhập hay AddMail vào Shopee."""
    if not vubel_key:
        return TestResult(False, "Thiếu Vubel API Key", None), ""

    proxy = normalize_proxy_url(proxy_url)
    headers = {
        "x-api-key": vubel_key,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    base_payload: dict[str, Any] = {
        "proxy": proxy,
        "mode": "random",
        "mailfree": "random",
    }
    flow: list[str] = []
    field_attempts = 0
    flow_state: dict[str, Any] = {}
    actual_email = ""

    def finish(
        success: bool,
        message: str,
        response: Any,
        data: Any,
    ) -> tuple[TestResult, str]:
        status_code = getattr(response, "status_code", None) if response is not None else None
        payload = _mailfree_result_data(data, flow=flow, email=actual_email, proxy=proxy)
        return TestResult(success, message, status_code, payload), actual_email

    try:
        response = client.post(VUBEL_ADDMAIL_URL, headers=headers, json=dict(base_payload))
        data = _decode_vubel_response(response)
        flow.append("MAILFREE:CREATE(json)")

        for _round in range(4):
            detected = extract_mailfree_email_from_data(data)
            if detected:
                actual_email = _normalize_email_for_addmail(detected)

            status_code = int(getattr(response, "status_code", 0) or 0)
            # Không gửi email đầu vào nên một trường generated_email hợp lệ là
            # bằng chứng mail đã được tạo. Dừng tại đây kể cả response đồng thời
            # mang HTTP lỗi/cookie error của bước AddMail phía sau.
            if actual_email:
                return finish(
                    True,
                    f"Đã tạo MailFree: {actual_email} | Chưa AddMail; không kiểm tra cookie Shopee.",
                    response,
                    data,
                )
            input_required = _vubel_is_input_required(data)
            if not input_required:
                if not 200 <= status_code < 300:
                    server = vubel_message(data, _vubel_response_text(data) or f"HTTP {status_code}")
                    return finish(False, f"Tạo MailFree thất bại: {server}", response, data)
                server = vubel_message(data, _vubel_response_text(data) or "API không trả địa chỉ MailFree")
                return finish(
                    False,
                    f"Vubel không trả địa chỉ MailFree: {server}",
                    response,
                    data,
                )

            flow_state.update(_vubel_continuation_state(data))
            field = _vubel_required_input(data)
            if field == "addmail":
                return finish(
                    False,
                    "API chưa trả địa chỉ MailFree và đang yêu cầu AddMail; tab MailFree không chạy AddMail.",
                    response,
                    data,
                )
            if not field:
                field = "mailfree"
            if field != "mailfree":
                return finish(
                    False,
                    f"API yêu cầu input không thuộc luồng tạo MailFree: {field}",
                    response,
                    data,
                )
            if field_attempts >= 2:
                server = _vubel_response_text(data)[:260]
                suffix = f" | Server: {server}" if server else ""
                return finish(
                    False,
                    f"API lặp lại input MailFree không thể hoàn tất{suffix}",
                    response,
                    data,
                )

            body = dict(base_payload)
            body.update(flow_state)
            body.update({
                "field": "mailfree",
                "required_input": "mailfree",
                "input": "random",
                "mailfree": "random",
            })
            if field_attempts == 0:
                response = client.post(VUBEL_ADDMAIL_URL, headers=headers, json=body)
                flow.append("MAILFREE:mailfree=random(json)")
            else:
                form_headers = {
                    key: value for key, value in headers.items()
                    if key.casefold() != "content-type"
                }
                response = client.post(VUBEL_ADDMAIL_URL, headers=form_headers, data=body)
                flow.append("MAILFREE:mailfree=random(form)")
            field_attempts += 1
            data = _decode_vubel_response(response)

        return finish(False, "Luồng tạo MailFree vượt số vòng cho phép.", response, data)
    except Exception as exc:
        return TestResult(
            False,
            f"Lỗi tạo MailFree qua Vubel: {exc} | Flow: {' → '.join(flow) or 'chưa gửi'}",
            None,
        ), actual_email


def extract_vubel_inbox_messages(data: Any) -> list[dict[str, Any]]:
    """Lấy danh sách thư từ response ``/v1/email/get`` với schema linh hoạt."""
    if not isinstance(data, dict):
        return []
    roots: list[Any] = [data]
    nested = data.get("data")
    if isinstance(nested, dict):
        roots.insert(0, nested)
    for root in roots:
        if not isinstance(root, dict):
            continue
        for key in ("emails", "messages", "items"):
            candidate = root.get(key)
            if isinstance(candidate, list):
                return [item if isinstance(item, dict) else {"content": item} for item in candidate]
            if isinstance(candidate, dict):
                nested_items = candidate.get("items") or candidate.get("data")
                if isinstance(nested_items, list):
                    return [item if isinstance(item, dict) else {"content": item} for item in nested_items]
                values = list(candidate.values())
                if values and all(isinstance(item, dict) for item in values):
                    return values
    return []


def _vubel_has_explicit_failure(data: Any) -> bool:
    failure_keys = {
        "error", "errors", "errormessage", "errormsg", "errordescription",
        "errorcode", "exception",
    }

    def walk(obj: Any, depth: int = 0) -> bool:
        if depth > 6:
            return False
        if isinstance(obj, dict):
            if obj.get("ok") is False or obj.get("success") is False:
                return True
            status = str(obj.get("status") or "").strip().casefold()
            if status in {"failed", "fail", "error", "input_required"}:
                return True
            for key, candidate in obj.items():
                normalized = re.sub(r"[^a-z0-9]", "", str(key).casefold())
                if normalized in failure_keys and candidate not in (None, "", 0, "0", False, [], {}):
                    return True
                if isinstance(candidate, (dict, list)) and walk(candidate, depth + 1):
                    return True
        elif isinstance(obj, list):
            return any(walk(candidate, depth + 1) for candidate in obj[:100])
        return False

    return walk(data)


def mailfree_message_summary(item: dict[str, Any]) -> tuple[str, str, str]:
    """Chuẩn hóa thời gian/người gửi/tiêu đề để hiển thị inbox Vubel."""
    def first(keys: tuple[str, ...], default: str = "") -> str:
        for key in keys:
            candidate = item.get(key)
            if candidate not in (None, "", [], {}):
                if isinstance(candidate, dict):
                    candidate = (
                        candidate.get("address") or candidate.get("email")
                        or candidate.get("name") or candidate
                    )
                return str(candidate).strip()
        return default

    received = first(("receivedAt", "received_at", "createdAt", "created_at", "date", "timestamp", "time"), "—")
    sender = first(("from", "sender", "fromAddress", "from_address", "mailFrom", "mail_from"), "—")
    subject = first(("subject", "title", "name"), "(Không có tiêu đề)")
    return received, sender, subject


def format_mailfree_message(item: dict[str, Any]) -> str:
    """Hiển thị nội dung thư dạng text; fallback JSON nếu API đổi schema."""
    received, sender, subject = mailfree_message_summary(item)
    body: Any = ""
    for key in ("text", "body", "content", "html", "message", "bodyText", "bodyHtml"):
        candidate = item.get(key)
        if candidate not in (None, "", [], {}):
            body = candidate
            break
    if isinstance(body, dict):
        body = body.get("text") or body.get("html") or body.get("content") or body
    if isinstance(body, (dict, list)):
        body_text = json.dumps(body, ensure_ascii=False, indent=2)
    else:
        body_text = str(body or "").strip()
    if "<" in body_text and ">" in body_text:
        body_text = re.sub(r"(?i)<br\s*/?>", "\n", body_text)
        body_text = re.sub(r"(?i)</p\s*>", "\n", body_text)
        body_text = re.sub(r"<[^>]+>", "", body_text)
        body_text = html_lib.unescape(body_text)
    if not body_text:
        body_text = json.dumps(item, ensure_ascii=False, indent=2)
    return (
        f"Từ: {sender}\n"
        f"Thời gian: {received}\n"
        f"Tiêu đề: {subject}\n"
        f"{'─' * 72}\n{body_text}"
    )


def vubel_read_mailfree_inbox(
    client: Any,
    vubel_key: str,
    email: str,
    password: str,
) -> tuple[TestResult, list[dict[str, Any]]]:
    """Đọc inbox MailFree bằng Vubel, không sử dụng cookie Shopee."""
    email = _normalize_email_for_addmail(email)
    password = str(password or "").strip()
    if not vubel_key:
        return TestResult(False, "Thiếu Vubel API Key", None), []
    if not EMAIL_ADDRESS_RE.fullmatch(email):
        return TestResult(False, f"Email MailFree không hợp lệ: {email!r}", None), []
    if not password:
        return TestResult(False, "Thiếu mật khẩu inbox MailFree.", None), []

    headers = {
        "x-api-key": vubel_key,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    try:
        response = client.post(
            VUBEL_EMAIL_GET_URL,
            headers=headers,
            json={"email": email, "password": password},
        )
        data = _decode_vubel_response(response)
        status_code = int(getattr(response, "status_code", 0) or 0)
        explicit_failure = _vubel_has_explicit_failure(data)
        if not 200 <= status_code < 300 or explicit_failure:
            message = vubel_message(data, _vubel_response_text(data) or f"HTTP {status_code}")
            return TestResult(False, f"Không đọc được MailFree: {message}", status_code, data), []
        messages = extract_vubel_inbox_messages(data)
        return TestResult(True, f"Đã tải {len(messages)} thư của {email}.", status_code, data), messages
    except Exception as exc:
        return TestResult(False, f"Lỗi đọc MailFree qua Vubel: {exc}", None), []


def vubel_find_mailfree_password(
    client: Any,
    vubel_key: str,
    email: str,
    *,
    max_pages: int = 3,
) -> str:
    """Tìm lại mật khẩu inbox trong lịch sử Vubel khi response tạo chưa có."""
    target = _normalize_email_for_addmail(email).casefold()
    if not vubel_key or not target:
        return ""
    headers = {"x-api-key": vubel_key, "Accept": "application/json"}
    limit = 100
    for page in range(max(1, min(int(max_pages), 10))):
        try:
            response = client.get(
                VUBEL_EMAIL_HISTORY_URL,
                headers=headers,
                params={"limit": limit, "offset": page * limit},
            )
            if not 200 <= int(getattr(response, "status_code", 0) or 0) < 300:
                return ""
            data = _decode_vubel_response(response)
        except Exception:
            return ""
        container = data.get("data") if isinstance(data, dict) else None
        items = container.get("items") if isinstance(container, dict) else None
        if not isinstance(items, list):
            return ""
        for item in items:
            if not isinstance(item, dict):
                continue
            candidate = _normalize_email_for_addmail(str(item.get("email") or ""))
            if candidate.casefold() == target:
                return str(item.get("password") or "").strip()
        try:
            total = int(container.get("total") or 0) if isinstance(container, dict) else 0
        except (TypeError, ValueError):
            total = 0
        if len(items) < limit or (total and (page + 1) * limit >= total):
            break
    return ""


def vubel_add_mail(
    client: Any,
    account: str,
    email: str,
    random_mode: bool,
    vubel_key: str,
    proxy_url: str | None,
    on_ivs_needed: Callable[[str, list, str | None], str | None] | None = None,
) -> tuple[TestResult, str]:
    """Vubel AddMail; giữ ``random_mode`` chỉ để tương thích bản cấu hình cũ.

    QUY TẮC:
      - Nếu đầu vào là User|Pass|SPC_F (hoặc SPC_F|User|Pass), BẮT BUỘC
        gọi Vubel login để làm mới sang SPC_ST trước, rồi mới chạy AddMail.
      - Nếu đầu vào đã có SPC_ST thì bỏ qua bước làm mới và AddMail trực tiếp.
      - AddMail: gửi email người dùng dán và không được đổi sang mail random.
      - MailFree được chuyển ngay sang ``vubel_generate_mailfree`` và không đi
        qua bất kỳ đoạn parse/login/cookie nào của AddMail.
      - Proxy Shopee chỉ nằm trong body Vubel; request tới api.vubel.store đi direct.
    """
    if not vubel_key:
        return TestResult(False, "Thiếu Vubel API Key", None), email

    if random_mode:
        return vubel_generate_mailfree(client, vubel_key, proxy_url)

    resolved_cookie = ""

    def make_result(
        success: bool, message: str, status_code: int | None, response_data: Any = None,
    ) -> TestResult:
        # Cookie đã lấy được vẫn cần xuất Excel nếu bước AddMail sau đó lỗi.
        return TestResult(success, message, status_code, response_data, resolved_cookie=resolved_cookie)

    parsed = _parse_addmail_account(account)
    resolved_cookie = parsed["cookie"]
    if parsed["kind"] == "unknown":
        return make_result(
            False,
            "Không nhận diện được tài khoản. Hãy dán User|Pass|SPC_F=... hoặc Cookie có SPC_ST=...",
            None,
        ), email

    if parsed["kind"] == "spc_f" and (not parsed["username"] or not parsed["password"]):
        return make_result(
            False,
            "Đã nhận SPC_F nhưng thiếu User/Pass. Cần SPC_F=...|User|Pass hoặc User|Pass|SPC_F=...",
            None,
        ), email

    # Giữ bản parse gốc cho debug của AddMail. MailFree đã return ở trên và
    # không bao giờ đi qua bước refresh SPC_ST này.
    source_parsed = dict(parsed)
    refreshed_spcst = False
    refresh_result: dict[str, Any] | None = None

    if parsed["kind"] == "spc_f":
        combo = f"{parsed['username']}|{parsed['password']}|SPC_F={parsed['spc_f']}"
        # Làm mới SPC_ST trực tiếp qua Shopee (không qua Vubel). Login phải đi qua
        # proxy của account nên tạo client proxied riêng; ``client`` (addmail) là
        # client direct dành cho request Vubel nên không dùng ở đây.
        try:
            with build_httpx_client(proxy_url=proxy_url) as login_client:
                refresh_result = shopee_login_spcst(
                    login_client,
                    combo,
                    proxy_url,
                    on_ivs_needed,
                )
        except Exception as exc:
            refresh_result = {
                "success": False,
                "message": f"Không tạo được client login Shopee: {exc}",
                "data": None,
            }
        if not refresh_result.get("success"):
            return make_result(
                False,
                "Làm mới SPC_ST trước AddMail thất bại: "
                + str(refresh_result.get("message") or "Không lấy được SPC_ST"),
                refresh_result.get("status_code"),
                refresh_result.get("data"),
            ), email

        resolved_cookie = str(refresh_result.get("cookie") or "").strip()
        parsed = _parse_addmail_account(resolved_cookie)
        if parsed.get("kind") != "spc_st" or not parsed.get("spc_st"):
            return make_result(
                False,
                "Shopee báo refresh thành công nhưng không tìm thấy SPC_ST trong cookie trả về.",
                refresh_result.get("status_code"),
                refresh_result.get("data"),
            ), email
        refreshed_spcst = True

    if not random_mode:
        email = _normalize_email_for_addmail(email)
        if not EMAIL_ADDRESS_RE.fullmatch(email):
            return make_result(False, f"Email không hợp lệ: {email!r}", None), email

    proxy = normalize_proxy_url(proxy_url)
    headers = {
        "x-api-key": vubel_key,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    base_payload: dict[str, Any] = {"proxy": proxy}

    if parsed["kind"] == "spc_st":
        headers["Cookie"] = parsed["cookie"] or f"SPC_ST={parsed['spc_st']}"
    else:
        base_payload.update({
            "username": parsed["username"],
            "password": parsed["password"],
            "spc_f": parsed["spc_f"],
        })

    def decode(resp: Any) -> Any:
        try:
            return resp.json()
        except Exception:
            return {"raw_text": str(getattr(resp, "text", "") or "")[:4000]}

    def response_text(data: Any) -> str:
        out: list[str] = []
        def walk(obj: Any, depth: int = 0):
            if depth > 3:
                return
            if isinstance(obj, dict):
                for k, v in obj.items():
                    normalized_key = re.sub(r"[^a-z0-9]", "", str(k).casefold())
                    if normalized_key in {
                        "message", "msg", "error", "errors", "errormessage",
                        "errormsg", "errordescription", "detail", "status",
                        "code", "errorcode", "rawtext",
                    } and v not in (None, ""):
                        out.append(str(v))
                    elif isinstance(v, (dict, list)):
                        walk(v, depth + 1)
            elif isinstance(obj, list):
                for v in obj[:20]:
                    walk(v, depth + 1)
        walk(data)
        return " ".join(out).strip()

    def missing_input(data: Any) -> str | None:
        text = response_text(data)
        m = re.search(r"(?:missing|invalid)\s+['\"]?([A-Za-z0-9_-]+)['\"]?\s+input", text, re.I)
        if m:
            return m.group(1).lower()
        if isinstance(data, dict):
            for obj in (data, data.get("data") if isinstance(data.get("data"), dict) else None):
                if not isinstance(obj, dict):
                    continue
                for k in ("required_input", "requiredInput", "field"):
                    v = obj.get(k)
                    if isinstance(v, str) and v.strip():
                        return v.strip().lower()
        return None

    def is_input_required(data: Any) -> bool:
        text = response_text(data).lower()
        if "input_required" in text or "input required" in text or ("missing " in text and " input" in text):
            return True
        if isinstance(data, dict):
            for obj in (data, data.get("data") if isinstance(data.get("data"), dict) else None):
                if isinstance(obj, dict) and str(obj.get("status") or "").lower() == "input_required":
                    return True
        return False

    def is_explicit_error(data: Any) -> bool:
        def walk(obj: Any, depth: int = 0) -> bool:
            if depth > 6:
                return False
            if isinstance(obj, dict):
                if obj.get("success") is False:
                    return True
                if obj.get("ok") is False:
                    return True
                for key, value in obj.items():
                    normalized_key = re.sub(r"[^a-z0-9]", "", str(key).casefold())
                    if normalized_key in {
                        "error", "errors", "errormessage", "errormsg",
                        "errordescription", "errorcode", "exception",
                    } and value not in (None, 0, "0", False, "", [], {}):
                        return True
                status = str(obj.get("status") or "").strip().lower()
                if status in {"failed", "error", "fail", "input_required"}:
                    return True
                return any(walk(child, depth + 1) for child in obj.values() if isinstance(child, (dict, list)))
            if isinstance(obj, list):
                return any(walk(child, depth + 1) for child in obj[:100])
            return False

        return walk(data)

    def is_explicit_success(data: Any) -> bool:
        def walk(obj: Any, depth: int = 0) -> bool:
            if depth > 6:
                return False
            if isinstance(obj, dict):
                if obj.get("success") is True:
                    return True
                if obj.get("ok") is True:
                    return True
                status = str(obj.get("status") or "").strip().lower()
                if status in {"success", "succeeded", "ok", "completed", "done"}:
                    return True
                return any(walk(child, depth + 1) for child in obj.values() if isinstance(child, (dict, list)))
            if isinstance(obj, list):
                return any(walk(child, depth + 1) for child in obj[:100])
            return False

        if walk(data):
            return True
        text = response_text(data).strip().casefold()
        negative = ("không thành công", "thất bại", "that bai", "failed", "error")
        if any(marker in text for marker in negative):
            return False
        return any(marker in text for marker in (
            "thêm mail thành công",
            "them mail thanh cong",
            "add mail success",
            "email added successfully",
        ))

    def continuation_state(data: Any) -> dict[str, Any]:
        wanted = {
            "sessionId", "session_id", "requestId", "request_id", "taskId", "task_id",
            "requestState", "request_state", "flowId", "flow_id", "stateId", "state_id",
            "tokenId", "token_id",
        }
        found: dict[str, Any] = {}
        def walk(obj: Any, depth: int = 0):
            if depth > 4:
                return
            if isinstance(obj, dict):
                for k, v in obj.items():
                    if k in wanted and v not in (None, "", [], {}):
                        found[k] = v
                    elif isinstance(v, (dict, list)):
                        walk(v, depth + 1)
            elif isinstance(obj, list):
                for v in obj[:20]:
                    walk(v, depth + 1)
        walk(data)
        return found

    def send(body: dict[str, Any]):
        return client.post(VUBEL_ADDMAIL_URL, headers=headers, json=body)

    attempts: list[str] = []
    if refreshed_spcst:
        attempts.append("REFRESH:SPC_F→SPC_ST")
    actual_email = "" if random_mode else email

    try:
        # ================================================================
        # DIRECT EMAIL FLOW: dùng ĐÚNG email người dùng dán, KHÔNG mailfree2
        # ================================================================
        if not random_mode:
            add_arg = _build_addmail_argument(email, parsed)
            if not add_arg:
                return make_result(False, "Không tạo được input addmail từ dữ liệu hiện tại.", None), email

            # Manual mode phải khai báo email ngay từ request đầu. Request INIT
            # trống có thể khiến backend tự chọn mail free trước khi client kịp
            # trả lời flow input_required. Giữ cả mailfree và addmail trong mọi
            # request tiếp theo để backend không bao giờ suy diễn sang random.
            manual_payload = dict(base_payload)
            manual_payload.update({
                "email": email,
                "mail": email,
                "mailfree": email,
                "addmail": add_arg,
            })

            def mismatched_response_email(payload: Any) -> str:
                for detected_email in extract_emails_from_data(payload):
                    normalized = _normalize_email_for_addmail(detected_email)
                    if normalized and normalized.lower() != email.lower():
                        return normalized
                return ""

            response = send(dict(manual_payload))
            data = decode(response)
            attempts.append("DIRECT:INIT(manual-fields)")
            actual_email = email
            prefixed_retry_used = False

            for _round in range(8):
                text_now = response_text(data)
                text_lower = text_now.lower()
                field = missing_input(data) or ""

                mismatch = mismatched_response_email(data)
                if mismatch:
                    return make_result(
                        False,
                        f"Vubel trả về email khác email đã dán "
                        f"(yêu cầu: {email}; trả về: {mismatch}). "
                        f"Đã đánh dấu THẤT BẠI để không lưu nhầm mail random. "
                        f"Flow: {' → '.join(attempts)}",
                        getattr(response, "status_code", None),
                        data,
                    ), mismatch

                # Nếu backend báo invalid addmail sau một lần gửi, thử biến thể
                # có tiền tố lệnh /addmail nhưng tuyệt đối không đụng mailfree2.
                if "invalid addmail input" in text_lower:
                    if prefixed_retry_used:
                        return make_result(
                            False,
                            f"Vubel tiếp tục báo invalid addmail sau lần thử /addmail. "
                            f"Đã dừng để tránh gửi lặp. Flow: {' → '.join(attempts)}",
                            getattr(response, "status_code", None),
                            data,
                        ), email
                    prefixed_retry_used = True
                    state = continuation_state(data)
                    retry = dict(manual_payload)
                    retry.update(state)
                    prefixed = f"/addmail {add_arg}"
                    retry.update({
                        "field": "addmail",
                        "required_input": "addmail",
                        "input": prefixed,
                        "addmail": prefixed,
                    })
                    response = send(retry)
                    data = decode(response)
                    attempts.append("DIRECT:addmail:/command")
                    continue

                if not is_input_required(data):
                    break

                state = continuation_state(data)

                if field == "mailfree":
                    # DIRECT = mailfree input chính là email đã dán.
                    # KHÔNG gọi /mailfree2 và KHÔNG chấp nhận email random khác.
                    value = email
                    body = dict(manual_payload)
                    body.update(state)
                    body.update({
                        "field": "mailfree",
                        "required_input": "mailfree",
                        "input": value,
                        "mailfree": value,
                        "email": value,
                    })
                    response = send(body)
                    data = decode(response)
                    attempts.append("DIRECT:mailfree=email")
                    continue

                if field == "addmail":
                    body = dict(manual_payload)
                    body.update(state)
                    body.update({
                        "field": "addmail",
                        "required_input": "addmail",
                        "input": add_arg,
                        "addmail": add_arg,
                        "email": email,
                    })
                    response = send(body)
                    data = decode(response)
                    attempts.append("DIRECT:addmail=argument")
                    continue

                # Một số response chỉ ghi input_required mà không chỉ rõ field.
                # Suy luận theo flow an toàn: trước tiên mailfree=email, sau đó addmail.
                if not field:
                    if not any(x.startswith("DIRECT:mailfree=") for x in attempts):
                        body = dict(manual_payload)
                        body.update(state)
                        body.update({
                            "field": "mailfree",
                            "required_input": "mailfree",
                            "input": email,
                            "mailfree": email,
                            "email": email,
                        })
                        response = send(body)
                        data = decode(response)
                        attempts.append("DIRECT:mailfree=email(inferred)")
                        continue
                    if not any(x.startswith("DIRECT:addmail=") for x in attempts):
                        body = dict(manual_payload)
                        body.update(state)
                        body.update({
                            "field": "addmail",
                            "required_input": "addmail",
                            "input": add_arg,
                            "addmail": add_arg,
                            "email": email,
                        })
                        response = send(body)
                        data = decode(response)
                        attempts.append("DIRECT:addmail=argument(inferred)")
                        continue

                # Nếu backend yêu cầu credential theo từng vòng, cung cấp đúng field.
                credential_map = {
                    "username": parsed.get("username", ""),
                    "password": parsed.get("password", ""),
                    "spc_f": parsed.get("spc_f", ""),
                    "spcf": parsed.get("spc_f", ""),
                    "proxy": proxy,
                    "email": email,
                }
                value = credential_map.get(field, "")
                if value not in (None, ""):
                    body = dict(manual_payload)
                    body.update(state)
                    body.update({
                        "field": field,
                        "required_input": field,
                        "input": value,
                        field: value,
                    })
                    response = send(body)
                    data = decode(response)
                    attempts.append(f"DIRECT:{field}=value")
                    continue

                return make_result(
                    False,
                    f"API yêu cầu input chưa hỗ trợ: {field or '(không rõ)'} | "
                    f"Flow: {' → '.join(attempts)} | Server: {text_now[:220]}",
                    getattr(response, "status_code", None),
                    data,
                ), email

            mismatch = mismatched_response_email(data)
            if mismatch:
                return make_result(
                    False,
                    f"Vubel trả về email khác email đã dán "
                    f"(yêu cầu: {email}; trả về: {mismatch}). "
                    f"Đã đánh dấu THẤT BẠI để không lưu nhầm mail random. "
                    f"Flow: {' → '.join(attempts)}",
                    getattr(response, "status_code", None),
                    data,
                ), mismatch

            detected = extract_email_from_data(data)
            if detected:
                norm = _normalize_email_for_addmail(detected)
                actual_email = norm

            ok = (
                200 <= getattr(response, "status_code", 0) < 300
                and is_explicit_success(data)
                and not is_explicit_error(data)
                and not is_input_required(data)
            )
            msg = vubel_message(data, "Thêm mail thành công" if ok else f"HTTP {getattr(response, 'status_code', '?')}")
            if not ok:
                server_text = response_text(data)[:260]
                msg = f"{msg} | Flow: {' → '.join(attempts)}"
                if server_text and server_text.lower() not in msg.lower():
                    msg += f" | Server: {server_text}"
            else:
                msg = f"{msg} | Flow: {' → '.join(attempts)}"

            if isinstance(data, dict):
                data = dict(data)
                dbg = data.setdefault("_client_debug", {})
                if isinstance(dbg, dict):
                    dbg.update({
                        "mode": "direct_email_flow",
                        "flow": attempts,
                        "actual_email": actual_email,
                        "requested_email": email,
                        "proxy": proxy,
                        "addmail_argument": add_arg,
                        "refreshed_spcst": refreshed_spcst,
                        "source_kind": source_parsed.get("kind", ""),
                    })
            return make_result(ok, msg, getattr(response, "status_code", None), data), actual_email

    except Exception as exc:
        return make_result(
            False,
            f"Lỗi gọi Vubel addmail: {exc} | Flow: {' → '.join(attempts) or 'chưa gửi'}",
            None,
        ), actual_email


def vubel_add_mail_manual(
    client: Any,
    account: str,
    email: str,
    vubel_key: str,
    proxy_url: str | None,
    on_ivs_needed: Callable[[str, list, str | None], str | None] | None = None,
) -> tuple[TestResult, str]:
    """Luồng AddMail riêng: chỉ liên kết đúng email được truyền vào."""
    return vubel_add_mail(
        client,
        account,
        email,
        False,
        vubel_key,
        proxy_url,
        on_ivs_needed,
    )


def vubel_mailfree(
    client: Any,
    account: str,
    vubel_key: str,
    proxy_url: str | None,
    on_ivs_needed: Callable[[str, list, str | None], str | None] | None = None,
) -> tuple[TestResult, str]:
    """Luồng MailFree riêng; ``account`` chỉ dùng để ghép dòng tại giao diện."""
    _ = account, on_ivs_needed
    return vubel_generate_mailfree(client, vubel_key, proxy_url)

def build_httpx_client(proxy_url: str | None = None) -> Any:
    if httpx is None:
        raise RuntimeError('Thiếu httpx. Chạy: python -m pip install "httpx[http2]"')
    kwargs: dict[str, Any] = {
        "http2": True,
        "timeout": 30.0,  # Tăng thời gian timeout lên 30 giây để tránh bị nghẽn
        "follow_redirects": False,
        "trust_env": False,
    }
    proxy_norm = normalize_proxy_url(proxy_url)
    if proxy_norm:
        kwargs["proxy"] = proxy_norm
    try:
        return httpx.Client(**kwargs)
    except TypeError:
        kwargs.pop("proxy", None)
        if proxy_norm:
            kwargs["proxies"] = {"all://": proxy_norm}
        try:
            return httpx.Client(**kwargs)
        except ImportError as exc:
            if proxy_norm.startswith(("socks5://", "socks5h://")):
                raise RuntimeError(
                    'Thiếu gói SOCKS5 cho httpx. Hãy cài: python -m pip install "httpx[socks]"'
                ) from exc
            raise
    except ImportError as exc:
        if proxy_norm.startswith(("socks5://", "socks5h://")):
            raise RuntimeError(
                'Thiếu gói SOCKS5 cho httpx. Hãy cài: python -m pip install "httpx[socks]"'
            ) from exc
        raise

class KiotProxyKeyExpiredError(RuntimeError):
    """Nhà cung cấp xác nhận key hết hạn."""


def acquire_proxy_until_ready(manager, stop_event, on_retry=None, force_new=False):
    """Giữ nguyên tài khoản đang chờ IP; nút Dừng luôn ngắt được thời gian chờ."""
    while not stop_event.is_set():
        try:
            return manager.rotate(force_new=force_new)
        except KiotProxyKeyExpiredError:
            raise
        except Exception as exc:
            if manager.static_proxy or manager.key_or_url.startswith(("http://", "https://")):
                raise
            for remaining in range(5, 0, -1):
                if on_retry:
                    on_retry(f"KiotProxy: {exc} • lấy lại IP sau {remaining} giây")
                if stop_event.wait(1):
                    return None
    return None


class ProxyManager:
    def __init__(
        self,
        key_or_url: str,
        rotate_interval: float | None = None,
        proxy_protocol: Any = PROXY_PROTOCOL_HTTP,
    ):
        self.key_or_url = str(key_or_url or "").strip()
        try:
            self.key_or_url = normalize_manual_proxy(self.key_or_url, proxy_protocol)
        except ValueError:
            pass
        self.proxy_protocol = infer_manual_proxy_protocol(self.key_or_url, proxy_protocol)
        configured_interval = parse_proxy_rotate_seconds(rotate_interval)
        self.has_configured_interval = configured_interval is not None
        self.rotate_interval = (
            DEFAULT_PROXY_ROTATE_SECONDS if configured_interval is None else configured_interval
        )
        self.last_rotate_time = 0.0
        self.current_proxy: str | None = None
        self.http_proxy: str | None = None
        self.socks5_proxy: str | None = None
        self.key_expiry_timestamp: int | None = None
        self.proxy_expiry_timestamp: int | None = None
        self.last_status_message = "Chưa kiểm tra KiotProxy"
        self.static_proxy = self._looks_like_static_proxy(self.key_or_url)

    @staticmethod
    def _looks_like_static_proxy(value: str) -> bool:
        """Nhận IP:PORT / user:pass@IP:PORT là proxy tĩnh, không coi là URL API đổi IP."""
        v = str(value or "").strip()
        if not v:
            return False
        no_scheme = re.sub(r"^[A-Za-z0-9+.-]+://", "", v)
        # URL API thường có /path; proxy tĩnh thường chỉ authority host:port.
        authority, sep, path = no_scheme.partition("/")
        if sep and path.strip():
            return False
        hostpart = authority.rsplit("@", 1)[-1]
        return bool(re.fullmatch(r"(?:\[[0-9A-Fa-f:]+\]|[^:]+):\d{2,5}", hostpart))

    def needs_rotation(self) -> bool:
        if not self.key_or_url:
            return False
        if self.static_proxy:
            return self.current_proxy is None
        if self.current_proxy is None:
            return True
        return time.monotonic() - self.last_rotate_time >= self.rotate_interval

    def get_remaining_time(self) -> int:
        if not self.key_or_url or self.current_proxy is None:
            return 0
        elapsed = time.monotonic() - self.last_rotate_time
        remaining = int(self.rotate_interval - elapsed)
        return max(0, remaining)

    def get_key_remaining_time(self) -> int | None:
        """Số giây còn lại của key khi API trả ngày hết hạn."""

        if self.key_expiry_timestamp is None:
            return None
        return max(0, int(self.key_expiry_timestamp - time.time()))

    def get_proxy_remaining_time(self) -> int | None:
        """Số giây còn lại của IP hiện tại; proxy tĩnh không có TTL đổi IP."""

        if self.static_proxy or self.current_proxy is None:
            return None
        if self.proxy_expiry_timestamp is None:
            return None
        return max(0, int(self.proxy_expiry_timestamp - time.time()))

    def status_snapshot(self, message: str | None = None) -> dict[str, Any]:
        """Trạng thái an toàn để đẩy lên UI, không bao gồm API key."""

        return {
            "success": bool(self.current_proxy),
            "message": message or self.last_status_message,
            "protocol": self.proxy_protocol,
            "current_proxy": self.current_proxy or "",
            "http_proxy": self.http_proxy or "",
            "socks5_proxy": self.socks5_proxy or "",
            "key_expiry_timestamp": self.key_expiry_timestamp,
            "key_remaining_seconds": self.get_key_remaining_time(),
            "proxy_remaining_seconds": self.get_proxy_remaining_time(),
        }

    def _use_provider_interval(self, provider_seconds: float) -> None:
        """Giữ số giây người dùng nhập; caller cũ vẫn dùng TTC của nhà cung cấp."""

        if not self.has_configured_interval:
            self.rotate_interval = (
                provider_seconds if provider_seconds > 0 else DEFAULT_PROXY_ROTATE_SECONDS
            )

    def _extract_proxy(self, data: dict) -> tuple[str | None, float]:
        details = extract_kiotproxy_proxy_info(data, self.proxy_protocol)
        proxy_str = details.get("selected_proxy") or None
        ttc = details.get("ttc_seconds")
        return proxy_str, float(ttc if ttc is not None else 120.0)

    @staticmethod
    def _response_data(response: Any) -> Any:
        try:
            return response.json()
        except Exception:
            try:
                response.raise_for_status()
            except Exception:
                raise
            raise ValueError(f"Phản hồi không hợp lệ: {str(getattr(response, 'text', '') or '')[:160]}")

    @staticmethod
    def _payload_is_success(data: Any, status_code: int | None) -> bool:
        if not isinstance(data, dict):
            return 200 <= int(status_code or 0) < 300
        if data.get("success") is False or data.get("ok") is False:
            return False
        if data.get("success") is True or data.get("ok") is True:
            return True
        if str(data.get("code") or "") == "200":
            return True
        return "data" in data and 200 <= int(status_code or 0) < 300

    @staticmethod
    def _api_error_message(data: Any) -> str:
        if isinstance(data, dict):
            for key in ("message", "msg", "error", "error_message", "error_msg"):
                value = data.get(key)
                if value not in (None, "", {}, []):
                    return str(value)[:240]
        return str(data)[:240]

    @staticmethod
    def _raise_if_key_expired(data: Any) -> None:
        if not isinstance(data, dict) or data.get("success") is True:
            return
        error = str(data.get("error") or "").upper()
        message = str(data.get("message") or "").casefold()
        if error in {"KEY_EXPIRED", "KEY_EXPIRE", "KEY_EXPIRATION", "PACKAGE_EXPIRED"} or (
            "key" in message and ("hết hạn" in message or "expired" in message)
        ):
            raise KiotProxyKeyExpiredError("Key KiotProxy đã hết hạn: " + str(data.get("message") or error))

    def _apply_proxy_data(self, data: Any) -> tuple[str | None, float]:
        inner = data.get("data", {}) if isinstance(data, dict) else {}
        if isinstance(inner, dict):
            self.proxy_expiry_timestamp = _coerce_timestamp(inner.get("expirationAt"))
        details = extract_kiotproxy_proxy_info(data, self.proxy_protocol)
        self.http_proxy = str(details.get("http_proxy") or "") or None
        self.socks5_proxy = str(details.get("socks5_proxy") or "") or None
        expiry_timestamp = details.get("key_expiry_timestamp")
        if expiry_timestamp is not None:
            self.key_expiry_timestamp = int(expiry_timestamp)

        proxy_str = str(details.get("selected_proxy") or "").strip()
        if not proxy_str or ":" not in proxy_str:
            return None, float(details.get("ttc_seconds") or 120.0)

        self.current_proxy = proxy_str
        ttc = float(details.get("ttc_seconds") or 120.0)
        self._use_provider_interval(ttc)
        self.last_rotate_time = time.monotonic()
        return self.current_proxy, ttc

    def _static_proxy(self) -> str:
        self.current_proxy = normalize_proxy_url(self.key_or_url, self.proxy_protocol)
        self.http_proxy = self.current_proxy if self.proxy_protocol == PROXY_PROTOCOL_HTTP else None
        self.socks5_proxy = self.current_proxy if self.proxy_protocol == PROXY_PROTOCOL_SOCKS5 else None
        self.last_rotate_time = time.monotonic()
        # Proxy tĩnh không tự đổi; giữ nguyên cho tới khi người dùng thay cấu hình.
        self.rotate_interval = 10**9
        self.last_status_message = "Proxy tĩnh đã sẵn sàng"
        return self.current_proxy

    def _api_url(self, endpoint: str) -> str:
        return f"{endpoint}?key={quote(self.key_or_url, safe='')}"

    def refresh_current(self) -> dict[str, Any]:
        """Lấy proxy hiện tại và metadata hạn key của KiotProxy."""

        if not self.key_or_url:
            raise ValueError("Chưa nhập KiotProxy Key/Proxy")
        if self.static_proxy:
            self._static_proxy()
            return self.status_snapshot("Địa chỉ proxy tĩnh hợp lệ")
        if httpx is None:
            raise RuntimeError('Thiếu httpx. Chạy: python -m pip install "httpx[http2]"')

        is_url = self.key_or_url.startswith(("http://", "https://"))
        api_url = self.key_or_url if is_url else self._api_url(KIOTPROXY_CURRENT_URL)
        try:
            with httpx.Client(timeout=15.0, trust_env=False) as client:
                response = client.get(api_url)
                data = self._response_data(response)
                self._raise_if_key_expired(data)
                if not self._payload_is_success(data, getattr(response, "status_code", None)):
                    raise ValueError(f"API báo lỗi: {self._api_error_message(data)}")
                proxy_str, _ttc = self._apply_proxy_data(data)
                if not proxy_str:
                    raise ValueError(f"API chưa trả proxy hiện tại: {self._api_error_message(data)}")
                self.last_status_message = "Đã lấy proxy hiện tại từ KiotProxy"
                return self.status_snapshot()
        except KiotProxyKeyExpiredError:
            raise
        except Exception as exc:
            self.last_status_message = f"KiotProxy lỗi: {exc}"
            raise RuntimeError(str(exc)) from exc

    def rotate(self, force_new: bool = False) -> str:
        """Lấy proxy cho lần chạy.

        ``force_new=True`` bỏ qua endpoint ``current`` ở lần lấy đầu tiên và
        gọi thẳng endpoint ``new``. Các lần xoay sau khi đã có current_proxy
        vốn luôn đi endpoint ``new`` như trước.
        """
        if self.static_proxy:
            return self._static_proxy()

        if httpx is None:
            raise RuntimeError('Thiếu httpx. Chạy: python -m pip install "httpx[http2]"')

        is_url = self.key_or_url.startswith(("http://", "https://"))
        try:
            with httpx.Client(timeout=15.0, trust_env=False) as client:
                if not is_url and self.current_proxy is None and not force_new:
                    try:
                        resp_cur = client.get(self._api_url(KIOTPROXY_CURRENT_URL))
                        data_cur = self._response_data(resp_cur)
                        self._raise_if_key_expired(data_cur)
                        if self._payload_is_success(data_cur, getattr(resp_cur, "status_code", None)):
                            p_str, _p_ttc = self._apply_proxy_data(data_cur)
                            if p_str:
                                self.last_status_message = "Đã lấy proxy hiện tại từ KiotProxy"
                                return p_str
                    except KiotProxyKeyExpiredError:
                        raise
                    except Exception:
                        pass 

                api_url = self.key_or_url if is_url else self._api_url(KIOTPROXY_NEW_URL)
                resp = client.get(api_url)
                data = self._response_data(resp)
                self._raise_if_key_expired(data)
                if not self._payload_is_success(data, getattr(resp, "status_code", None)):
                    raise ValueError(f"API báo lỗi: {self._api_error_message(data)}")

                p_str, p_ttc = self._apply_proxy_data(data)
                if p_str:
                    provider_interval = p_ttc if not is_url else DEFAULT_PROXY_ROTATE_SECONDS
                    self._use_provider_interval(provider_interval)
                    self.last_status_message = "Đã lấy proxy mới từ KiotProxy"
                    return p_str
                raise ValueError(f"API chưa trả proxy: {self._api_error_message(data)}")

        except KiotProxyKeyExpiredError:
            raise
        except Exception as e:
            raise RuntimeError(str(e))

def get_shopee_account_info(client: Any, cookie: str) -> dict:
    """Kiểm tra và lấy thông tin tài khoản Shopee từ Cookie"""
    headers = build_shopee_headers(cookie)
    try:
        response = client.get(ACCOUNT_INFO_URL, headers=headers)
        if response.status_code == 200:
            data = response.json()
            if data.get("error") or data.get("error_msg") or data.get("error") == 1:
                return {"status": "error", "msg": f"Cookie Die/Không hợp lệ ({data.get('error_msg','')})"}
            
            account_data = data.get("data", data)
            userid = account_data.get("userid")
            if userid:
                return {
                    "status": "success",
                    "username": account_data.get("username", ""),
                    "email": account_data.get("email", ""),
                    "phone": account_data.get("phone", "")
                }
            else:
                return {"status": "error", "msg": "Cookie bị giới hạn (Không có UserID)"}
        elif response.status_code == 403:
            return {"status": "error", "msg": "Bị chặn 403 (Hệ thống Anti-bot)"}
        else:
            return {"status": "error", "msg": f"Lỗi HTTP {response.status_code}"}
    except Exception as e:
        return {"status": "error", "msg": f"Lỗi kết nối: {str(e)}"}

def _message_from_error_config(config: Any) -> str:
    if not isinstance(config, dict):
        return ""
    for key in ("message", "error_msg", "content", "description"):
        value = config.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""

def interpret_shopee_response(data: Any) -> tuple[bool, str]:
    if not isinstance(data, dict):
        return False, "Phản hồi không đúng định dạng JSON"

    error = data.get("error")
    error_msg = str(data.get("error_msg") or "").strip()
    response_data = data.get("data")
    response_data = response_data if isinstance(response_data, dict) else {}

    invalid_code = response_data.get("invalid_message_code")
    error_config = response_data.get("error_config")
    configured_message = _message_from_error_config(error_config)
    voucher = response_data.get("voucher")
    voucher = voucher if isinstance(voucher, dict) else {}

    if error not in (0, "0"):
        return False, f"Từ chối: {error_msg or f'Lỗi {error}'}"

    if voucher.get("is_claimed_before") is True:
        return True, "Voucher đã có trong kho"

    if invalid_code not in (None, 0, "0"):
        try:
            explanation = INVALID_MESSAGE_EXPLANATIONS.get(int(invalid_code), "")
        except (TypeError, ValueError):
            explanation = ""
        detail = configured_message or explanation or "Không đủ điều kiện"
        return False, f"Từ chối: {detail} (Mã: {invalid_code})"

    if error_config:
        detail = configured_message or "Lỗi cấu hình"
        return False, f"Từ chối: {detail}"

    return True, "Lưu thành công"

def run_voucher_test_with_client(client: Any, cookie: str, code: str) -> TestResult:
    try:
        entry = parse_voucher_entry(code)
        if not entry["is_valid"]:
            detail = "thiếu promotionId/signature trong link" if entry["is_link"] else "mã voucher rỗng"
            return TestResult(False, f"Voucher không hợp lệ ({detail})", None)
        url, payload = build_voucher_request(entry)
        response = client.post(
            url,
            headers=build_shopee_headers(cookie),
            json=payload,
        )
        try:
            data = response.json()
        except Exception:
            data = {"raw_text": str(response.text or "")[:1000]}

        if response.status_code != 200:
            return TestResult(False, f"Lỗi HTTP {response.status_code}", response.status_code, data)

        success, message = interpret_shopee_response(data)
        return TestResult(success, message, response.status_code, data)
    except Exception as exc:
        return TestResult(False, f"Lỗi kết nối/Proxy: {exc}", None)

def run_tasks_batch(
    tasks: list[tuple[str, str]],
    delay: float,
    proxy_manager: ProxyManager | None,
    vubel_key: str,
    stop_event: threading.Event,
    cookie_info_cache: dict[str, dict],
    on_start: Callable[[int, str, str], None],
    on_result: Callable[[int, str, str, str, TestResult], None],
    on_proxy_log: Callable[[str], None],
    on_ivs_needed: Callable[[str, list, str | None], str | None],
    refresh_only: bool = False,
    skip_locked: bool = False,
) -> None:
    current_client = None
    active_cookie = None
    resolved_cookies: dict[str, dict[str, Any]] = {}
    locked_cookies: set[str] = set()

    def get_or_create_client(raw_cookie: str):
        nonlocal current_client
        # Giữ nguyên kết nối cho toàn bộ voucher của tài khoản đang xử lý,
        # kể cả khi bộ đếm đổi IP hết thời gian giữa các voucher.
        if current_client is not None and raw_cookie == active_cookie:
            return current_client

        change_account_proxy = (
            active_cookie is not None
            and proxy_manager is not None
            and not proxy_manager.static_proxy
        )
        if change_account_proxy:
            remaining = max(
                0.0,
                proxy_manager.rotate_interval
                - (time.monotonic() - proxy_manager.last_rotate_time),
            )
            if remaining > 0:
                on_proxy_log(f"Đang chờ {remaining:.1f}s để đổi IP cho cookie tiếp theo...")
                if stop_event.wait(remaining):
                    return None

        rotate_attempts = 0
        while proxy_manager and (change_account_proxy or proxy_manager.needs_rotation()):
            if stop_event.is_set():
                return None
            rotate_attempts += 1
            on_proxy_log("Đang lấy IP mới...")
            try:
                new_proxy = proxy_manager.rotate()
                on_proxy_log(f"IP Mới: {proxy_display_name(new_proxy)}")
                if current_client:
                    current_client.close()
                    current_client = None
                current_client = build_httpx_client(proxy_url=new_proxy)
                break
            except Exception as exc:
                if rotate_attempts >= MAX_PROXY_ROTATE_ATTEMPTS:
                    raise RuntimeError(
                        f"Không thể lấy IP sau {MAX_PROXY_ROTATE_ATTEMPTS} lần: {exc}"
                    ) from exc
                on_proxy_log(f"Lỗi IP: {exc} - Thử lại sau {PROXY_RETRY_DELAY_SECONDS:g}s")
                if stop_event.wait(PROXY_RETRY_DELAY_SECONDS):
                    return None

        if current_client is None:
            if stop_event.is_set():
                return None
            proxy_url = proxy_manager.current_proxy if proxy_manager else None
            current_client = build_httpx_client(proxy_url=proxy_url)
        return current_client

    try:
        for index, (raw_cookie, code) in enumerate(tasks, 1):
            if stop_event.is_set():
                break
            # Cookie đã bị khoá và bật tùy chọn bỏ qua: không gửi request nữa,
            # đánh dấu BỎ QUA và chuyển sang cookie kế tiếp.
            if skip_locked and raw_cookie in locked_cookies:
                on_start(index, raw_cookie, code)
                on_result(
                    index, raw_cookie, code, raw_cookie,
                    TestResult(False, "Bỏ qua: tài khoản bị khoá", None, skipped=True),
                )
                continue
            try:
                client = get_or_create_client(raw_cookie)
            except Exception as exc:
                message = f"Lỗi Proxy/khởi tạo kết nối của luồng: {exc}"
                on_proxy_log(message)
                # Proxy thuộc riêng lane này: đánh dấu phần còn lại của lane lỗi,
                # nhưng không chặn các lane/proxy khác hoàn tất.
                for failed_index in range(index, len(tasks) + 1):
                    failed_cookie, failed_code = tasks[failed_index - 1]
                    on_start(failed_index, failed_cookie, failed_code)
                    on_result(
                        failed_index,
                        failed_cookie,
                        failed_code,
                        failed_cookie,
                        TestResult(False, message, None),
                    )
                break
            if stop_event.is_set() or client is None:
                break

            active_cookie = raw_cookie
            on_start(index, raw_cookie, code)

            if raw_cookie not in resolved_cookies:
                if "|" in raw_cookie:
                    on_proxy_log("Đang làm mới SPC_ST trực tiếp qua Shopee...")
                    login_result = shopee_login_spcst(
                        client,
                        raw_cookie,
                        proxy_manager.current_proxy if proxy_manager else None,
                        on_ivs_needed,
                    )
                    resolved_cookies[raw_cookie] = login_result
                    if login_result.get("success"):
                        on_proxy_log("Shopee: Đã lấy SPC_ST thành công!")
                    elif skip_locked and login_result.get("locked"):
                        # Đánh dấu để các voucher còn lại của cookie này được bỏ qua.
                        locked_cookies.add(raw_cookie)
                        on_proxy_log("Tài khoản bị khoá — bỏ qua, chuyển cookie tiếp theo.")
                elif refresh_only:
                    resolved_cookies[raw_cookie] = {
                        "success": False,
                        "message": "Chế độ làm mới SPC_ST yêu cầu mỗi dòng dạng User|Pass|SPC_F",
                        "data": None,
                    }
                else:
                    resolved_cookies[raw_cookie] = {"success": True, "cookie": raw_cookie}

            res_cookie_obj = resolved_cookies[raw_cookie]
            if not res_cookie_obj.get("success"):
                fail = TestResult(False, str(res_cookie_obj.get("message") or "Lỗi"), res_cookie_obj.get("status_code"), res_cookie_obj.get("data"))
                on_result(index, raw_cookie, code, raw_cookie, fail)
                on_proxy_log("")
                if index < len(tasks) and stop_event.wait(delay):
                    break
                continue

            cookie = str(res_cookie_obj.get("cookie") or raw_cookie)

            if refresh_only:
                # Kiểm tra info chỉ để bổ sung dữ liệu Excel; SPC_ST đã lấy được vẫn được xem là thành công.
                if raw_cookie not in cookie_info_cache:
                    cookie_info_cache[raw_cookie] = get_shopee_account_info(client, cookie)
                refresh_res = TestResult(
                    True,
                    "Làm mới SPC_ST thành công",
                    res_cookie_obj.get("status_code") or 200,
                    res_cookie_obj.get("data"),
                )
                on_result(index, raw_cookie, code, cookie, refresh_res)
                if index < len(tasks) and stop_event.wait(delay):
                    break
                continue

            if raw_cookie not in cookie_info_cache:
                cookie_info_cache[raw_cookie] = get_shopee_account_info(client, cookie)

            info = cookie_info_cache[raw_cookie]
            if info.get("status") != "success":
                fail_res = TestResult(False, f"{info.get('msg')}", None)
                on_result(index, raw_cookie, code, cookie, fail_res)
                if index < len(tasks) and stop_event.wait(delay):
                    break
                continue

            result = run_voucher_test_with_client(client, cookie, code)
            on_result(index, raw_cookie, code, cookie, result)

            if index < len(tasks) and stop_event.wait(delay):
                break
    finally:
        if current_client:
            current_client.close()


def partition_voucher_tasks(
    tasks: list[tuple[str, str]],
    worker_count: int,
) -> list[list[tuple[int, str, str]]]:
    """Chia theo tài khoản để một cookie luôn giữ cùng luồng/proxy."""

    if worker_count < 1:
        raise ValueError("Cần ít nhất 1 luồng Voucher")
    lanes: list[list[tuple[int, str, str]]] = [[] for _ in range(worker_count)]
    tasks_by_account: dict[str, list[tuple[int, str, str]]] = {}
    for global_index, (raw_cookie, code) in enumerate(tasks, 1):
        tasks_by_account.setdefault(raw_cookie, []).append((global_index, raw_cookie, code))
    for account_index, account_tasks in enumerate(tasks_by_account.values()):
        lanes[account_index % worker_count].extend(account_tasks)
    return lanes


def run_voucher_workers(
    tasks: list[tuple[str, str]],
    delay: float,
    proxy_managers: list[ProxyManager | None],
    vubel_key: str,
    stop_event: threading.Event,
    cookie_info_cache: dict[str, dict],
    on_start: Callable[[int, str, str], None],
    on_result: Callable[[int, str, str, str, TestResult], None],
    on_proxy_log: Callable[[int, str], None],
    on_ivs_needed: Callable[[str, list, str | None], str | None],
    refresh_only: bool = False,
    skip_locked: bool = False,
) -> None:
    """Chạy đồng thời các lane; mỗi lane sở hữu ProxyManager/client riêng."""

    if not proxy_managers:
        raise ValueError("Cần ít nhất 1 cấu hình luồng Voucher")
    lanes = partition_voucher_tasks(tasks, len(proxy_managers))
    cache_lock = threading.Lock()
    ivs_lock = threading.Lock()
    error_lock = threading.Lock()
    worker_errors: list[tuple[int, Exception]] = []
    workers: list[threading.Thread] = []

    def run_lane(
        worker_number: int,
        indexed_tasks: list[tuple[int, str, str]],
        proxy_manager: ProxyManager | None,
    ) -> None:
        lane_tasks = [(raw_cookie, code) for _, raw_cookie, code in indexed_tasks]
        local_cache: dict[str, dict] = {}

        def global_index(local_index: int) -> int:
            if not 1 <= local_index <= len(indexed_tasks):
                raise IndexError(f"Luồng {worker_number}: task index không hợp lệ")
            return indexed_tasks[local_index - 1][0]

        def emit_start(local_index: int, raw_cookie: str, code: str) -> None:
            on_start(global_index(local_index), raw_cookie, code)

        def emit_result(
            local_index: int,
            raw_cookie: str,
            code: str,
            resolved_cookie: str,
            result: TestResult,
        ) -> None:
            if raw_cookie in local_cache:
                with cache_lock:
                    cookie_info_cache[raw_cookie] = local_cache[raw_cookie]
            on_result(
                global_index(local_index),
                raw_cookie,
                code,
                resolved_cookie,
                result,
            )

        def emit_proxy(message: str) -> None:
            on_proxy_log(worker_number, message)

        def serialized_ivs(session_id: str, methods: list, proxy_url: str | None) -> str | None:
            with ivs_lock:
                if stop_event.is_set():
                    return None
                return on_ivs_needed(session_id, methods, proxy_url)

        try:
            run_tasks_batch(
                lane_tasks,
                delay,
                proxy_manager,
                vubel_key,
                stop_event,
                local_cache,
                emit_start,
                emit_result,
                emit_proxy,
                serialized_ivs,
                refresh_only=refresh_only,
                skip_locked=skip_locked,
            )
        except Exception as exc:
            with error_lock:
                worker_errors.append((worker_number, exc))
            stop_event.set()

    for worker_number, (lane, manager) in enumerate(zip(lanes, proxy_managers), 1):
        if not lane:
            continue
        worker = threading.Thread(
            target=run_lane,
            args=(worker_number, lane, manager),
            name=f"VoucherWorker-{worker_number}",
            daemon=True,
        )
        workers.append(worker)
        worker.start()

    for worker in workers:
        worker.join()

    if worker_errors:
        worker_number, exc = worker_errors[0]
        raise RuntimeError(f"Luồng Voucher {worker_number}: {exc}") from exc


class IVSDialog(tk.Toplevel):
    def __init__(self, parent, session_id, methods, api_key, proxy_url, result_container, wait_event):
        super().__init__(parent)
        self.title("Vubel - Xác Minh Bảo Mật IVS")
        self.geometry("450x260")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.session_id = session_id
        self.api_key = api_key
        self.proxy_url = proxy_url
        self.result_container = result_container
        self.wait_event = wait_event
        
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        
        ttk.Label(self, text="Shopee yêu cầu xác minh bảo mật do IP lạ.", font=("Segoe UI", 10, "bold"), foreground="red").pack(pady=(15, 5))
        
        self.method_var = tk.IntVar()
        self.frame_radios = ttk.Frame(self)
        self.frame_radios.pack(anchor="w", padx=25, fill="x", pady=5)
        
        if methods:
            self.method_var.set(methods[0].get('type'))
            for m in methods:
                ttk.Radiobutton(self.frame_radios, text=f"{m.get('label')} (Type: {m.get('type')})", value=m.get('type'), variable=self.method_var).pack(anchor="w", pady=3)
        
        self.status_lbl = ttk.Label(self, text="Chọn phương thức và bấm Gửi Yêu Cầu.", foreground="blue")
        self.status_lbl.pack(pady=10)
        
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=5)
        
        self.btn_send = ttk.Button(btn_frame, text="1. Gửi Yêu Cầu", command=self.send_ivs)
        self.btn_send.pack(side="left", padx=5)
        
        self.btn_check = ttk.Button(btn_frame, text="2. Đã Duyệt - Kiểm Tra Cookie", command=self.check_status, state="disabled")
        self.btn_check.pack(side="left", padx=5)

    def get_client(self):
        # Request tới api.vubel.store đi trực tiếp. Proxy Shopee được truyền
        # trong payload để backend Vubel sử dụng cho phiên Shopee.
        return build_httpx_client(None)

    def send_ivs(self):
        sel_type = self.method_var.get()
        
        # Bổ sung proxy vào IVS payload đề phòng API cũng yêu cầu bắt buộc
        p_url = self.proxy_url or ""
        if p_url and not p_url.startswith("http"):
            p_url = f"http://{p_url}"
            
        payload = {
            'sessionId': self.session_id, 
            'type': sel_type, 
            'method': sel_type,
            'proxy': p_url
        }
        headers = {'x-api-key': self.api_key, 'Content-Type': 'application/json'}
        self.status_lbl.config(text="Đang gửi yêu cầu xác minh...", foreground="blue")
        self.btn_send.config(state="disabled")
        
        def task():
            try:
                with self.get_client() as client:
                    res = client.post("https://api.vubel.store/v1/shopee/login/select", json=payload, headers=headers)
                    try:
                        resp_data = res.json()
                    except Exception:
                        resp_data = {"raw_text": str(res.text or "")[:1000]}
                    if not (200 <= res.status_code < 300):
                        raise RuntimeError(f"HTTP {res.status_code}: {vubel_message(resp_data, str(resp_data)[:250])}")
                self.after(0, lambda: self.status_lbl.config(text="Đã gửi phương thức xác minh. Hãy duyệt OTP/Captcha/App rồi bấm Nút 2.", foreground="#087a2f"))
                self.after(0, lambda: self.btn_check.config(state="normal"))
            except Exception as e:
                self.after(0, lambda: self.status_lbl.config(text=f"Lỗi: {e}", foreground="red"))
            finally:
                self.after(0, lambda: self.btn_send.config(state="normal"))
        
        threading.Thread(target=task, daemon=True).start()

    def check_status(self):
        self.status_lbl.config(text="Đang kiểm tra trạng thái...", foreground="blue")
        self.btn_check.config(state="disabled")
        url = f"https://api.vubel.store/v1/shopee/login/status?sessionId={self.session_id}"
        headers = {'x-api-key': self.api_key}
        
        def task():
            try:
                with self.get_client() as client:
                    res = client.get(url, headers=headers)
                    spc_st = extract_spcst_from_httpx(res)
                    if spc_st:
                        self.result_container['spc_st'] = spc_st
                        self.after(0, self.on_success)
                    else:
                        data = res.json()
                        is_pending = data.get('status') == 'pending' or data.get('data', {}).get('status') == 'pending'
                        msg = "Vẫn đang chờ (Bạn chưa duyệt trên App/Email)" if is_pending else "Phiên lỗi / Không tìm thấy Cookie"
                        self.after(0, lambda: self.status_lbl.config(text=msg, foreground="red"))
            except Exception as e:
                self.after(0, lambda: self.status_lbl.config(text=f"Lỗi kiểm tra: {e}", foreground="red"))
            finally:
                self.after(0, lambda: self.btn_check.config(state="normal"))
                
        threading.Thread(target=task, daemon=True).start()
        
    def on_success(self):
        messagebox.showinfo("Thành công", "Đã lấy được SPC_ST, hệ thống sẽ tiếp tục xử lý!", parent=self)
        self.on_close()

    def on_close(self):
        self.wait_event.set()
        self.destroy()

class ExportDialog(tk.Toplevel):
    def __init__(self, parent, results_by_iid, tree, cookie_info_cache):
        super().__init__(parent)
        self.title("Xuất kết quả Excel")
        self.geometry("480x230")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.results_by_iid = results_by_iid
        self.tree = tree
        self.cookie_info_cache = cookie_info_cache
        self.export_success = tk.BooleanVar(value=True)
        self.export_fail = tk.BooleanVar(value=True)

        ttk.Label(self, text="Chọn nhóm kết quả đưa vào workbook:", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=15, pady=(15, 6))
        ttk.Checkbutton(self, text="Sheet Thành Công", variable=self.export_success).pack(anchor="w", padx=25)
        ttk.Checkbutton(self, text="Sheet Thất Bại", variable=self.export_fail).pack(anchor="w", padx=25, pady=2)
        ttk.Label(
            self,
            text="Excel có cột SPC_ST nếu đã lấy được cookie, kể cả kết quả thất bại, và cột 3 dòng: SPC_F → USER → MẬT KHẨU.",
            foreground="#555",
            wraplength=440,
        ).pack(anchor="w", padx=15, pady=(12, 4))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=15, pady=15)
        ttk.Button(btn_frame, text="Hủy", command=self.destroy).pack(side="right", padx=(5, 0))
        ttk.Button(btn_frame, text="Xuất Excel", command=self.process_export).pack(side="right")

    def _append_sheet(self, ws, rows):
        headers = [
            "STT", "Mã Voucher", "Tài khoản", "Email", "SĐT", "SPC_F", "SPC_ST",
            "SPC_F / USER / MẬT KHẨU (3 dòng)", "FULL COOKIE", "Trạng thái", "Thông báo"
        ]
        ws.append(headers)
        if Font is not None:
            for cell in ws[1]:
                cell.font = Font(bold=True)
        for row in rows:
            ws.append(row)
        # Cột H là định dạng copy/paste 3 dòng: SPC_F -> USER -> MẬT KHẨU.
        if Alignment is not None:
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row=row_idx, column=8).alignment = Alignment(wrap_text=True, vertical="top")
                ws.row_dimensions[row_idx].height = 45
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        widths = [7, 18, 22, 30, 18, 38, 55, 42, 65, 16, 55]
        for idx, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(idx)].width = width

    def process_export(self):
        if Workbook is None:
            messagebox.showerror("Thiếu thư viện", "Cần cài openpyxl:\npython -m pip install openpyxl", parent=self)
            return
        if not self.export_success.get() and not self.export_fail.get():
            messagebox.showwarning("Lỗi", "Vui lòng chọn ít nhất 1 nhóm để xuất!", parent=self)
            return

        path = filedialog.asksaveasfilename(
            title="Lưu kết quả Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile="Shopee_KetQua.xlsx",
            parent=self,
        )
        if not path:
            return

        success_rows = []
        fail_rows = []
        for iid in self.tree.get_children():
            values = self.tree.item(iid, "values")
            data = self.results_by_iid.get(iid)
            if not data:
                continue

            raw = data.raw_cookie
            resolved = data.resolved_cookie or raw
            info = self.cookie_info_cache.get(raw, {})
            username = info.get("username") if info.get("status") == "success" else ""
            email = info.get("email") if info.get("status") == "success" else ""
            phone = info.get("phone") if info.get("status") == "success" else ""
            combo_pass = ""
            if "|" in raw:
                combo_user, combo_pass, combo_spcf = parse_combo(raw)
                username = username or combo_user
                spcf = combo_spcf or cookie_value(resolved, "SPC_F")
            else:
                spcf = cookie_value(resolved, "SPC_F") or cookie_value(raw, "SPC_F")
            spcst = cookie_value(resolved, "SPC_ST") or cookie_value(raw, "SPC_ST")
            paste_3_lines = ""
            if spcf and username and combo_pass:
                paste_3_lines = f"SPC_F={spcf}\n{username}\n{combo_pass}"
            row = [
                int(values[0]),
                values[3],
                username,
                email,
                phone,
                f"SPC_F={spcf}" if spcf else "",
                f"SPC_ST={spcst}" if spcst else "",
                paste_3_lines,
                raw,
                values[4],
                values[5],
            ]
            (success_rows if data.test_result.success else fail_rows).append(row)

        try:
            wb = Workbook()
            default_ws = wb.active
            wb.remove(default_ws)
            if self.export_success.get():
                self._append_sheet(wb.create_sheet("Thanh Cong"), success_rows)
            if self.export_fail.get():
                self._append_sheet(wb.create_sheet("That Bai"), fail_rows)
            wb.save(path)
            try:
                open_file_in_default_app(path)
                open_note = "\n\nĐã tự mở file Excel."
            except Exception as open_exc:
                open_note = f"\n\n{open_exc}"
            messagebox.showinfo("Thành công", f"Đã xuất Excel:\n{path}{open_note}", parent=self)
            self.destroy()
        except Exception as exc:
            messagebox.showerror("Lỗi Lưu Excel", f"Không thể lưu file: {exc}", parent=self)


@dataclass
class IIDData:
    raw_cookie: str
    resolved_cookie: str
    test_result: TestResult

class VoucherTesterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"Shopee Voucher & SPC_ST Tool • {OWNER_NAME} • v3.18")
        apply_ryan_window_icon(self.root)
        self.root.geometry("1500x900")
        self.root.minsize(1180, 720)
        self.running = False
        self.stop_event = threading.Event()
        self.result_by_iid: dict[str, IIDData] = {}
        self.cookie_info_cache: dict[str, dict] = {}
        self.tasks: list[tuple[str, str]] = []
        self.ok_count = 0
        self.fail_count = 0
        self.already_saved_count = 0
        self.new_saved_count = 0
        self.skipped_count = 0
        self.processed_count = 0
        self.total_count = 0
        self.batch_refresh_only = False
        self.batch_skip_locked = True
        self.voucher_thread_configs = [
            VoucherThreadConfig("", DEFAULT_PROXY_ROTATE_SECONDS)
        ]
        self.proxy_managers: list[ProxyManager | None] = []
        self.proxy_worker_status: dict[int, str] = {}
        self.active_ivs_dialog: IVSDialog | None = None
        self.mail_running = False
        self.mail_stop_event = threading.Event()
        self.mail_results: list[dict[str, Any]] = []
        self.mailfree_running = False
        self.mailfree_stop_event = threading.Event()
        self.mailfree_results: list[dict[str, Any]] = []
        self.mailfree_mailboxes: list[dict[str, str]] = []
        self._last_alarm_target: str | None = None
        self._alarm_thread_running = False
        self._completion_alarm_thread_running = False
        self.native_mail_suite = None
        self._kiotproxy_manager: ProxyManager | None = None
        self._kiotproxy_check_running = False
        self._kiotproxy_check_generation = 0

        # Custom frameless window: 3 nút thu nhỏ / phóng to / đóng được lồng vào header.
        self._window_drag_x = 0
        self._window_drag_y = 0
        self._window_maximized = False
        self._window_normal_geometry: str | None = None
        self._restore_chrome_after_map = False
        try:
            self.root.overrideredirect(False)
            self.root.resizable(True, True)
        except Exception:
            pass
        self.root.bind("<Map>", self._on_window_map, add="+")

        self._setup_ui()
        self._load_config() 
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing) 

    def _save_config(self):
        """Lưu toàn bộ nội dung đã nhập vào file JSON"""
        config = {
            "codes": self.codes_text.get("1.0", "end-1c"),
            "cookies": self.cookies_text.get("1.0", "end-1c"),
            "proxy": self.proxy_var.get(),
            "inbox_refresh_all": self.inbox_refresh_all_var.get(),
            "inbox_refresh_seconds": {
                name: getattr(self.native_mail_suite, name).auto_refresh_seconds.get()
                for name in ("mailtm", "tinyhost", "outlook")
            } if self.native_mail_suite is not None else {},
            "proxy_protocol": self.proxy_protocol_var.get(),
            "kiotproxy_auto_new_ip": bool(self.kiotproxy_auto_new_ip_var.get()),
            "addmail_rotate_seconds": self.addmail_rotate_seconds_var.get(),
            "delay": self.delay_var.get(),
            "voucher_delay": self.voucher_delay_var.get(),
            "voucher_threads": [
                {
                    "proxy": item.proxy_key,
                    "rotate_seconds": item.rotate_seconds,
                }
                for item in self.voucher_thread_configs
            ],
            "use_timer": self.use_timer_var.get(),
            "timer_value": self.timer_var.get(),
            "alarm_5s": self.alarm_5s_var.get(),
            "completion_alarm": self.completion_alarm_var.get(),
            "refresh_only": self.refresh_only_var.get(),
            "skip_locked": self.skip_locked_var.get(),
            "voucher_mask_code": self.voucher_mask_code_var.get(),
            "voucher_hide_code_col": self.voucher_hide_code_col_var.get(),
            "background_mode": self.background_mode_var.get(),
            "mail_accounts": self.mail_accounts_text.get("1.0", "end-1c"),
            "mail_emails": self.mail_emails_text.get("1.0", "end-1c"),
            "mailfree_accounts": self.mailfree_accounts_text.get("1.0", "end-1c"),
            "mailfree_mailboxes": [
                {
                    "email": str(item.get("email") or ""),
                    "password": str(item.get("password") or ""),
                    "created_at": str(item.get("created_at") or ""),
                }
                for item in self.mailfree_mailboxes[-MAX_BATCH_COOKIES:]
                if item.get("email") and item.get("password")
            ],
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Lỗi khi lưu cấu hình: {e}")

    def _load_config(self):
        """Đọc và tự động điền cấu hình từ file JSON nếu có"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    config = json.load(f)
                
                if "codes" in config and config["codes"]:
                    self.codes_text.insert("1.0", config["codes"])
                if "cookies" in config and config["cookies"]:
                    self.cookies_text.insert("1.0", config["cookies"])
                # vubel_key không còn lưu/đọc từ file — đã chuyển sang RAM-only
                if "proxy" in config:
                    self.proxy_var.set(config["proxy"])
                self.inbox_refresh_all_var.set(str(config.get("inbox_refresh_all", "0")))
                intervals = config.get("inbox_refresh_seconds", {})
                if self.native_mail_suite is not None and isinstance(intervals, dict):
                    for name in ("mailtm", "tinyhost", "outlook"):
                        getattr(self.native_mail_suite, name).auto_refresh_seconds.set(str(intervals.get(name, self.inbox_refresh_all_var.get())))
                if "proxy_protocol" in config:
                    self.proxy_protocol_var.set(normalize_proxy_protocol(config["proxy_protocol"]))
                if str(config.get("manual_proxy") or "").strip():
                    self.proxy_var.set(str(config["manual_proxy"]).strip())
                    self.proxy_protocol_var.set(normalize_proxy_protocol(config.get("manual_proxy_protocol")))
                if "kiotproxy_auto_new_ip" in config:
                    self.kiotproxy_auto_new_ip_var.set(bool(config["kiotproxy_auto_new_ip"]))
                self.addmail_rotate_seconds_var.set(str(config.get("addmail_rotate_seconds") or ""))
                if "delay" in config:
                    self.delay_var.set(config["delay"])
                self.voucher_delay_var.set(f"{voucher_delay_from_saved(config):g}")
                try:
                    self.voucher_thread_configs = voucher_thread_configs_from_saved(config)
                except ValueError as exc:
                    print(f"Cấu hình luồng Voucher không hợp lệ: {exc}")
                    self.voucher_thread_configs = [
                        VoucherThreadConfig("", DEFAULT_PROXY_ROTATE_SECONDS)
                    ]
                self._update_voucher_thread_count()
                if "use_timer" in config:
                    self.use_timer_var.set(config["use_timer"])
                if "timer_value" in config:
                    self.timer_var.set(config["timer_value"])
                if "alarm_5s" in config:
                    self.alarm_5s_var.set(config["alarm_5s"])
                if "completion_alarm" in config:
                    self.completion_alarm_var.set(config["completion_alarm"])
                if "refresh_only" in config:
                    self.refresh_only_var.set(config["refresh_only"])
                if "skip_locked" in config:
                    self.skip_locked_var.set(bool(config["skip_locked"]))
                if "voucher_mask_code" in config:
                    self.voucher_mask_code_var.set(bool(config["voucher_mask_code"]))
                if "voucher_hide_code_col" in config:
                    self.voucher_hide_code_col_var.set(bool(config["voucher_hide_code_col"]))
                if "background_mode" in config:
                    self.background_mode_var.set(config["background_mode"])
                if "mail_accounts" in config and config["mail_accounts"]:
                    self.mail_accounts_text.insert("1.0", config["mail_accounts"])
                if "mail_emails" in config and config["mail_emails"]:
                    self.mail_emails_text.insert("1.0", config["mail_emails"])
                # v3.18: MailFree là tab riêng. Cờ mail_random cũ bị bỏ qua để
                # không thể âm thầm đổi luồng AddMail sau khi khởi động lại.
                legacy_accounts = str(config.get("mail_accounts") or "")
                if "mailfree_accounts" in config:
                    mailfree_accounts = str(config.get("mailfree_accounts") or "")
                else:
                    # Nâng cấp lần đầu: sao chép danh sách tài khoản cũ sang tab
                    # MailFree để người dùng không phải dán lại. Hai tab sẽ lưu
                    # độc lập từ lần lưu tiếp theo và không tự chạy.
                    mailfree_accounts = legacy_accounts
                if mailfree_accounts:
                    self.mailfree_accounts_text.insert("1.0", mailfree_accounts)
                stored_mailboxes = config.get("mailfree_mailboxes")
                if isinstance(stored_mailboxes, list):
                    for item in stored_mailboxes[-MAX_BATCH_COOKIES:]:
                        if not isinstance(item, dict):
                            continue
                        email = _normalize_email_for_addmail(str(item.get("email") or ""))
                        password = str(item.get("password") or "").strip()
                        if EMAIL_ADDRESS_RE.fullmatch(email) and password:
                            self._remember_mailfree_mailbox(
                                email,
                                password,
                                created_at=str(item.get("created_at") or ""),
                            )
            except Exception as e:
                print(f"Lỗi khi đọc cấu hình: {e}")
        # Áp dụng nền sau khi toàn bộ giao diện đã được tạo và cấu hình được đọc.
        try:
            self._apply_background_mode()
        except Exception as e:
            print(f"Lỗi khi áp dụng Background: {e}")
        # Cập nhật số lượng + bảng xem trước voucher sau khi ô mã voucher đã được điền từ config.
        self._update_input_counts()
        self._update_voucher_preview()

    def _on_closing(self):
        self._save_config()
        self.stop_event.set()
        self.mail_stop_event.set()
        self.mailfree_stop_event.set()
        try:
            if self.native_mail_suite is not None:
                self.native_mail_suite.stop_all()
        except Exception:
            pass
        self.root.destroy()

    # ===== Custom traffic-light window controls =====
    def _on_window_map(self, _event=None):
        """Khôi phục chế độ không viền sau khi cửa sổ được restore từ taskbar."""
        if not self._restore_chrome_after_map:
            return
        self._restore_chrome_after_map = False
        try:
            self.root.after(20, lambda: self.root.overrideredirect(False))
        except Exception:
            pass

    def _minimize_window(self):
        """Thu nhỏ cửa sổ; tạm bật native frame để Windows iconify ổn định."""
        try:
            self._restore_chrome_after_map = True
            self.root.overrideredirect(False)
            self.root.iconify()
        except Exception:
            try:
                self.root.iconify()
            except Exception:
                pass

    def _get_work_area(self):
        """Lấy vùng làm việc không che taskbar trên Windows; fallback về toàn màn hình."""
        if os.name == "nt":
            try:
                import ctypes
                from ctypes import wintypes
                rect = wintypes.RECT()
                SPI_GETWORKAREA = 0x0030
                ok = ctypes.windll.user32.SystemParametersInfoW(
                    SPI_GETWORKAREA, 0, ctypes.byref(rect), 0
                )
                if ok:
                    return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top
            except Exception:
                pass
        return 0, 0, self.root.winfo_screenwidth(), self.root.winfo_screenheight()

    def _toggle_maximize_window(self, _event=None):
        """Phóng to / khôi phục cửa sổ bằng nút xanh."""
        try:
            if not self._window_maximized:
                self._window_normal_geometry = self.root.geometry()
                x, y, w, h = self._get_work_area()
                self.root.geometry(f"{w}x{h}+{x}+{y}")
                self._window_maximized = True
            else:
                if self._window_normal_geometry:
                    self.root.geometry(self._window_normal_geometry)
                self._window_maximized = False
        except Exception:
            pass

    def _start_window_drag(self, event):
        if self._window_maximized:
            return
        try:
            self._window_drag_x = event.x_root - self.root.winfo_x()
            self._window_drag_y = event.y_root - self.root.winfo_y()
        except Exception:
            pass

    def _drag_window(self, event):
        if self._window_maximized:
            return
        try:
            x = event.x_root - self._window_drag_x
            y = event.y_root - self._window_drag_y
            self.root.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _make_traffic_button(self, parent, color: str, command, tooltip: str, symbol: str = "", symbol_color: str = "#2f2f2f"):
        """Tạo nút tròn kiểu traffic-light; có thể hiển thị ký hiệu ngay bên trong nút."""
        wrap = tk.Frame(parent, bg="#ffffff", width=24, height=28, cursor="hand2")
        wrap.pack_propagate(False)
        canvas = tk.Canvas(
            wrap, width=24, height=28, bg="#ffffff", highlightthickness=0, bd=0, cursor="hand2"
        )
        canvas.pack(fill="both", expand=True)
        dot = canvas.create_oval(4, 6, 20, 22, fill=color, outline=color)
        symbol_item = None
        if symbol:
            symbol_item = canvas.create_text(
                12, 14,
                text=symbol,
                fill=symbol_color,
                font=("Segoe UI", 9, "bold"),
                anchor="center",
            )

        def on_enter(_e):
            try:
                canvas.itemconfigure(dot, outline="#6f6f6f", width=1)
            except Exception:
                pass

        def on_leave(_e):
            try:
                canvas.itemconfigure(dot, outline=color, width=1)
            except Exception:
                pass

        def on_click(_e):
            command()

        for widget in (wrap, canvas):
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
            widget.bind("<Button-1>", on_click)
        if symbol_item is not None:
            canvas.tag_bind(symbol_item, "<Button-1>", on_click)
            canvas.tag_bind(symbol_item, "<Enter>", on_enter)
            canvas.tag_bind(symbol_item, "<Leave>", on_leave)
        # Dùng tên tooltip cho khả năng truy vết/debug, không hiển thị popup gây rối UI.
        wrap._traffic_tooltip = tooltip  # type: ignore[attr-defined]
        return wrap

    def _get_next_run_time(self, milestones_str: str) -> datetime | None:
        tz_hn = timezone(timedelta(hours=7))
        now = datetime.now(tz_hn)
        next_run = None

        for ms in milestones_str.split(','):
            ms = ms.strip()
            if not ms: continue
            try:
                parts = ms.split(':')
                h = int(parts[0])
                m = int(parts[1]) if len(parts) > 1 else 0
                s = int(parts[2]) if len(parts) > 2 else 0

                target = now.replace(hour=h, minute=m, second=s, microsecond=0)
                if target <= now:
                    target += timedelta(days=1)

                if next_run is None or target < next_run:
                    next_run = target
            except ValueError:
                continue
        return next_run

    def _play_alarm_5s(self):
        """Phát chuông cảnh báo ngắn ở T-5 giây mà không làm treo giao diện."""
        if self._alarm_thread_running:
            return
        self._alarm_thread_running = True

        def worker():
            try:
                if os.name == "nt":
                    import winsound
                    for _ in range(3):
                        winsound.Beep(1250, 180)
                        time.sleep(0.10)
                else:
                    # Fallback cho macOS/Linux: chuông hệ thống Tk.
                    for i in range(3):
                        self.root.after(i * 280, self.root.bell)
            except Exception:
                try:
                    self.root.after(0, self.root.bell)
                except Exception:
                    pass
            finally:
                self._alarm_thread_running = False

        threading.Thread(target=worker, daemon=True).start()

    def _play_completion_alarm(self):
        """Phát chuông hoàn tất sau khi toàn bộ tác vụ đã xử lý xong."""
        if self._completion_alarm_thread_running:
            return
        self._completion_alarm_thread_running = True

        def worker():
            try:
                if os.name == "nt":
                    import winsound
                    # Giai điệu tăng dần, khác chuông cảnh báo T-5s.
                    for freq, dur in ((880, 150), (1100, 160), (1320, 180), (1760, 260)):
                        winsound.Beep(freq, dur)
                        time.sleep(0.07)
                else:
                    for i in range(4):
                        self.root.after(i * 260, self.root.bell)
            except Exception:
                try:
                    self.root.after(0, self.root.bell)
                except Exception:
                    pass
            finally:
                self._completion_alarm_thread_running = False

        threading.Thread(target=worker, daemon=True).start()

    def _check_timer(self):
        if self.use_timer_var.get() and not self.running:
            tz_hn = timezone(timedelta(hours=7))
            now = datetime.now(tz_hn)
            next_run = self._get_next_run_time(self.timer_var.get())

            if next_run:
                diff = (next_run - now).total_seconds()
                target_key = next_run.isoformat()
                remaining = max(0, int(diff + 0.999))
                try:
                    self.timer_countdown_var.set(f"Còn {remaining}s")
                except Exception:
                    pass

                # Chuông chỉ phát một lần cho mỗi mốc, khi bước vào khoảng T-5 giây.
                if self.alarm_5s_var.get() and 1.0 < diff <= 5.0 and self._last_alarm_target != target_key:
                    self._last_alarm_target = target_key
                    self.status_var.set(f"🔔 Còn {remaining}s tới mốc {next_run.strftime('%H:%M:%S')}")
                    self._play_alarm_5s()
                elif diff <= 1.0:
                    self.status_var.set(f"Đã đến mốc ({next_run.strftime('%H:%M:%S')}), tự động khởi chạy!")
                    try:
                        self.timer_countdown_var.set("Đang chạy")
                    except Exception:
                        pass
                    self.start_batch()
                else:
                    self.status_var.set(f"Chờ tự động chạy lúc: {next_run.strftime('%H:%M:%S')} (còn {remaining}s)")
            else:
                try:
                    self.timer_countdown_var.set("Sai giờ")
                except Exception:
                    pass
                self.status_var.set("Lỗi: Sai định dạng giờ (Cần nhập dạng HH:MM hoặc HH:MM:SS)")

        elif not self.running and not self.use_timer_var.get():
            try:
                self.timer_countdown_var.set("Tắt")
            except Exception:
                pass
            if "Chờ tự động" in self.status_var.get() or self.status_var.get().startswith("🔔"):
                self.status_var.set("Sẵn sàng.")

        self.root.after(200, self._check_timer)

    def _setup_ui(self):
        # ===== Modern orange/white dashboard UI =====
        self.root.configure(bg="#f6f7f9")
        self.root.option_add("*Font", ("Segoe UI", 9))

        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background="#f6f7f9")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("TLabel", background="#f6f7f9", foreground="#222222")
        style.configure("Card.TLabel", background="#ffffff", foreground="#222222")
        style.configure("Muted.Card.TLabel", background="#ffffff", foreground="#7a7f87")
        style.configure("Accent.TButton", background="#ee4d2d", foreground="white", borderwidth=0, padding=(14, 8), font=("Segoe UI", 9, "bold"))
        style.map("Accent.TButton", background=[("active", "#d84225"), ("disabled", "#f2a28f")])
        style.configure("Soft.TButton", background="#ffffff", foreground="#333333", bordercolor="#dfe3e8", borderwidth=1, padding=(12, 8))
        style.map("Soft.TButton", background=[("active", "#f7f7f7")])
        style.configure("Danger.TButton", background="#ffffff", foreground="#e53935", bordercolor="#f1c5c4", borderwidth=1, padding=(12, 8))
        style.map("Danger.TButton", background=[("active", "#fff2f1")])
        style.configure("Treeview", background="#ffffff", fieldbackground="#ffffff", foreground="#3a3a3a", rowheight=31, borderwidth=0, relief="flat")
        style.configure("Treeview.Heading", background="#fbfbfc", foreground="#444444", borderwidth=0, relief="flat", font=("Segoe UI", 9, "bold"), padding=(8, 8))
        style.map("Treeview", background=[("selected", "#fff0eb")], foreground=[("selected", "#222222")])
        style.configure("Horizontal.TProgressbar", troughcolor="#edf0f3", background="#ee4d2d", bordercolor="#edf0f3", lightcolor="#ee4d2d", darkcolor="#ee4d2d")
        style.configure("TNotebook", background="#f6f7f9", borderwidth=0)
        style.configure("TNotebook.Tab", background="#eceff2", foreground="#333333", padding=(14, 8), borderwidth=0)
        style.map("TNotebook.Tab", background=[("selected", "#ffffff"), ("active", "#f7f7f7")], foreground=[("selected", "#ee4d2d"), ("active", "#333333")])

        # Shared variables are created once and reused by Voucher / Mail / Settings pages.
        # vubel_key_var đã xóa — key Vubel chỉ tồn tại server-side trong RAM (_VUBEL_KEY_RAM)
        self.inbox_refresh_all_var = tk.StringVar(value="0")
        self.proxy_var = tk.StringVar()
        self.proxy_protocol_var = tk.StringVar(value=PROXY_PROTOCOL_HTTP)
        self.kiotproxy_auto_new_ip_var = tk.BooleanVar(value=False)
        self.addmail_rotate_seconds_var = tk.StringVar(value="")
        self.kiotproxy_status_var = tk.StringVar(value="KiotProxy: Chưa kiểm tra")
        self.kiotproxy_expiry_var = tk.StringVar(value="Hạn key: chưa kiểm tra | IP hiện tại: chưa có dữ liệu")
        self.delay_var = tk.StringVar(value=str(DEFAULT_DELAY_SECONDS))
        self.voucher_delay_var = tk.StringVar(value=str(DEFAULT_VOUCHER_DELAY_SECONDS))
        self.voucher_thread_count_var = tk.StringVar(value="1 luồng")
        self.use_timer_var = tk.BooleanVar(value=False)
        self.timer_var = tk.StringVar(value="00:00:00, 12:00:00, 20:00:00")
        self.alarm_5s_var = tk.BooleanVar(value=True)
        self.completion_alarm_var = tk.BooleanVar(value=True)
        self.timer_countdown_var = tk.StringVar(value="Tắt")
        self.refresh_only_var = tk.BooleanVar(value=False)
        self.skip_locked_var = tk.BooleanVar(value=True)
        self.voucher_mask_code_var = tk.BooleanVar(value=False)
        self.voucher_hide_code_col_var = tk.BooleanVar(value=False)
        self.voucher_notes: dict[str, str] = {}
        self.voucher_selected: dict[str, bool] = {}
        self.voucher_selected_count_var = tk.StringVar(value="Đã chọn: 0/0")
        self._voucher_preview_raw_lines: list[str] = []
        # Chế độ nền: Sáng (mặc định) hoặc Đen. Giá trị được lưu trong config.
        self.background_mode_var = tk.StringVar(value="Sáng")
        self._theme_originals = {}
        self.code_count_var = tk.StringVar(value="0")
        self.cookie_count_var = tk.StringVar(value="0")

        # Summary variables (right dashboard)
        self.summary_total_var = tk.StringVar(value="0")
        self.summary_ok_var = tk.StringVar(value="0")
        self.summary_fail_var = tk.StringVar(value="0")
        self.summary_spcst_var = tk.StringVar(value="0")
        self.summary_already_var = tk.StringVar(value="0")
        self.summary_new_saved_var = tk.StringVar(value="0")
        self.summary_progress_var = tk.StringVar(value="0 / 0 (0%)")

        # ===== Header / custom title bar =====
        header = tk.Frame(self.root, bg="#ffffff", height=64, highlightthickness=1, highlightbackground="#e8eaed")
        header.pack(side="top", fill="x")
        header.pack_propagate(False)
        self.custom_header = header

        # 3 nút traffic-light được LỒNG TRỰC TIẾP vào header và đặt bên PHẢI.
        # Thứ tự: vàng (thu nhỏ) - xanh (phóng to) - đỏ (đóng).
        traffic_box = tk.Frame(header, bg="#ffffff")
        traffic_box.pack(side="right", padx=(6, 14), pady=18)
        btn_min = self._make_traffic_button(
            traffic_box, "#f7b527", self._minimize_window, "Thu nhỏ", symbol="−", symbol_color="#6b4b00"
        )
        btn_min.pack(side="left", padx=(0, 2))
        btn_max = self._make_traffic_button(
            traffic_box, "#25c759", self._toggle_maximize_window, "Phóng to / Khôi phục"
        )
        btn_max.pack(side="left", padx=2)
        btn_close = self._make_traffic_button(
            traffic_box, "#ff5f57", self._on_closing, "Đóng", symbol="×", symbol_color="#7a1712"
        )
        btn_close.pack(side="left", padx=(2, 0))

        brand = tk.Frame(header, bg="#ffffff")
        brand.pack(side="left", padx=(8, 0), pady=12)
        logo = tk.Label(brand, text="S", bg="#ee4d2d", fg="white", width=2, height=1, font=("Segoe UI", 16, "bold"), relief="flat")
        logo.pack(side="left")
        brand_title = tk.Label(brand, text="Shopee Voucher & SPC_ST Tool", bg="#ffffff", fg="#171717", font=("Segoe UI", 15, "bold"))
        brand_title.pack(side="left", padx=(12, 0))
        owner_badge = tk.Label(brand, text=f"  {OWNER_NAME}  ", bg="#fff0ea", fg="#ee4d2d", font=("Segoe UI", 9, "bold"))
        owner_badge.pack(side="left", padx=(14, 0))
        ready_badge = tk.Label(brand, text="  ●  Ready  ", bg="#e8f8ef", fg="#168a49", font=("Segoe UI", 9, "bold"))
        ready_badge.pack(side="left", padx=(8, 0))
        license_exp = str(CURRENT_LICENSE_INFO.get("expiry_text") or "Đã kích hoạt").split(" ")[0]
        license_badge = tk.Label(brand, text=f"  License: {license_exp}  ", bg="#eef4ff", fg="#3569b7", font=("Segoe UI", 8, "bold"))
        license_badge.pack(side="left", padx=(8, 0))

        menu_dots = tk.Label(header, text="⋮", bg="#ffffff", fg="#555555", font=("Segoe UI", 19))
        # Nút menu nằm ngay bên trái cụm điều khiển cửa sổ.
        menu_dots.pack(side="right", padx=(8, 4))

        # Có thể kéo cửa sổ bằng vùng header; double-click để phóng to/khôi phục.
        for drag_widget in (header, brand, logo, brand_title, owner_badge, ready_badge, license_badge, menu_dots):
            drag_widget.bind("<ButtonPress-1>", self._start_window_drag, add="+")
            drag_widget.bind("<B1-Motion>", self._drag_window, add="+")
            drag_widget.bind("<Double-Button-1>", self._toggle_maximize_window, add="+")

        # ===== Shell =====
        shell = tk.Frame(self.root, bg="#f6f7f9")
        shell.pack(fill="both", expand=True)

        self.sidebar = tk.Frame(shell, bg="#ffffff", width=210, highlightthickness=1, highlightbackground="#eceef1")
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        self.nav_buttons = {}
        nav_items = [
            ("voucher", "▣   Voucher", lambda: self._show_page("voucher", refresh=False)),
            ("readmail", "📨   Đọc Mail", lambda: self._show_page("readmail")),
            ("mail", "✉   Thêm Mail", lambda: self._show_page("mail")),
            ("excel", "▤   Xuất Excel", self.open_export_dialog),
            ("settings", "⚙   Cài đặt", lambda: self._show_page("settings")),
        ]
        nav_wrap = tk.Frame(self.sidebar, bg="#ffffff")
        nav_wrap.pack(fill="x", pady=(22, 0))
        for key, label, cmd in nav_items:
            btn = tk.Button(
                nav_wrap, text=label, command=cmd, anchor="w", relief="flat", bd=0,
                bg="#ffffff", fg="#4d535a", activebackground="#fff1ec", activeforeground="#ee4d2d",
                padx=22, pady=13, font=("Segoe UI", 10)
            )
            btn.pack(fill="x", padx=8, pady=2)
            self.nav_buttons[key] = btn

        sidebar_bottom = tk.Frame(self.sidebar, bg="#ffffff", highlightthickness=1, highlightbackground="#eceef1")
        sidebar_bottom.pack(side="bottom", fill="x", padx=16, pady=16)
        tk.Label(sidebar_bottom, text="S  Shopee", bg="#ffffff", fg="#ee4d2d", font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=12, pady=(12, 2))
        tk.Label(sidebar_bottom, text=OWNER_NAME, bg="#ffffff", fg="#222222", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=12, pady=(0, 3))
        tk.Label(sidebar_bottom, text="Phiên bản: 3.18 Ryan Nguyễn", bg="#ffffff", fg="#777777", font=("Segoe UI", 8)).pack(anchor="w", padx=12)
        tk.Label(sidebar_bottom, text="License • Native Mail • Timer + Alarm", bg="#ffffff", fg="#999999", font=("Segoe UI", 8)).pack(anchor="w", padx=12, pady=(1, 12))

        self.page_stack = tk.Frame(shell, bg="#f6f7f9")
        self.page_stack.pack(side="left", fill="both", expand=True)
        self.page_stack.grid_rowconfigure(0, weight=1)
        self.page_stack.grid_columnconfigure(0, weight=1)

        self.voucher_page = tk.Frame(self.page_stack, bg="#f6f7f9")
        self.readmail_page = tk.Frame(self.page_stack, bg="#f6f7f9")
        self.mail_page = tk.Frame(self.page_stack, bg="#f6f7f9")
        self.settings_page = tk.Frame(self.page_stack, bg="#f6f7f9")
        for page in (self.voucher_page, self.readmail_page, self.mail_page, self.settings_page):
            page.grid(row=0, column=0, sticky="nsew")

        # ===== Voucher page =====
        voucher_outer = tk.Frame(self.voucher_page, bg="#f6f7f9")
        voucher_outer.pack(fill="both", expand=True, padx=16, pady=16)
        voucher_outer.grid_columnconfigure(0, weight=1)
        voucher_outer.grid_columnconfigure(1, minsize=270)
        voucher_outer.grid_rowconfigure(0, weight=1)

        main = tk.Frame(voucher_outer, bg="#f6f7f9")
        main.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(4, weight=3)
        main.grid_rowconfigure(5, weight=2)

        # Inputs row
        inputs = tk.Frame(main, bg="#f6f7f9")
        inputs.grid(row=0, column=0, sticky="ew")
        inputs.grid_columnconfigure(0, weight=1)
        inputs.grid_columnconfigure(1, weight=1)

        code_card = self._make_card(inputs)
        code_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        code_head = tk.Frame(code_card, bg="#ffffff")
        code_head.pack(fill="x", padx=16, pady=(13, 6))
        tk.Label(code_head, text="▣  Danh sách mã Voucher", bg="#ffffff", fg="#2f3338", font=("Segoe UI", 10, "bold")).pack(side="left")
        tk.Label(code_head, textvariable=self.code_count_var, bg="#f5f6f7", fg="#50545a", padx=10, pady=3, font=("Segoe UI", 8, "bold")).pack(side="right")
        self.codes_text = scrolledtext.ScrolledText(code_card, height=6, font=("Consolas", 10), bd=0, relief="flat", wrap="word", background="#ffffff", foreground="#30343a", insertbackground="#30343a")
        self.codes_text.pack(fill="both", expand=True, padx=14, pady=(0, 5))
        tk.Label(code_card, text="Mỗi dòng 1 mã voucher HOẶC 1 link voucher Shopee (tự bóc tách promotionId/signature/evcode) — voucher trùng sẽ tự lọc bỏ", bg="#ffffff", fg="#9a9da2", font=("Segoe UI", 8), wraplength=430, justify="left").pack(anchor="w", padx=16, pady=(0, 11))

        cookie_card = self._make_card(inputs)
        cookie_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        cookie_head = tk.Frame(cookie_card, bg="#ffffff")
        cookie_head.pack(fill="x", padx=16, pady=(13, 6))
        tk.Label(cookie_head, text="◉  Danh sách Cookie / User|Pass|SPC_F", bg="#ffffff", fg="#2f3338", font=("Segoe UI", 10, "bold")).pack(side="left")
        tk.Label(cookie_head, textvariable=self.cookie_count_var, bg="#f5f6f7", fg="#50545a", padx=10, pady=3, font=("Segoe UI", 8, "bold")).pack(side="right")
        self.cookies_text = scrolledtext.ScrolledText(cookie_card, height=6, font=("Consolas", 9), bd=0, relief="flat", wrap="none", background="#ffffff", foreground="#30343a", insertbackground="#30343a")
        self.cookies_text.pack(fill="both", expand=True, padx=14, pady=(0, 5))
        tk.Label(cookie_card, text="Hỗ trợ Full Cookie hoặc User|Pass|SPC_F — Mỗi dòng 1 tài khoản", bg="#ffffff", fg="#9a9da2", font=("Segoe UI", 8)).pack(anchor="w", padx=16, pady=(0, 11))

        self.codes_text.bind("<KeyRelease>", self._update_input_counts)
        self.codes_text.bind("<KeyRelease>", self._update_voucher_preview, add="+")
        self.cookies_text.bind("<KeyRelease>", self._update_input_counts)

        # Voucher preview card (bóc tách link/mã theo từng dòng đã nhập)
        preview_card = self._make_card(main)
        preview_card.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        preview_card.grid_columnconfigure(0, weight=1)
        preview_head = tk.Frame(preview_card, bg="#ffffff")
        preview_head.pack(fill="x", padx=16, pady=(13, 6))
        tk.Label(preview_head, text="▤  Danh sách Voucher cần lưu (xem trước)", bg="#ffffff", fg="#2f3338", font=("Segoe UI", 10, "bold")).pack(side="left")
        tk.Label(preview_head, textvariable=self.voucher_selected_count_var, bg="#f5f6f7", fg="#50545a", padx=10, pady=3, font=("Segoe UI", 8, "bold")).pack(side="left", padx=(10, 0))
        ttk.Checkbutton(
            preview_head, text="Ẩn cột Mã Code", variable=self.voucher_hide_code_col_var,
            command=self._apply_voucher_code_column_visibility,
        ).pack(side="right")
        ttk.Checkbutton(
            preview_head, text="Ẩn mã (••••)", variable=self.voucher_mask_code_var,
            command=self._update_voucher_preview,
        ).pack(side="right", padx=(0, 10))

        preview_actions = tk.Frame(preview_card, bg="#ffffff")
        preview_actions.pack(fill="x", padx=16, pady=(0, 6))
        ttk.Button(
            preview_actions, text="☑ Chọn tất cả", style="Soft.TButton",
            command=self._select_all_voucher_preview,
        ).pack(side="left")
        ttk.Button(
            preview_actions, text="☐ Bỏ chọn tất cả", style="Soft.TButton",
            command=self._deselect_all_voucher_preview,
        ).pack(side="left", padx=(8, 0))

        preview_tree_frame = tk.Frame(preview_card, bg="#ffffff")
        preview_tree_frame.pack(fill="both", expand=True, padx=14, pady=(0, 12))
        preview_tree_frame.grid_rowconfigure(0, weight=1)
        preview_tree_frame.grid_columnconfigure(0, weight=1)
        # PromotionId/Signature/Link Gốc chỉ giữ trong dữ liệu dòng đã nhập
        # (đã lưu nguyên văn trong ô mã voucher / config), KHÔNG hiển thị lên bảng.
        self.voucher_preview_tree = ttk.Treeview(
            preview_tree_frame,
            columns=("select", "stt", "note", "type", "code", "valid"),
            show="headings", selectmode="browse", height=4,
        )
        for col, title, width, anchor, stretch in (
            ("select", "Chọn", 48, "center", False),
            ("stt", "STT", 40, "center", False),
            ("note", "Ghi chú", 130, "w", False),
            ("type", "Loại", 90, "center", False),
            ("code", "Mã Code (Evcode)", 160, "center", False),
            ("valid", "Hợp Lệ", 90, "center", True),
        ):
            self.voucher_preview_tree.heading(col, text=title)
            self.voucher_preview_tree.column(col, width=width, anchor=anchor, stretch=stretch)
        self.voucher_preview_tree.grid(row=0, column=0, sticky="nsew")
        preview_scroll = ttk.Scrollbar(preview_tree_frame, orient="vertical", command=self.voucher_preview_tree.yview)
        preview_scroll.grid(row=0, column=1, sticky="ns")
        self.voucher_preview_tree.configure(yscrollcommand=preview_scroll.set)
        self.voucher_preview_tree.tag_configure("valid", foreground="#168a49")
        self.voucher_preview_tree.tag_configure("invalid", foreground="#df3b3b")
        self.voucher_preview_tree.tag_configure("dup", foreground="#b76d00", background="#fff8e7")
        self.voucher_preview_tree.tag_configure("unselected", foreground="#b3b6bb")
        self.voucher_preview_tree.bind("<Double-1>", self._on_voucher_preview_double_click)
        self.voucher_preview_tree.bind("<Button-1>", self._on_voucher_preview_click)
        tk.Label(
            preview_card,
            text="Tích ô Chọn để đưa voucher vào lượt chạy • Nhấp đúp ô Ghi chú để đặt tên gợi nhớ (không gửi lên Shopee).",
            bg="#ffffff", fg="#9a9da2", font=("Segoe UI", 8),
        ).pack(anchor="w", padx=16, pady=(0, 11))

        # Settings strip
        cfg = self._make_card(main)
        cfg.grid(row=2, column=0, sticky="ew", pady=(12, 10))
        cfg.grid_columnconfigure(0, weight=1)
        cfg.grid_columnconfigure(1, weight=1)
        cfg.grid_columnconfigure(2, weight=0)
        cfg.grid_columnconfigure(3, weight=0)

        timer_box = tk.Frame(cfg, bg="#ffffff")
        timer_box.grid(row=0, column=0, sticky="ew", padx=(16, 8), pady=12)
        timer_head = tk.Frame(timer_box, bg="#ffffff")
        timer_head.pack(fill="x")
        tk.Label(timer_head, text="Mốc hẹn giờ  ⓘ", bg="#ffffff", fg="#444444", font=("Segoe UI", 8, "bold")).pack(side="left")
        tk.Label(timer_head, textvariable=self.timer_countdown_var, bg="#fff3ee", fg="#ee4d2d", padx=6, pady=1, font=("Segoe UI", 8, "bold")).pack(side="right")
        ttk.Entry(timer_box, textvariable=self.timer_var).pack(fill="x", pady=(5, 4))
        timer_opts = tk.Frame(timer_box, bg="#ffffff")
        timer_opts.pack(fill="x")
        ttk.Checkbutton(timer_opts, text="Bật hẹn giờ", variable=self.use_timer_var).pack(side="left")
        ttk.Checkbutton(timer_opts, text="🔔 Chuông T-5s", variable=self.alarm_5s_var).pack(side="left", padx=(10, 0))
        ttk.Checkbutton(timer_opts, text="🔔 Chuông hoàn tất", variable=self.completion_alarm_var).pack(side="left", padx=(10, 0))

        proxy_box = tk.Frame(cfg, bg="#ffffff")
        proxy_box.grid(row=0, column=1, sticky="ew", padx=8, pady=12)
        proxy_head = tk.Frame(proxy_box, bg="#ffffff")
        proxy_head.pack(fill="x")
        tk.Label(proxy_head, text="Luồng Voucher / Proxy  ⓘ", bg="#ffffff", fg="#444444", font=("Segoe UI", 8, "bold")).pack(side="left")
        tk.Label(proxy_head, textvariable=self.voucher_thread_count_var, bg="#fff3ee", fg="#ee4d2d", padx=6, pady=1, font=("Segoe UI", 8, "bold")).pack(side="right")
        self.voucher_threads_button = ttk.Button(
            proxy_box,
            text="Cấu hình Proxy + giây đổi IP",
            command=self._open_voucher_threads_dialog,
            style="Soft.TButton",
        )
        self.voucher_threads_button.pack(fill="x", pady=(5, 0))

        delay_box = tk.Frame(cfg, bg="#ffffff")
        delay_box.grid(row=0, column=2, sticky="w", padx=8, pady=12)
        tk.Label(delay_box, text="Độ trễ request  ⓘ", bg="#ffffff", fg="#444444", font=("Segoe UI", 8, "bold")).pack(anchor="w")
        spin_wrap = tk.Frame(delay_box, bg="#ffffff")
        spin_wrap.pack(anchor="w", pady=(5, 0))
        self.voucher_delay_spinbox = ttk.Spinbox(
            spin_wrap,
            from_=MIN_VOUCHER_DELAY_SECONDS,
            to=MAX_VOUCHER_DELAY_SECONDS,
            increment=0.5,
            width=7,
            textvariable=self.voucher_delay_var,
        )
        self.voucher_delay_spinbox.pack(side="left")
        tk.Label(spin_wrap, text=" giây", bg="#ffffff", fg="#777777", font=("Segoe UI", 8)).pack(side="left")

        mode_box = tk.Frame(cfg, bg="#ffffff")
        mode_box.grid(row=0, column=3, sticky="e", padx=(8, 16), pady=12)
        tk.Label(mode_box, text="Chỉ làm mới SPC_ST  ⓘ", bg="#ffffff", fg="#444444", font=("Segoe UI", 8, "bold")).pack(anchor="w")
        ttk.Checkbutton(mode_box, variable=self.refresh_only_var).pack(anchor="w", pady=(5, 0))
        ttk.Checkbutton(
            mode_box,
            text="Bỏ qua cookie bị khoá",
            variable=self.skip_locked_var,
        ).pack(anchor="w", pady=(6, 0))

        # Actions
        actions = tk.Frame(main, bg="#f6f7f9")
        actions.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        for i in range(4):
            actions.grid_columnconfigure(i, weight=1)
        self.run_button = ttk.Button(actions, text="▶  Bắt đầu xử lý", command=self.start_batch, style="Accent.TButton")
        self.run_button.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self.stop_button = ttk.Button(actions, text="⏹  Dừng", state="disabled", command=self.stop_batch, style="Danger.TButton")
        self.stop_button.grid(row=0, column=1, sticky="ew", padx=5)
        self.export_button = ttk.Button(actions, text="▤  Xuất Excel", command=self.open_export_dialog, style="Soft.TButton")
        self.export_button.grid(row=0, column=2, sticky="ew", padx=5)
        self.clear_button = ttk.Button(actions, text="⌫  Xóa kết quả", command=self.clear_results, style="Danger.TButton")
        self.clear_button.grid(row=0, column=3, sticky="ew", padx=(5, 0))
        self.copy_spcst_button = ttk.Button(
            actions, text="Copy nhanh SPC_ST", command=self.copy_refreshed_spcst,
            style="Soft.TButton",
        )
        self.copy_spcst_button.grid(row=1, column=2, columnspan=2, sticky="ew", padx=5, pady=(6, 0))
        self.refresh_only_var.trace_add("write", lambda *_: self._update_copy_spcst_button())
        self._update_copy_spcst_button()

        # Results card
        result_card = self._make_card(main)
        result_card.grid(row=4, column=0, sticky="nsew")
        result_card.grid_rowconfigure(1, weight=1)
        result_card.grid_columnconfigure(0, weight=1)
        tk.Label(result_card, text="Kết quả xử lý", bg="#ffffff", fg="#333333", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", padx=16, pady=(12, 8))
        tree_frame = tk.Frame(result_card, bg="#ffffff")
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=1, pady=(0, 1))
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.results_tree = ttk.Treeview(
            tree_frame,
            columns=("index", "cookie_preview", "username", "code", "status", "message"),
            show="headings", selectmode="browse"
        )
        for col, title in (
            ("index", "STT"), ("cookie_preview", "FULL COOKIE"), ("username", "Tài khoản"),
            ("code", "Mã Voucher"), ("status", "Kết quả"), ("message", "Thông báo")
        ):
            self.results_tree.heading(col, text=title)
        self.results_tree.column("index", width=48, anchor="center", stretch=False)
        self.results_tree.column("cookie_preview", width=190, anchor="w")
        self.results_tree.column("username", width=145, anchor="w")
        self.results_tree.column("code", width=120, anchor="w")
        self.results_tree.column("status", width=150, anchor="center", stretch=False)
        self.results_tree.column("message", width=360, anchor="w")
        self.results_tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.results_tree.yview)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        self.results_tree.configure(yscrollcommand=tree_scroll.set)
        self.results_tree.tag_configure("ok", foreground="#168a49", background="#effaf3")
        self.results_tree.tag_configure("fail", foreground="#df3b3b", background="#fff1f1")
        self.results_tree.tag_configure("running", foreground="#b76d00", background="#fff8e7")
        self.results_tree.tag_configure("skipped", foreground="#7a7d82", background="#f1f2f4")
        self.results_tree.tag_configure("pending", foreground="#7d8085")
        self.results_tree.bind("<<TreeviewSelect>>", self._show_selected_result)

        # Log card
        log_card = self._make_card(main)
        log_card.grid(row=5, column=0, sticky="nsew", pady=(10, 0))
        log_card.grid_rowconfigure(1, weight=1)
        log_card.grid_columnconfigure(0, weight=1)
        log_head = tk.Frame(log_card, bg="#ffffff")
        log_head.grid(row=0, column=0, sticky="ew", padx=16, pady=(10, 6))
        tk.Label(log_head, text="Chi tiết / Log", bg="#ffffff", fg="#333333", font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Button(log_head, text="Sao chép", command=self._copy_log, style="Soft.TButton").pack(side="right", padx=(6, 0))
        ttk.Button(log_head, text="Xóa log", command=lambda: self._set_details(""), style="Danger.TButton").pack(side="right")
        self.details = scrolledtext.ScrolledText(log_card, height=6, state="disabled", font=("Consolas", 9), bd=0, relief="flat", background="#ffffff", foreground="#253b59")
        self.details.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 10))

        # Status line
        status_line = tk.Frame(main, bg="#f6f7f9")
        status_line.grid(row=6, column=0, sticky="ew", pady=(7, 0))
        self.status_var = tk.StringVar(value="Sẵn sàng.")
        self.proxy_status_var = tk.StringVar(value="")
        tk.Label(status_line, textvariable=self.status_var, bg="#f6f7f9", fg="#565b61", font=("Segoe UI", 8)).pack(side="left")
        tk.Label(status_line, textvariable=self.proxy_status_var, bg="#f6f7f9", fg="#3976c1", font=("Segoe UI", 8, "bold")).pack(side="right")
        self.status_var.trace_add("write", lambda *_: self._update_summary())

        # ===== Right summary =====
        side = tk.Frame(voucher_outer, bg="#f6f7f9", width=270)
        side.grid(row=0, column=1, sticky="nsew")
        side.grid_propagate(False)
        summary = self._make_card(side)
        summary.pack(fill="x")
        tk.Label(summary, text="▥  Tổng quan", bg="#ffffff", fg="#333333", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=16, pady=(14, 8))
        self._add_stat_card(summary, "Tổng tác vụ", self.summary_total_var, "#20242a", "◎")
        self._add_stat_card(summary, "Thành công", self.summary_ok_var, "#168a49", "✓")
        self._add_stat_card(summary, "Voucher đã được lưu rồi", self.summary_already_var, "#e08a1e", "⚠")
        self._add_stat_card(summary, "Đã lưu voucher (mới)", self.summary_new_saved_var, "#168a49", "＋")
        self._add_stat_card(summary, "Thất bại", self.summary_fail_var, "#df3b3b", "×")
        self._add_stat_card(summary, "SPC_ST mới", self.summary_spcst_var, "#3575df", "⟳")

        progress_card = self._make_card(side)
        progress_card.pack(fill="x", pady=(10, 0))
        top_prog = tk.Frame(progress_card, bg="#ffffff")
        top_prog.pack(fill="x", padx=16, pady=(13, 8))
        tk.Label(top_prog, text="◷  Tiến trình", bg="#ffffff", fg="#333333", font=("Segoe UI", 10, "bold")).pack(side="left")
        tk.Label(progress_card, textvariable=self.summary_progress_var, bg="#ffffff", fg="#666666", font=("Segoe UI", 8)).pack(anchor="w", padx=16)
        self.summary_progress = ttk.Progressbar(progress_card, orient="horizontal", mode="determinate", maximum=100, style="Horizontal.TProgressbar")
        self.summary_progress.pack(fill="x", padx=16, pady=(8, 14))

        tip = self._make_card(side)
        tip.pack(fill="x", pady=(10, 0))
        tk.Label(tip, text="💡  Mẹo", bg="#ffffff", fg="#ee4d2d", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=16, pady=(13, 6))
        tk.Label(tip, text="Dùng proxy ổn định, đặt độ trễ phù hợp và kiểm tra định dạng Cookie / User|Pass|SPC_F trước khi chạy.", bg="#ffffff", fg="#777777", justify="left", wraplength=230, font=("Segoe UI", 8)).pack(anchor="w", padx=16, pady=(0, 14))

        # ===== Đọc Mail native + Thêm Mail + Settings =====
        self._setup_read_mail_page(self.readmail_page)
        self._setup_mail_tab(self.mail_page)
        self._setup_settings_page(self.settings_page)

        self._show_page("voucher", refresh=False)
        self._update_input_counts()
        self.root.after(500, self._check_timer)
        self.root.after(1000, self._update_kiotproxy_countdown)

    def _make_card(self, parent):
        return tk.Frame(parent, bg="#ffffff", highlightthickness=1, highlightbackground="#e4e7eb", bd=0)

    def _add_stat_card(self, parent, title, value_var, value_color, icon):
        box = tk.Frame(parent, bg="#ffffff", highlightthickness=1, highlightbackground="#eceef1")
        box.pack(fill="x", padx=14, pady=5)
        left = tk.Frame(box, bg="#ffffff")
        left.pack(side="left", fill="both", expand=True, padx=14, pady=10)
        tk.Label(left, text=title, bg="#ffffff", fg="#4f5358", font=("Segoe UI", 9, "bold")).pack(anchor="w")
        tk.Label(left, textvariable=value_var, bg="#ffffff", fg=value_color, font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(3, 0))
        tk.Label(box, text=icon, bg="#fff4ef", fg="#ee4d2d", width=3, font=("Segoe UI", 16, "bold")).pack(side="right", padx=12, pady=12)

    def _show_page(self, name: str, refresh: bool | None = None):
        self._active_page = name if name in {"voucher", "readmail", "mail", "settings"} else "voucher"
        if name == "readmail":
            self.readmail_page.tkraise()
            active = "readmail"
        elif name == "mail":
            self.mail_page.tkraise()
            active = "mail"
        elif name == "settings":
            self.settings_page.tkraise()
            active = "settings"
        else:
            self.voucher_page.tkraise()
            if refresh is not None:
                self.refresh_only_var.set(bool(refresh))
            active = "voucher"
        dark = str(self.background_mode_var.get()).strip().lower() in {"đen", "dark", "black"}
        for key, btn in self.nav_buttons.items():
            if dark:
                if key == active:
                    btn.configure(bg="#1b1b1b", fg="#ffffff", activebackground="#2a2a2a", activeforeground="#ffffff", font=("Segoe UI", 10, "bold"))
                else:
                    btn.configure(bg="#000000", fg="#ffffff", activebackground="#1a1a1a", activeforeground="#ffffff", font=("Segoe UI", 10))
            else:
                if key == active:
                    btn.configure(bg="#fff0ea", fg="#ee4d2d", activebackground="#fff1ec", activeforeground="#ee4d2d", font=("Segoe UI", 10, "bold"))
                else:
                    btn.configure(bg="#ffffff", fg="#4d535a", activebackground="#fff1ec", activeforeground="#ee4d2d", font=("Segoe UI", 10))

    def _apply_background_mode(self):
        """Áp dụng theme Sáng/Đen cho TOÀN BỘ giao diện.

        Ở chế độ Đen: toàn bộ frame/card/header/sidebar/ô nhập/bảng/log chuyển
        sang nền đen hoặc đen rất tối; chữ mặc định chuyển sang trắng. Các nút
        điều khiển vàng/xanh/đỏ và nút hành động màu cam vẫn giữ màu nhận diện.
        """
        dark = str(self.background_mode_var.get()).strip().lower() in {"đen", "dark", "black"}
        self._is_dark_mode = dark

        # Snapshot màu gốc đúng một lần để có thể quay lại theme Sáng chính xác.
        def snapshot(widget):
            if widget not in self._theme_originals:
                original = {}
                for opt in (
                    "background", "foreground", "activebackground", "activeforeground",
                    "highlightbackground", "highlightcolor", "insertbackground",
                    "selectbackground", "selectforeground", "readonlybackground",
                ):
                    try:
                        original[opt] = widget.cget(opt)
                    except Exception:
                        pass
                self._theme_originals[widget] = original
            try:
                children = widget.winfo_children()
            except Exception:
                children = []
            for child in children:
                snapshot(child)

        snapshot(self.root)

        # Tk widgets.
        def walk(widget):
            original = self._theme_originals.get(widget, {})
            cls = widget.winfo_class() if hasattr(widget, "winfo_class") else ""

            if not dark:
                for opt, value in original.items():
                    try:
                        widget.configure(**{opt: value})
                    except Exception:
                        pass
            else:
                try:
                    if cls in {"Frame", "Labelframe", "Toplevel"}:
                        widget.configure(background="#000000")
                        if cls == "Labelframe":
                            try:widget.configure(foreground="#ffffff")
                            except Exception:pass
                        try:
                            widget.configure(highlightbackground="#2a2a2a", highlightcolor="#2a2a2a")
                        except Exception:
                            pass
                    elif cls == "Label":
                        widget.configure(background="#000000", foreground="#ffffff")
                    elif cls in {"Entry", "Spinbox"}:
                        widget.configure(
                            background="#101010", foreground="#ffffff", insertbackground="#ffffff",
                            selectbackground="#ee4d2d", selectforeground="#ffffff",
                        )
                        try:
                            widget.configure(readonlybackground="#101010")
                        except Exception:
                            pass
                    elif cls in {"Text"}:
                        widget.configure(
                            background="#0b0b0b", foreground="#ffffff", insertbackground="#ffffff",
                            selectbackground="#ee4d2d", selectforeground="#ffffff",
                        )
                    elif cls == "Button":
                        orig_bg = str(original.get("background", "")).lower()
                        # Giữ màu của 3 nút cửa sổ và nút hành động cam.
                        preserve = orig_bg in {"#f7b527", "#25c759", "#ff5f57", "#ee4d2d", "#d84325"}
                        if preserve:
                            widget.configure(foreground="#ffffff", activeforeground="#ffffff")
                        else:
                            widget.configure(
                                background="#141414", foreground="#ffffff",
                                activebackground="#252525", activeforeground="#ffffff",
                                highlightbackground="#2a2a2a",
                            )
                    elif cls == "Checkbutton":
                        widget.configure(
                            background="#000000", foreground="#ffffff",
                            activebackground="#000000", activeforeground="#ffffff",
                            selectcolor="#151515",
                        )
                    elif cls == "Listbox":
                        widget.configure(
                            background="#0b0b0b", foreground="#ffffff",
                            selectbackground="#ee4d2d", selectforeground="#ffffff",
                        )
                except Exception:
                    pass

            try:
                children = widget.winfo_children()
            except Exception:
                children = []
            for child in children:
                walk(child)

        try:
            self.root.configure(bg="#000000" if dark else "#f6f7f9")
        except Exception:
            pass
        walk(self.root)

        # ttk widgets cần style riêng.
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        if dark:
            style.configure("TFrame", background="#000000")
            style.configure("Card.TFrame", background="#000000")
            style.configure("TLabel", background="#000000", foreground="#ffffff")
            style.configure("Card.TLabel", background="#000000", foreground="#ffffff")
            style.configure("Muted.Card.TLabel", background="#000000", foreground="#d2d2d2")

            style.configure("TEntry", fieldbackground="#101010", background="#101010", foreground="#ffffff", bordercolor="#343434", lightcolor="#343434", darkcolor="#343434", insertcolor="#ffffff")
            style.map("TEntry", fieldbackground=[("readonly", "#101010"), ("disabled", "#101010")], foreground=[("readonly", "#ffffff"), ("disabled", "#aaaaaa")])
            style.configure("TCombobox", fieldbackground="#101010", background="#171717", foreground="#ffffff", arrowcolor="#ffffff", bordercolor="#343434", lightcolor="#343434", darkcolor="#343434")
            style.map("TCombobox", fieldbackground=[("readonly", "#101010")], foreground=[("readonly", "#ffffff")], selectbackground=[("readonly", "#101010")], selectforeground=[("readonly", "#ffffff")])
            style.configure("TCheckbutton", background="#000000", foreground="#ffffff", focuscolor="#000000")
            style.map("TCheckbutton", background=[("active", "#000000")], foreground=[("active", "#ffffff")])

            style.configure("Accent.TButton", background="#ee4d2d", foreground="#ffffff", borderwidth=0, padding=(14, 8), font=("Segoe UI", 9, "bold"))
            style.map("Accent.TButton", background=[("active", "#d84225"), ("disabled", "#6b2c20")], foreground=[("disabled", "#dddddd")])
            style.configure("Soft.TButton", background="#151515", foreground="#ffffff", bordercolor="#333333", borderwidth=1, padding=(12, 8))
            style.map("Soft.TButton", background=[("active", "#252525")], foreground=[("active", "#ffffff")])
            style.configure("Danger.TButton", background="#151515", foreground="#ffffff", bordercolor="#513030", borderwidth=1, padding=(12, 8))
            style.map("Danger.TButton", background=[("active", "#2a1717")], foreground=[("active", "#ffffff")])

            style.configure("Treeview", background="#080808", fieldbackground="#080808", foreground="#ffffff", rowheight=31, borderwidth=0, relief="flat")
            style.configure("Treeview.Heading", background="#151515", foreground="#ffffff", borderwidth=0, relief="flat", font=("Segoe UI", 9, "bold"), padding=(8, 8))
            style.map("Treeview", background=[("selected", "#3b1c12")], foreground=[("selected", "#ffffff")])
            style.configure("Horizontal.TProgressbar", troughcolor="#242424", background="#ee4d2d", bordercolor="#242424", lightcolor="#ee4d2d", darkcolor="#ee4d2d")
            style.configure("TScrollbar", background="#181818", troughcolor="#050505", bordercolor="#2c2c2c", arrowcolor="#ffffff", lightcolor="#181818", darkcolor="#181818")
            style.map("TScrollbar", background=[("active", "#2b2b2b")])
            style.configure("TNotebook", background="#000000", borderwidth=0)
            style.configure("TNotebook.Tab", background="#151515", foreground="#ffffff", padding=(14, 8), borderwidth=0)
            style.map("TNotebook.Tab", background=[("selected", "#2a2a2a"), ("active", "#202020")], foreground=[("selected", "#ffffff"), ("active", "#ffffff")])
            try:
                self.root.option_add("*TCombobox*Listbox.background", "#101010")
                self.root.option_add("*TCombobox*Listbox.foreground", "#ffffff")
                self.root.option_add("*TCombobox*Listbox.selectBackground", "#ee4d2d")
                self.root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
            except Exception:
                pass

            # Tree tags không dùng nền sáng trong dark mode.
            for tree_name in ("results_tree", "mail_tree", "mailfree_tree"):
                tree = getattr(self, tree_name, None)
                if tree is not None:
                    try:
                        tree.tag_configure("ok", foreground="#ffffff", background="#0c1a11")
                        tree.tag_configure("fail", foreground="#ffffff", background="#211010")
                        tree.tag_configure("running", foreground="#ffffff", background="#211b0c")
                    except Exception:
                        pass
            try:
                self.voucher_preview_tree.tag_configure("valid", foreground="#4fd88a", background="")
                self.voucher_preview_tree.tag_configure("invalid", foreground="#ff6b6b", background="")
                self.voucher_preview_tree.tag_configure("dup", foreground="#ffffff", background="#211b0c")
                self.voucher_preview_tree.tag_configure("unselected", foreground="#6b6e73", background="")
            except Exception:
                pass
        else:
            # Khôi phục ttk style sáng nguyên bản.
            style.configure("TFrame", background="#f6f7f9")
            style.configure("Card.TFrame", background="#ffffff")
            style.configure("TLabel", background="#f6f7f9", foreground="#222222")
            style.configure("Card.TLabel", background="#ffffff", foreground="#222222")
            style.configure("Muted.Card.TLabel", background="#ffffff", foreground="#7a7f87")
            style.configure("TEntry", fieldbackground="#ffffff", background="#ffffff", foreground="#222222", bordercolor="#d9dde2", lightcolor="#d9dde2", darkcolor="#d9dde2")
            style.map("TEntry", fieldbackground=[("readonly", "#ffffff")], foreground=[("readonly", "#222222")])
            style.configure("TCombobox", fieldbackground="#ffffff", background="#ffffff", foreground="#222222", arrowcolor="#444444", bordercolor="#d9dde2", lightcolor="#d9dde2", darkcolor="#d9dde2")
            style.map("TCombobox", fieldbackground=[("readonly", "#ffffff")], foreground=[("readonly", "#222222")])
            style.configure("TCheckbutton", background="#ffffff", foreground="#222222")
            style.configure("Accent.TButton", background="#ee4d2d", foreground="white", borderwidth=0, padding=(14, 8), font=("Segoe UI", 9, "bold"))
            style.map("Accent.TButton", background=[("active", "#d84225"), ("disabled", "#f2a28f")])
            style.configure("Soft.TButton", background="#ffffff", foreground="#333333", bordercolor="#dfe3e8", borderwidth=1, padding=(12, 8))
            style.map("Soft.TButton", background=[("active", "#f7f7f7")])
            style.configure("Danger.TButton", background="#ffffff", foreground="#e53935", bordercolor="#f1c5c4", borderwidth=1, padding=(12, 8))
            style.map("Danger.TButton", background=[("active", "#fff2f1")])
            style.configure("Treeview", background="#ffffff", fieldbackground="#ffffff", foreground="#3a3a3a", rowheight=31, borderwidth=0, relief="flat")
            style.configure("Treeview.Heading", background="#fbfbfc", foreground="#444444", borderwidth=0, relief="flat", font=("Segoe UI", 9, "bold"), padding=(8, 8))
            style.map("Treeview", background=[("selected", "#fff0eb")], foreground=[("selected", "#222222")])
            style.configure("Horizontal.TProgressbar", troughcolor="#edf0f3", background="#ee4d2d", bordercolor="#edf0f3", lightcolor="#ee4d2d", darkcolor="#ee4d2d")
            style.configure("TScrollbar", background="#e8eaed", troughcolor="#f6f7f9", bordercolor="#dfe3e8", arrowcolor="#555555", lightcolor="#e8eaed", darkcolor="#e8eaed")
            style.configure("TNotebook", background="#f6f7f9", borderwidth=0)
            style.configure("TNotebook.Tab", background="#eceff2", foreground="#333333", padding=(14, 8), borderwidth=0)
            style.map("TNotebook.Tab", background=[("selected", "#ffffff"), ("active", "#f7f7f7")], foreground=[("selected", "#ee4d2d"), ("active", "#333333")])
            try:
                self.root.option_add("*TCombobox*Listbox.background", "#ffffff")
                self.root.option_add("*TCombobox*Listbox.foreground", "#222222")
                self.root.option_add("*TCombobox*Listbox.selectBackground", "#fff0eb")
                self.root.option_add("*TCombobox*Listbox.selectForeground", "#222222")
            except Exception:
                pass
            try:
                self.results_tree.tag_configure("ok", foreground="#168a49", background="#effaf3")
                self.results_tree.tag_configure("fail", foreground="#df3b3b", background="#fff1f1")
                self.results_tree.tag_configure("running", foreground="#b76d00", background="#fff8e7")
            except Exception:
                pass
            try:
                self.mail_tree.tag_configure("ok", foreground="#168a49", background="#effaf3")
                self.mail_tree.tag_configure("fail", foreground="#df3b3b", background="#fff1f1")
                self.mail_tree.tag_configure("running", foreground="#b76d00", background="#fff8e7")
            except Exception:
                pass
            try:
                self.mailfree_tree.tag_configure("ok", foreground="#168a49", background="#effaf3")
                self.mailfree_tree.tag_configure("fail", foreground="#df3b3b", background="#fff1f1")
                self.mailfree_tree.tag_configure("running", foreground="#b76d00", background="#fff8e7")
            except Exception:
                pass
            try:
                self.voucher_preview_tree.tag_configure("valid", foreground="#168a49", background="")
                self.voucher_preview_tree.tag_configure("invalid", foreground="#df3b3b", background="")
                self.voucher_preview_tree.tag_configure("dup", foreground="#b76d00", background="#fff8e7")
                self.voucher_preview_tree.tag_configure("unselected", foreground="#b3b6bb", background="")
            except Exception:
                pass

        # Cập nhật menu đang mở để không bị đổi về màu sáng sau khi chuyển trang.
        try:
            self._show_page(getattr(self, "_active_page", "voucher"), refresh=None)
        except Exception:
            pass

        try:
            self.status_var.set("Giao diện Đen: toàn bộ nền đen / chữ trắng" if dark else "Giao diện Sáng")
        except Exception:
            pass

    def _setup_settings_page(self, parent):
        outer = tk.Frame(parent, bg="#f6f7f9")
        outer.pack(fill="both", expand=True, padx=22, pady=22)
        tk.Label(outer, text="Cài đặt", bg="#f6f7f9", fg="#222222", font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(0, 14))
        card = self._make_card(outer)
        card.pack(fill="x")
        refresh_row = tk.Frame(card, bg="#ffffff")
        refresh_row.pack(fill="x", padx=18, pady=10)
        tk.Label(refresh_row, text="Tự làm mới cả 3 inbox (giây, 0 = tắt)", bg="#ffffff").pack(side="left")
        ttk.Spinbox(refresh_row, from_=0, to=86400, textvariable=self.inbox_refresh_all_var, width=8).pack(side="left", padx=8)
        ttk.Button(refresh_row, text="Áp dụng tất cả", command=self._apply_all_inbox_refresh).pack(side="left")

        rows = [
            # Vubel API Key đã xóa — key được lấy tự động từ server sau khi xác minh license
            ("Proxy / Key", self.proxy_var, False, "Dán key hoặc IP:PORT; http:// và socks5:// được tự nhận."),
            ("Độ trễ AddMail/MailFree", self.delay_var, False, f"Chờ giữa các tài khoản: từ {MIN_DELAY_SECONDS:g} giây, không giới hạn tối đa. Đổi IP chỉnh riêng trong AddMail."),
            ("Mốc hẹn giờ", self.timer_var, False, "Ví dụ: 00:00:00, 12:00:00, 20:00:00"),
        ]
        for idx, (label, var, secret, note) in enumerate(rows):
            row = tk.Frame(card, bg="#ffffff")
            row.pack(fill="x", padx=18, pady=(14 if idx == 0 else 7, 7))
            row.grid_columnconfigure(1, weight=1)
            tk.Label(row, text=label, bg="#ffffff", fg="#333333", width=23, anchor="w", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w")
            ttk.Entry(row, textvariable=var, show="•" if secret else "").grid(row=0, column=1, sticky="ew", padx=(12, 0))
            tk.Label(row, text=note, bg="#ffffff", fg="#888888", font=("Segoe UI", 8)).grid(row=1, column=1, sticky="w", padx=(12, 0), pady=(4, 0))

        proxy_options_row = tk.Frame(card, bg="#ffffff")
        proxy_options_row.pack(fill="x", padx=18, pady=(2, 7))
        proxy_options_row.grid_columnconfigure(1, weight=1)
        tk.Label(
            proxy_options_row,
            text="Kiểu kết nối",
            bg="#ffffff",
            fg="#333333",
            width=23,
            anchor="w",
            font=("Segoe UI", 9, "bold"),
        ).grid(row=0, column=0, sticky="w")
        self.proxy_protocol_combo = ttk.Combobox(
            proxy_options_row,
            textvariable=self.proxy_protocol_var,
            values=PROXY_PROTOCOL_OPTIONS,
            state="readonly",
            width=14,
        )
        self.proxy_protocol_combo.grid(row=0, column=1, sticky="w", padx=(12, 0))
        self.kiotproxy_check_button = ttk.Button(
            proxy_options_row,
            text="↻  Lấy key/proxy hiện tại",
            command=self.check_kiotproxy,
            style="Soft.TButton",
        )
        self.kiotproxy_check_button.grid(row=0, column=2, sticky="e", padx=(10, 0))
        self._register_kiotproxy_check_button(self.kiotproxy_check_button)
        ttk.Checkbutton(
            proxy_options_row,
            text="Đổi IP mới mỗi lần chạy",
            variable=self.kiotproxy_auto_new_ip_var,
            command=self._on_kiotproxy_auto_new_ip_changed,
        ).grid(row=0, column=3, sticky="w", padx=(10, 0))
        tk.Label(
            proxy_options_row,
            textvariable=self.kiotproxy_status_var,
            bg="#ffffff",
            fg="#3976c1",
            font=("Segoe UI", 8, "bold"),
        ).grid(row=1, column=1, columnspan=3, sticky="w", padx=(12, 0), pady=(5, 0))
        tk.Label(
            proxy_options_row,
            textvariable=self.kiotproxy_expiry_var,
            bg="#ffffff",
            fg="#168a49",
            font=("Segoe UI", 8, "bold"),
        ).grid(row=2, column=1, columnspan=3, sticky="w", padx=(12, 0), pady=(2, 0))


        # Background / Nền giao diện
        bg_row = tk.Frame(card, bg="#ffffff")
        bg_row.pack(fill="x", padx=18, pady=(8, 7))
        bg_row.grid_columnconfigure(1, weight=1)
        tk.Label(bg_row, text="Background (Nền)", bg="#ffffff", fg="#333333", width=23, anchor="w",
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w")
        self.background_combo = ttk.Combobox(
            bg_row, textvariable=self.background_mode_var, values=("Sáng", "Đen"), state="readonly", width=18
        )
        self.background_combo.grid(row=0, column=1, sticky="w", padx=(12, 0))
        self.background_combo.bind("<<ComboboxSelected>>", lambda _e: self._apply_background_mode())
        ttk.Button(bg_row, text="Áp dụng", command=self._apply_background_mode, style="Soft.TButton").grid(
            row=0, column=2, sticky="e", padx=(10, 0)
        )
        tk.Label(
            bg_row, text="Chọn Đen: toàn bộ giao diện, card, ô nhập, bảng và log chuyển nền đen; chữ chuyển trắng.",
            bg="#ffffff", fg="#888888", font=("Segoe UI", 8)
        ).grid(row=1, column=1, columnspan=2, sticky="w", padx=(12, 0), pady=(4, 0))

        timer_row = tk.Frame(card, bg="#ffffff")
        timer_row.pack(fill="x", padx=18, pady=(7, 14))
        ttk.Checkbutton(timer_row, text="Bật hẹn giờ tự động", variable=self.use_timer_var).pack(side="left")
        ttk.Checkbutton(timer_row, text="Chuông cảnh báo trước 5 giây", variable=self.alarm_5s_var).pack(side="left", padx=(25, 0))
        ttk.Checkbutton(timer_row, text="Chuông báo khi hoàn tất", variable=self.completion_alarm_var).pack(side="left", padx=(25, 0))
        ttk.Checkbutton(timer_row, text="Chỉ làm mới SPC_ST", variable=self.refresh_only_var).pack(side="left", padx=(25, 0))
        ttk.Button(timer_row, text="Lưu cấu hình", command=self._save_config, style="Accent.TButton").pack(side="right")

    def _apply_all_inbox_refresh(self):
        try:
            seconds = int(self.inbox_refresh_all_var.get())
            if not 0 <= seconds <= 86400:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Thời gian không hợp lệ", "Nhập số nguyên từ 0 đến 86400 giây; 0 để tắt.")
            return
        if self.native_mail_suite is not None:
            for name in ("mailtm", "tinyhost", "outlook"):
                getattr(self.native_mail_suite, name).auto_refresh_seconds.set(str(seconds))
        self._save_config()

    def _register_kiotproxy_check_button(self, button: Any) -> None:
        buttons = getattr(self, "_kiotproxy_check_buttons", None)
        if buttons is None:
            buttons = []
            self._kiotproxy_check_buttons = buttons
        if button not in buttons:
            buttons.append(button)

    def _set_kiotproxy_check_buttons_state(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for button in getattr(self, "_kiotproxy_check_buttons", []):
            try:
                button.configure(state=state)
            except Exception:
                pass

    def _kiotproxy_expiry_text(self, manager: ProxyManager) -> str:
        key_remaining = manager.get_key_remaining_time()
        if manager.key_expiry_timestamp is None:
            key_text = "Hạn key: API chưa trả thời hạn"
        elif key_remaining is not None and key_remaining <= 0:
            key_text = "Hạn key: ĐÃ HẾT HẠN"
        else:
            expiry_text = datetime.fromtimestamp(manager.key_expiry_timestamp).strftime("%d/%m/%Y %H:%M:%S")
            key_text = f"Key còn {format_remaining_duration(key_remaining)} • hết {expiry_text}"

        proxy_remaining = manager.get_proxy_remaining_time()
        if manager.static_proxy:
            proxy_text = f"{manager.proxy_protocol}: proxy tĩnh"
        elif proxy_remaining is None:
            proxy_text = "IP hiện tại: chưa có dữ liệu"
        else:
            proxy_text = f"IP hiện tại còn {format_remaining_duration(proxy_remaining)}"
        return f"{key_text} | {proxy_text}"

    def _kiotproxy_status_text(self, snapshot: dict[str, Any]) -> str:
        protocol = str(snapshot.get("protocol") or PROXY_PROTOCOL_HTTP)
        endpoints: list[str] = []
        http_proxy = str(snapshot.get("http_proxy") or "").strip()
        socks5_proxy = str(snapshot.get("socks5_proxy") or "").strip()
        if http_proxy:
            endpoints.append(f"HTTP {proxy_display_name(http_proxy)}")
        if socks5_proxy:
            endpoints.append(f"SOCKS5 {proxy_display_name(socks5_proxy)}")
        current = str(snapshot.get("current_proxy") or "").strip()
        if not endpoints and current:
            endpoints.append(f"{protocol} {proxy_display_name(current)}")
        endpoint_text = " | ".join(endpoints) if endpoints else "chưa có endpoint"
        return f"KiotProxy: Đã lấy hiện tại ({protocol}) • {endpoint_text}"

    def _on_kiotproxy_auto_new_ip_changed(self) -> None:
        """Lưu ngay lựa chọn để lần mở/chạy sau vẫn giữ chế độ đổi IP mới."""

        self._save_config()
        state = "BẬT" if self.kiotproxy_auto_new_ip_var.get() else "TẮT"
        try:
            self.status_var.set(f"Tự đổi IP mới khi bắt đầu AddMail/MailFree: {state}")
        except Exception:
            pass

    def request_new_kiotproxy(self) -> None:
        if self.mail_running or self.mailfree_running:
            messagebox.showinfo("Đang chạy", "Bật tùy chọn Fake IP mới trước lần chạy kế tiếp, hoặc dừng luồng rồi đổi IP.")
            return
        value = self.proxy_var.get().strip()
        if not value or ProxyManager._looks_like_static_proxy(value):
            messagebox.showinfo("Cần key KiotProxy", "Đổi IP mới cần key KiotProxy; địa chỉ IP:PORT là proxy cố định.")
            return
        self.check_kiotproxy(force_new=True)

    def check_kiotproxy(self, force_new: bool = False) -> None:
        """Kiểm tra key/host:port và lấy proxy hiện tại ở luồng nền."""

        if self._kiotproxy_check_running:
            return
        key_or_proxy = self.proxy_var.get().strip()
        protocol = normalize_proxy_protocol(self.proxy_protocol_var.get())
        self.proxy_protocol_var.set(protocol)
        if not key_or_proxy:
            self._kiotproxy_manager = None
            self.kiotproxy_status_var.set("KiotProxy: Chưa nhập Key hoặc host:port")
            self.kiotproxy_expiry_var.set("Hạn key: chưa kiểm tra | IP hiện tại: chưa có dữ liệu")
            return

        self._save_config()
        self._kiotproxy_check_running = True
        self._kiotproxy_check_generation += 1
        generation = self._kiotproxy_check_generation
        self._set_kiotproxy_check_buttons_state(False)
        self.kiotproxy_status_var.set(f"KiotProxy: Đang lấy proxy hiện tại ({protocol})...")
        self.kiotproxy_expiry_var.set("Hạn key: đang kiểm tra...")
        threading.Thread(
            target=self._check_kiotproxy_worker,
            args=(key_or_proxy, protocol, generation, force_new),
            name="KiotProxyStatus",
            daemon=True,
        ).start()

    def _check_kiotproxy_worker(self, key_or_proxy: str, protocol: str, generation: int, force_new: bool = False) -> None:
        manager: ProxyManager | None = None
        snapshot: dict[str, Any] | None = None
        error_message = ""
        try:
            manager = ProxyManager(key_or_proxy, proxy_protocol=protocol)
            if force_new:
                manager.rotate(force_new=True)
                snapshot = manager.status_snapshot()
            else:
                snapshot = manager.refresh_current()
        except Exception as exc:
            error_message = str(exc)
        self.root.after(
            0,
            self._finish_kiotproxy_check,
            generation,
            key_or_proxy,
            protocol,
            manager,
            snapshot,
            error_message,
        )

    def _finish_kiotproxy_check(
        self,
        generation: int,
        key_or_proxy: str,
        protocol: str,
        manager: ProxyManager | None,
        snapshot: dict[str, Any] | None,
        error_message: str,
    ) -> None:
        if generation != self._kiotproxy_check_generation:
            return
        self._kiotproxy_check_running = False
        self._set_kiotproxy_check_buttons_state(True)
        # Không để response của lần kiểm tra cũ ghi đè cấu hình mới.
        if self.proxy_var.get().strip() != key_or_proxy or normalize_proxy_protocol(self.proxy_protocol_var.get()) != protocol:
            self.kiotproxy_status_var.set("KiotProxy: Cấu hình đã thay đổi; bấm lấy lại để kiểm tra.")
            self.kiotproxy_expiry_var.set("Hạn key: chưa kiểm tra | IP hiện tại: chưa có dữ liệu")
            return
        if error_message or manager is None or snapshot is None:
            self._kiotproxy_manager = None
            self.kiotproxy_status_var.set(f"KiotProxy: Lỗi — {error_message or 'không nhận được response'}")
            self.kiotproxy_expiry_var.set("Hạn key: chưa lấy được | IP hiện tại: chưa có dữ liệu")
            return

        self._kiotproxy_manager = manager
        self.kiotproxy_status_var.set(self._kiotproxy_status_text(snapshot))
        self.kiotproxy_expiry_var.set(self._kiotproxy_expiry_text(manager))

    def _update_kiotproxy_countdown(self) -> None:
        manager = getattr(self, "_kiotproxy_manager", None)
        if manager is not None:
            try:
                self.kiotproxy_expiry_var.set(self._kiotproxy_expiry_text(manager))
            except Exception:
                pass
        try:
            self.root.after(1000, self._update_kiotproxy_countdown)
        except Exception:
            pass

    def _sync_kiotproxy_manager_status(self, manager: ProxyManager) -> None:
        """Đồng bộ endpoint/TTL của manager thực tế đang chạy lên hai khu vực UI."""

        if self.proxy_var.get().strip() != manager.key_or_url:
            return
        if normalize_proxy_protocol(self.proxy_protocol_var.get()) != manager.proxy_protocol:
            return
        self._kiotproxy_manager = manager
        snapshot = manager.status_snapshot()
        self.kiotproxy_status_var.set(self._kiotproxy_status_text(snapshot))
        self.kiotproxy_expiry_var.set(self._kiotproxy_expiry_text(manager))

    def _update_voucher_thread_count(self) -> None:
        count = max(1, len(self.voucher_thread_configs))
        self.voucher_thread_count_var.set(f"{count} luồng")

    def _set_voucher_config_controls_state(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        self.voucher_threads_button.configure(state=state)
        self.voucher_delay_spinbox.configure(state=state)

    def _open_voucher_threads_dialog(self) -> None:
        if self.running:
            messagebox.showinfo("Đang chạy", "Hãy dừng Voucher trước khi sửa cấu hình luồng.")
            return

        configs = list(self.voucher_thread_configs)
        dialog = tk.Toplevel(self.root)
        dialog.title("Cấu hình đa luồng Voucher")
        dialog.geometry("900x510")
        dialog.minsize(720, 440)
        dialog.transient(self.root)
        apply_ryan_window_icon(dialog)

        outer = tk.Frame(dialog, bg="#f6f7f9")
        outer.pack(fill="both", expand=True, padx=18, pady=18)
        tk.Label(
            outer,
            text="Cấu hình luồng Voucher",
            bg="#f6f7f9",
            fg="#222222",
            font=("Segoe UI", 16, "bold"),
        ).pack(anchor="w")
        tk.Label(
            outer,
            text=(
                "Mỗi dòng = 1 luồng. Mỗi luồng dùng Proxy/Key Proxy và số giây đổi IP API riêng. "
                "Để trống Proxy/Key nếu muốn chạy Direct; proxy tĩnh sẽ được giữ nguyên. "
                "Mỗi luồng chạy hết voucher của một cookie rồi mới đổi IP cho cookie tiếp theo. "
                "Nếu chưa đủ số giây đã cấu hình, luồng sẽ chờ trước khi đổi IP."
            ),
            bg="#f6f7f9",
            fg="#666666",
            justify="left",
            wraplength=840,
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(5, 12))

        table_card = self._make_card(outer)
        table_card.pack(fill="both", expand=True)
        table_frame = tk.Frame(table_card, bg="#ffffff")
        table_frame.pack(fill="both", expand=True, padx=1, pady=1)
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        tree = ttk.Treeview(
            table_frame,
            columns=("thread", "proxy", "rotate"),
            show="headings",
            selectmode="browse",
            height=8,
        )
        tree.heading("thread", text="Luồng")
        tree.heading("proxy", text="Proxy / Key Proxy")
        tree.heading("rotate", text="Đổi IP API sau (giây)")
        tree.column("thread", width=75, anchor="center", stretch=False)
        tree.column("proxy", width=540, anchor="w")
        tree.column("rotate", width=180, anchor="center", stretch=False)
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=scrollbar.set)

        editor = tk.Frame(outer, bg="#f6f7f9")
        editor.pack(fill="x", pady=(12, 0))
        editor.grid_columnconfigure(0, weight=1)
        proxy_input = tk.StringVar()
        rotate_input = tk.StringVar(value=f"{DEFAULT_PROXY_ROTATE_SECONDS:g}")
        tk.Label(editor, text="Proxy / Key Proxy (trống = Direct)", bg="#f6f7f9", fg="#444444", font=("Segoe UI", 8, "bold")).grid(row=0, column=0, sticky="w")
        tk.Label(editor, text="Số giây đổi IP API", bg="#f6f7f9", fg="#444444", font=("Segoe UI", 8, "bold")).grid(row=0, column=1, sticky="w", padx=(12, 0))
        proxy_entry = ttk.Entry(editor, textvariable=proxy_input)
        proxy_entry.grid(row=1, column=0, sticky="ew", pady=(5, 0))
        rotate_spin = ttk.Spinbox(
            editor,
            from_=MIN_PROXY_ROTATE_SECONDS,
            to=MAX_PROXY_ROTATE_SECONDS,
            increment=1,
            width=15,
            textvariable=rotate_input,
        )
        rotate_spin.grid(row=1, column=1, sticky="w", padx=(12, 0), pady=(5, 0))

        def refresh_tree(select_index: int | None = None) -> None:
            tree.delete(*tree.get_children())
            for index, item in enumerate(configs, 1):
                tree.insert(
                    "",
                    "end",
                    iid=f"voucher-thread-{index}",
                    values=(index, item.proxy_key or "Direct", f"{item.rotate_seconds:g}"),
                )
            if select_index is not None and 0 <= select_index < len(configs):
                iid = f"voucher-thread-{select_index + 1}"
                tree.selection_set(iid)
                tree.focus(iid)
                tree.see(iid)

        def editor_config() -> VoucherThreadConfig:
            return normalize_voucher_thread_configs([{
                "proxy": proxy_input.get(),
                "rotate_seconds": rotate_input.get(),
            }])[0]

        def selected_index() -> int | None:
            selected = tree.selection()
            return tree.index(selected[0]) if selected else None

        def load_selected(_event=None) -> None:
            index = selected_index()
            if index is None:
                return
            item = configs[index]
            proxy_input.set(item.proxy_key)
            rotate_input.set(f"{item.rotate_seconds:g}")

        def add_config() -> None:
            if len(configs) >= MAX_VOUCHER_THREADS:
                messagebox.showwarning("Đủ số luồng", f"Tối đa {MAX_VOUCHER_THREADS} luồng Voucher.", parent=dialog)
                return
            try:
                configs.append(editor_config())
            except ValueError as exc:
                messagebox.showwarning("Cấu hình không hợp lệ", str(exc), parent=dialog)
                return
            refresh_tree(len(configs) - 1)

        def update_config() -> None:
            index = selected_index()
            if index is None:
                messagebox.showinfo("Chưa chọn", "Hãy chọn một luồng trong bảng để cập nhật.", parent=dialog)
                return
            try:
                configs[index] = editor_config()
            except ValueError as exc:
                messagebox.showwarning("Cấu hình không hợp lệ", str(exc), parent=dialog)
                return
            refresh_tree(index)

        def delete_config() -> None:
            index = selected_index()
            if index is None:
                messagebox.showinfo("Chưa chọn", "Hãy chọn một luồng trong bảng để xóa.", parent=dialog)
                return
            if len(configs) == 1:
                messagebox.showinfo("Giữ một luồng", "Voucher luôn cần ít nhất một luồng. Có thể để Proxy trống để chạy Direct.", parent=dialog)
                return
            del configs[index]
            refresh_tree(min(index, len(configs) - 1))

        def apply_configs() -> None:
            try:
                index = selected_index()
                if index is not None:
                    # "Áp dụng" cũng lưu nội dung đang sửa, không bắt buộc bấm
                    # thêm nút "Cập nhật dòng chọn" trước đó.
                    configs[index] = editor_config()
                normalized = normalize_voucher_thread_configs(configs)
            except ValueError as exc:
                messagebox.showwarning("Cấu hình không hợp lệ", str(exc), parent=dialog)
                return
            self.voucher_thread_configs = normalized
            self._update_voucher_thread_count()
            self._save_config()
            dialog.destroy()

        tree.bind("<<TreeviewSelect>>", load_selected)
        buttons = tk.Frame(outer, bg="#f6f7f9")
        buttons.pack(fill="x", pady=(12, 0))
        ttk.Button(buttons, text="＋ Thêm luồng", command=add_config, style="Soft.TButton").pack(side="left")
        ttk.Button(buttons, text="Cập nhật dòng chọn", command=update_config, style="Soft.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="Xóa dòng chọn", command=delete_config, style="Danger.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="Hủy", command=dialog.destroy, style="Soft.TButton").pack(side="right")
        ttk.Button(buttons, text="Áp dụng", command=apply_configs, style="Accent.TButton").pack(side="right", padx=(0, 8))

        refresh_tree(0)
        load_selected()
        proxy_entry.focus_set()
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        dialog.grab_set()

    def _update_input_counts(self, _event=None):
        try:
            codes = parse_text_list(self.codes_text.get("1.0", "end"), MAX_BATCH_CODES, "mã voucher")
            codes, duplicate_count = dedupe_voucher_entries(codes)
            if duplicate_count:
                self.code_count_var.set(f"{len(codes)} (-{duplicate_count} trùng)")
            else:
                self.code_count_var.set(str(len(codes)))
        except Exception:
            self.code_count_var.set("!")
        try:
            cookies = parse_line_list(self.cookies_text.get("1.0", "end"), MAX_BATCH_COOKIES, "cookie/tài khoản")
            self.cookie_count_var.set(str(len(cookies)))
        except Exception:
            self.cookie_count_var.set("!")

    def _update_voucher_preview(self, _event=None):
        """Bóc tách từng dòng đã dán vào ô mã voucher để hiển thị bảng xem trước
        (Loại/Mã Code/Hợp lệ), kèm ô Chọn và Ghi chú tự đặt. PromotionId/Signature/
        Link Gốc chỉ được lưu nguyên văn trong nội dung dòng nhập (và trong config
        khi lưu), KHÔNG hiển thị thành cột riêng trên bảng này."""
        tree = getattr(self, "voucher_preview_tree", None)
        if tree is None:
            return
        raw_text = self.codes_text.get("1.0", "end").replace("\r", "")
        lines = [line.strip() for line in raw_text.split("\n") if line.strip()]
        self._voucher_preview_raw_lines = lines

        tree.delete(*tree.get_children())
        mask = self.voucher_mask_code_var.get()
        seen_keys: set[str] = set()
        selected_count = 0
        for i, raw in enumerate(lines, 1):
            entry = parse_voucher_entry(raw)
            key = _voucher_entry_dedupe_key(entry, raw)
            is_dup = key in seen_keys
            seen_keys.add(key)

            type_label = "🔗 Theo Link" if entry["is_link"] else "🔤 Theo Mã"
            code_value = str(entry.get("code") or "")
            code_display = ("•" * min(len(code_value), 10)) if (mask and code_value) else code_value
            note = self.voucher_notes.get(raw, "")
            is_selected = self.voucher_selected.get(raw, True)
            if is_selected:
                selected_count += 1
            select_display = "☑" if is_selected else "☐"

            if is_dup:
                valid_display, tag = "TRÙNG", "dup"
            elif entry["is_valid"]:
                valid_display, tag = "✅ HỢP LỆ", "valid"
            else:
                valid_display, tag = "❌ THIẾU", "invalid"
            if not is_selected:
                tag = "unselected"

            tree.insert(
                "", "end", iid=str(i),
                values=(select_display, i, note, type_label, code_display, valid_display),
                tags=(tag,),
            )
        self._apply_voucher_code_column_visibility()
        self.voucher_selected_count_var.set(f"Đã chọn: {selected_count}/{len(lines)}")

    def _apply_voucher_code_column_visibility(self):
        tree = getattr(self, "voucher_preview_tree", None)
        if tree is None:
            return
        all_columns = ("select", "stt", "note", "type", "code", "valid")
        if self.voucher_hide_code_col_var.get():
            tree["displaycolumns"] = tuple(c for c in all_columns if c != "code")
        else:
            tree["displaycolumns"] = all_columns

    def _resolve_voucher_preview_column(self, tree, x: int) -> str | None:
        col_id = tree.identify_column(x)
        displaycols = tree["displaycolumns"]
        if displaycols == "#all":
            displaycols = tree["columns"]
        try:
            col_index = int(col_id.replace("#", "")) - 1
            return displaycols[col_index]
        except (ValueError, IndexError):
            return None

    def _on_voucher_preview_click(self, event):
        """Tích/bỏ tích ô Chọn của 1 dòng, hoặc bấm vào tiêu đề cột Chọn để đảo chọn tất cả."""
        tree = self.voucher_preview_tree
        region = tree.identify("region", event.x, event.y)
        column_name = self._resolve_voucher_preview_column(tree, event.x)
        if column_name != "select":
            return

        if region == "heading":
            any_unselected = any(
                not self.voucher_selected.get(line, True) for line in self._voucher_preview_raw_lines
            )
            for line in self._voucher_preview_raw_lines:
                self.voucher_selected[line] = any_unselected
            self._update_voucher_preview()
            return

        if region != "cell":
            return
        row_id = tree.identify_row(event.y)
        if not row_id:
            return
        try:
            row_index = int(row_id) - 1
        except ValueError:
            return
        if not 0 <= row_index < len(self._voucher_preview_raw_lines):
            return
        raw_line = self._voucher_preview_raw_lines[row_index]
        self.voucher_selected[raw_line] = not self.voucher_selected.get(raw_line, True)
        self._update_voucher_preview()

    def _select_all_voucher_preview(self):
        for line in self._voucher_preview_raw_lines:
            self.voucher_selected[line] = True
        self._update_voucher_preview()

    def _deselect_all_voucher_preview(self):
        for line in self._voucher_preview_raw_lines:
            self.voucher_selected[line] = False
        self._update_voucher_preview()

    def _on_voucher_preview_double_click(self, event):
        """Cho phép sửa ô Ghi chú tại chỗ (không gửi lên Shopee, chỉ để tự nhớ)."""
        tree = self.voucher_preview_tree
        if tree.identify("region", event.x, event.y) != "cell":
            return
        row_id = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)
        if not row_id:
            return

        displaycols = tree["displaycolumns"]
        if displaycols == "#all":
            displaycols = tree["columns"]
        try:
            col_index = int(col_id.replace("#", "")) - 1
            column_name = displaycols[col_index]
        except (ValueError, IndexError):
            return
        if column_name != "note":
            return

        try:
            row_index = int(row_id) - 1
        except ValueError:
            return
        if not 0 <= row_index < len(self._voucher_preview_raw_lines):
            return
        raw_line = self._voucher_preview_raw_lines[row_index]

        bbox = tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, width, height = bbox

        edit_var = tk.StringVar(value=self.voucher_notes.get(raw_line, ""))
        editor = ttk.Entry(tree, textvariable=edit_var)
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()
        editor.select_range(0, "end")

        def commit(_evt=None):
            value = edit_var.get().strip()
            if value:
                self.voucher_notes[raw_line] = value
            else:
                self.voucher_notes.pop(raw_line, None)
            editor.destroy()
            self._update_voucher_preview()

        def cancel(_evt=None):
            editor.destroy()

        editor.bind("<Return>", commit)
        editor.bind("<FocusOut>", commit)
        editor.bind("<Escape>", cancel)

    def _update_summary(self):
        total = int(getattr(self, "total_count", 0) or 0)
        processed = int(getattr(self, "processed_count", 0) or 0)
        ok = int(getattr(self, "ok_count", 0) or 0)
        fail = int(getattr(self, "fail_count", 0) or 0)
        already_saved = int(getattr(self, "already_saved_count", 0) or 0)
        new_saved = int(getattr(self, "new_saved_count", 0) or 0)
        self.summary_total_var.set(str(total))
        self.summary_ok_var.set(str(ok))
        self.summary_fail_var.set(str(fail))
        self.summary_already_var.set(str(already_saved))
        self.summary_new_saved_var.set(str(new_saved))
        self.summary_spcst_var.set(str(ok if getattr(self, "batch_refresh_only", False) else 0))
        pct = int(round(processed * 100 / total)) if total else 0
        self.summary_progress_var.set(f"{processed} / {total} ({pct}%)")
        try:
            self.summary_progress.configure(value=pct)
        except Exception:
            pass

    def _copy_log(self):
        try:
            text = self.details.get("1.0", "end-1c")
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.status_var.set("Đã sao chép log vào clipboard.")
        except Exception as exc:
            messagebox.showerror("Lỗi", f"Không thể sao chép log: {exc}")


    def _setup_read_mail_page(self, parent):
        """Trang Đọc Mail native: 3 tab con chạy trực tiếp trong Tkinter, không HTML/browser."""
        if NativeMailSuiteFrame is None:
            outer = tk.Frame(parent, bg="#f6f7f9")
            outer.pack(fill="both", expand=True, padx=22, pady=22)
            tk.Label(outer, text="Đọc Mail", bg="#f6f7f9", fg="#222222", font=("Segoe UI", 18, "bold")).pack(anchor="w")
            tk.Label(outer, text="Không tải được module mail_native_suite.py. Hãy giữ file này cùng thư mục tool.", bg="#f6f7f9", fg="#df3b3b", font=("Segoe UI", 10)).pack(anchor="w", pady=12)
            return
        try:
            self.native_mail_suite = NativeMailSuiteFrame(parent)
            self.native_mail_suite.pack(fill="both", expand=True)
        except Exception as exc:
            outer = tk.Frame(parent, bg="#f6f7f9")
            outer.pack(fill="both", expand=True, padx=22, pady=22)
            tk.Label(outer, text="Lỗi khởi tạo Đọc Mail", bg="#f6f7f9", fg="#df3b3b", font=("Segoe UI", 16, "bold")).pack(anchor="w")
            tk.Label(outer, text=str(exc), bg="#f6f7f9", fg="#555555", wraplength=900, justify="left").pack(anchor="w", pady=10)

    def _setup_mail_tab(self, parent):
        """Hai luồng Vubel độc lập: AddMail nhập tay và MailFree random."""
        outer = tk.Frame(parent, bg="#f6f7f9")
        outer.pack(fill="both", expand=True, padx=18, pady=18)
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=1)

        title_row = tk.Frame(outer, bg="#f6f7f9")
        title_row.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        tk.Label(title_row, text="Thêm Mail Shopee • Ryan Nguyễn • v3.18", bg="#f6f7f9", fg="#222222", font=("Segoe UI", 18, "bold")).pack(side="left")
        tk.Label(title_row, text="AddMail và MailFree chạy ở hai tab riêng theo Vubel API", bg="#f6f7f9", fg="#777777", font=("Segoe UI", 9)).pack(side="left", padx=14, pady=(6, 0))

        self.mail_flow_notebook = ttk.Notebook(outer)
        self.mail_flow_notebook.grid(row=1, column=0, sticky="nsew")
        self.addmail_tab = tk.Frame(self.mail_flow_notebook, bg="#f6f7f9")
        self.mailfree_tab = tk.Frame(self.mail_flow_notebook, bg="#f6f7f9")
        self.mail_flow_notebook.add(self.addmail_tab, text="AddMail — Email nhập tay")
        self.mail_flow_notebook.add(self.mailfree_tab, text="MailFree — Tạo & đọc inbox")
        self._setup_addmail_flow_tab(self.addmail_tab)
        self._setup_mailfree_flow_tab(self.mailfree_tab)

    def _setup_addmail_flow_tab(self, parent):
        panel = tk.Frame(parent, bg="#f6f7f9")
        panel.pack(fill="both", expand=True, padx=4, pady=8)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=0)
        panel.grid_rowconfigure(2, weight=1)

        input_row = tk.Frame(panel, bg="#f6f7f9")
        input_row.grid(row=0, column=0, sticky="ew")
        input_row.grid_columnconfigure(0, weight=1)
        input_row.grid_columnconfigure(1, weight=1)

        acc_card = self._make_card(input_row)
        acc_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        tk.Label(acc_card, text="Tài khoản / Cookie (Tự nhận SPC_F / SPC_ST)", bg="#ffffff", fg="#333333", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=16, pady=(13, 6))
        self.mail_accounts_text = scrolledtext.ScrolledText(acc_card, height=8, font=("Consolas", 9), bd=0, relief="flat", wrap="none", background="#ffffff")
        self.mail_accounts_text.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        mail_card = self._make_card(input_row)
        mail_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        tk.Label(mail_card, text="Danh sách Email", bg="#ffffff", fg="#333333", font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=16, pady=(13, 6))
        self.mail_emails_text = scrolledtext.ScrolledText(mail_card, height=8, font=("Consolas", 9), bd=0, relief="flat", wrap="none", background="#ffffff")
        self.mail_emails_text.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        proxy_card = self._make_card(panel)
        proxy_card.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        proxy_card.grid_columnconfigure(1, weight=1)
        proxy_head = tk.Frame(proxy_card, bg="#ffffff")
        proxy_head.grid(row=0, column=0, columnspan=5, sticky="ew", padx=15, pady=(10, 4))
        tk.Label(
            proxy_head,
            text="⚡ KiotProxy cho AddMail / MailFree",
            bg="#ffffff",
            fg="#333333",
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left")
        tk.Label(
            proxy_head,
            text="Tự nhận key hoặc proxy • IP:PORT dùng giao thức bên cạnh",
            bg="#ffffff",
            fg="#888888",
            font=("Segoe UI", 8),
        ).pack(side="left", padx=(12, 0))
        tk.Label(proxy_card, text="Key / Proxy", bg="#ffffff", fg="#555555", font=("Segoe UI", 8, "bold")).grid(
            row=1, column=0, sticky="w", padx=(15, 8), pady=(3, 5)
        )
        ttk.Entry(proxy_card, textvariable=self.proxy_var).grid(row=1, column=1, sticky="ew", pady=(3, 5))
        self.mail_proxy_protocol_combo = ttk.Combobox(
            proxy_card,
            textvariable=self.proxy_protocol_var,
            values=PROXY_PROTOCOL_OPTIONS,
            state="readonly",
            width=11,
        )
        self.mail_proxy_protocol_combo.grid(row=1, column=2, sticky="w", padx=(8, 0), pady=(3, 5))
        mail_kiot_button = ttk.Button(
            proxy_card,
            text="↻  Lấy hiện tại",
            command=self.check_kiotproxy,
            style="Soft.TButton",
        )
        mail_kiot_button.grid(row=1, column=3, sticky="e", padx=(8, 5), pady=(3, 5))
        self._register_kiotproxy_check_button(mail_kiot_button)
        new_ip_button = ttk.Button(proxy_card, text="Fake IP mới", command=self.request_new_kiotproxy, style="Soft.TButton")
        new_ip_button.grid(row=5, column=1, sticky="w", pady=(0, 8))
        self._register_kiotproxy_check_button(new_ip_button)
        ttk.Checkbutton(
            proxy_card,
            text="Đổi IP mới mỗi lần chạy",
            variable=self.kiotproxy_auto_new_ip_var,
            command=self._on_kiotproxy_auto_new_ip_changed,
        ).grid(row=1, column=4, sticky="w", padx=(5, 15), pady=(3, 5))
        tk.Label(
            proxy_card, text="Đổi IP AddMail", bg="#ffffff", fg="#555555",
            font=("Segoe UI", 8, "bold"),
        ).grid(row=2, column=0, sticky="w", padx=(15, 8), pady=(3, 7))
        rotation_row = tk.Frame(proxy_card, bg="#ffffff")
        rotation_row.grid(row=2, column=1, columnspan=4, sticky="ew", padx=(0, 15), pady=(3, 7))
        self.addmail_rotate_seconds_entry = ttk.Entry(
            rotation_row, textvariable=self.addmail_rotate_seconds_var, width=12,
        )
        self.addmail_rotate_seconds_entry.pack(side="left")
        tk.Label(
            rotation_row,
            text="giây • Số > 0, không giới hạn tối đa; trống: theo nhà cung cấp.",
            bg="#ffffff", fg="#666666", font=("Segoe UI", 8),
        ).pack(side="left", padx=(8, 0))
        tk.Label(
            proxy_card,
            textvariable=self.kiotproxy_status_var,
            bg="#ffffff",
            fg="#3976c1",
            font=("Segoe UI", 8, "bold"),
        ).grid(row=3, column=1, columnspan=4, sticky="w", pady=(0, 2))
        tk.Label(
            proxy_card,
            textvariable=self.kiotproxy_expiry_var,
            bg="#ffffff",
            fg="#168a49",
            font=("Segoe UI", 8, "bold"),
        ).grid(row=4, column=1, columnspan=4, sticky="w", pady=(0, 9))

        result_card = self._make_card(panel)
        result_card.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        result_card.grid_rowconfigure(2, weight=1)
        result_card.grid_columnconfigure(0, weight=1)

        opts = tk.Frame(result_card, bg="#ffffff")
        opts.grid(row=0, column=0, sticky="ew", padx=15, pady=(12, 8))
        tk.Label(opts, text="LUỒNG CỐ ĐỊNH: chỉ AddMail email đã dán — không gọi MailFree, không tạo mail random.", bg="#ffffff", fg="#168a49", font=("Segoe UI", 8, "bold")).pack(side="left")
        self.mail_proxy_status_var = tk.StringVar(value="Proxy Shopee AddMail: Direct")
        tk.Label(opts, textvariable=self.mail_proxy_status_var, bg="#ffffff", fg="#3976c1", font=("Segoe UI", 8, "bold")).pack(side="right")

        btns = tk.Frame(result_card, bg="#ffffff")
        btns.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 9))
        for i in range(4):
            btns.grid_columnconfigure(i, weight=1)
        self.mail_run_button = ttk.Button(btns, text="▶  Bắt đầu thêm mail", command=self.start_add_mail, style="Accent.TButton")
        self.mail_run_button.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.mail_stop_button = ttk.Button(btns, text="⏹  Dừng", command=self.stop_add_mail, state="disabled", style="Danger.TButton")
        self.mail_stop_button.grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(btns, text="⌫  Xóa kết quả", command=self.clear_mail_results, style="Danger.TButton").grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(btns, text="▤  Xuất Excel Mail", command=self.export_mail_excel, style="Soft.TButton").grid(row=0, column=3, sticky="ew", padx=(4, 0))

        tree_frame = tk.Frame(result_card, bg="#ffffff")
        tree_frame.grid(row=2, column=0, sticky="nsew", padx=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.mail_tree = ttk.Treeview(tree_frame, columns=("stt", "account", "email", "proxy", "status", "message"), show="headings")
        for col, title in (("stt", "STT"), ("account", "Tài khoản"), ("email", "Email"), ("proxy", "Proxy Shopee AddMail"), ("status", "Kết quả"), ("message", "Thông báo")):
            self.mail_tree.heading(col, text=title)
        self.mail_tree.column("stt", width=50, anchor="center", stretch=False)
        self.mail_tree.column("account", width=280)
        self.mail_tree.column("email", width=240)
        self.mail_tree.column("proxy", width=180, anchor="center")
        self.mail_tree.column("status", width=125, anchor="center", stretch=False)
        self.mail_tree.column("message", width=420)
        self.mail_tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.mail_tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.mail_tree.configure(yscrollcommand=scroll.set)
        self.mail_tree.tag_configure("ok", foreground="#168a49", background="#effaf3")
        self.mail_tree.tag_configure("fail", foreground="#df3b3b", background="#fff1f1")
        self.mail_tree.tag_configure("running", foreground="#b76d00", background="#fff8e7")

        self.mail_status_var = tk.StringVar(value="Sẵn sàng.")
        tk.Label(result_card, textvariable=self.mail_status_var, bg="#ffffff", fg="#666666", font=("Segoe UI", 8)).grid(row=3, column=0, sticky="w", padx=15, pady=(7, 12))

    def _setup_mailfree_flow_tab(self, parent):
        panel = tk.Frame(parent, bg="#f6f7f9")
        panel.pack(fill="both", expand=True, padx=4, pady=8)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=1)

        acc_card = self._make_card(panel)
        acc_card.grid(row=0, column=0, sticky="ew")
        tk.Label(
            acc_card,
            text="Tài khoản Shopee để ghép dòng khi chuyển sang AddMail (MailFree KHÔNG gửi/kiểm tra cookie)",
            bg="#ffffff",
            fg="#333333",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=16, pady=(13, 8))
        self.mailfree_accounts_text = scrolledtext.ScrolledText(acc_card, height=8, font=("Consolas", 9), bd=0, relief="flat", wrap="none", background="#ffffff")
        self.mailfree_accounts_text.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        result_card = self._make_card(panel)
        result_card.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        result_card.grid_rowconfigure(2, weight=1)
        result_card.grid_columnconfigure(0, weight=1)

        opts = tk.Frame(result_card, bg="#ffffff")
        opts.grid(row=0, column=0, sticky="ew", padx=15, pady=(12, 8))
        tk.Label(
            opts,
            text="MAILFREE: chỉ tạo và đọc email — không đăng nhập Shopee, không chạy AddMail.",
            bg="#ffffff",
            fg="#168a49",
            font=("Segoe UI", 8, "bold"),
        ).pack(side="left")
        self.mailfree_proxy_status_var = tk.StringVar(value="Proxy MailFree: Direct")
        tk.Label(opts, textvariable=self.mailfree_proxy_status_var, bg="#ffffff", fg="#3976c1", font=("Segoe UI", 8, "bold")).pack(side="right")

        btns = tk.Frame(result_card, bg="#ffffff")
        btns.grid(row=1, column=0, sticky="ew", padx=15, pady=(0, 9))
        for i in range(6):
            btns.grid_columnconfigure(i, weight=1)
        self.mailfree_run_button = ttk.Button(btns, text="▶  Bắt đầu MailFree", command=self.start_mailfree, style="Accent.TButton")
        self.mailfree_run_button.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.mailfree_stop_button = ttk.Button(btns, text="⏹  Dừng", command=self.stop_mailfree, state="disabled", style="Danger.TButton")
        self.mailfree_stop_button.grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(
            btns,
            text="⇢  Chuyển sang AddMail",
            command=self.transfer_mailfree_to_addmail,
            style="Soft.TButton",
        ).grid(row=0, column=2, sticky="ew", padx=4)
        self.mailfree_read_button = ttk.Button(
            btns,
            text="✉  Đọc MailFree",
            command=self.read_selected_mailfree,
            style="Soft.TButton",
        )
        self.mailfree_read_button.grid(row=0, column=3, sticky="ew", padx=4)
        ttk.Button(btns, text="⌫  Xóa kết quả", command=self.clear_mailfree_results, style="Danger.TButton").grid(row=0, column=4, sticky="ew", padx=4)
        ttk.Button(btns, text="▤  Xuất Excel", command=self.export_mailfree_excel, style="Soft.TButton").grid(row=0, column=5, sticky="ew", padx=(4, 0))

        tree_frame = tk.Frame(result_card, bg="#ffffff")
        tree_frame.grid(row=2, column=0, sticky="nsew", padx=1)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self.mailfree_tree = ttk.Treeview(tree_frame, columns=("stt", "account", "email", "proxy", "status", "message"), show="headings")
        for col, title in (("stt", "STT"), ("account", "Tài khoản"), ("email", "MailFree Vubel tạo"), ("proxy", "Proxy MailFree"), ("status", "Kết quả"), ("message", "Thông báo")):
            self.mailfree_tree.heading(col, text=title)
        self.mailfree_tree.column("stt", width=50, anchor="center", stretch=False)
        self.mailfree_tree.column("account", width=280)
        self.mailfree_tree.column("email", width=240)
        self.mailfree_tree.column("proxy", width=180, anchor="center")
        self.mailfree_tree.column("status", width=125, anchor="center", stretch=False)
        self.mailfree_tree.column("message", width=420)
        self.mailfree_tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.mailfree_tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.mailfree_tree.configure(yscrollcommand=scroll.set)
        self.mailfree_tree.bind("<Double-1>", lambda _event: self.read_selected_mailfree())
        self.mailfree_tree.tag_configure("ok", foreground="#168a49", background="#effaf3")
        self.mailfree_tree.tag_configure("fail", foreground="#df3b3b", background="#fff1f1")
        self.mailfree_tree.tag_configure("running", foreground="#b76d00", background="#fff8e7")

        self.mailfree_status_var = tk.StringVar(value="Sẵn sàng.")
        tk.Label(result_card, textvariable=self.mailfree_status_var, bg="#ffffff", fg="#666666", font=("Segoe UI", 8)).grid(row=3, column=0, sticky="w", padx=15, pady=(7, 12))

    def start_add_mail(self):
        if self.mail_running:
            return
        if self.mailfree_running:
            messagebox.showwarning("MailFree đang chạy", "Hãy đợi hoặc dừng tab MailFree trước khi chạy AddMail.")
            return
        try:
            accounts = parse_line_list(self.mail_accounts_text.get("1.0", "end"), MAX_BATCH_COOKIES, "tài khoản")
            emails = parse_line_list(self.mail_emails_text.get("1.0", "end"), MAX_BATCH_COOKIES, "email")
            delay = self._read_delay()
            rotate_seconds = parse_proxy_rotate_seconds(self.addmail_rotate_seconds_var.get())
        except ValueError as exc:
            messagebox.showwarning("Dữ liệu không hợp lệ", str(exc))
            return
        if not accounts:
            messagebox.showwarning("Thiếu dữ liệu", "Cần nhập ít nhất 1 tài khoản.")
            return
        api_key = _VUBEL_KEY_RAM
        proxy_key = self.proxy_var.get().strip()
        proxy_protocol = normalize_proxy_protocol(self.proxy_protocol_var.get())
        manual_proxy = ""
        manual_proxy_protocol_value = proxy_protocol
        auto_new_ip = bool(self.kiotproxy_auto_new_ip_var.get())
        if not api_key:
            messagebox.showwarning("Chưa xác thực", "Vubel key chưa sẵn sàng. Hãy khởi động lại và kích hoạt license.")
            return
        if not emails:
            messagebox.showwarning("Thiếu Email", "Hãy dán Email cần AddMail trong tab này.")
            return
        if len(emails) != len(accounts):
            messagebox.showwarning("Sai số lượng", "Số dòng Email phải bằng số dòng tài khoản.")
            return

        self._save_config()
        self.clear_mail_results()
        tasks = list(zip(accounts, emails))
        for idx, (account, email) in enumerate(tasks, 1):
            preview = f"{account[:18]}...{account[-12:]}" if len(account) > 34 else account
            self.mail_tree.insert("", "end", iid=f"addmail-{idx}", values=(idx, preview, email, "Direct", "CHỜ", "Chưa xử lý"))

        self.mail_running = True
        self.mail_stop_event.clear()
        self.mail_run_button.configure(state="disabled")
        self.mail_stop_button.configure(state="normal")
        self.mail_accounts_text.configure(state="disabled")
        self.mail_emails_text.configure(state="disabled")
        self.addmail_rotate_seconds_entry.configure(state="disabled")
        self.mail_status_var.set(f"Đang xử lý 0/{len(tasks)}...")
        threading.Thread(
            target=self._mail_worker,
            args=(tasks, delay, api_key, proxy_key, proxy_protocol, auto_new_ip, manual_proxy, manual_proxy_protocol_value, rotate_seconds),
            daemon=True,
        ).start()

    def _mail_worker(
        self,
        tasks: list[tuple[str, str]],
        delay: float,
        api_key: str,
        proxy_key: str,
        proxy_protocol: str = PROXY_PROTOCOL_HTTP,
        auto_new_ip: bool = False,
        manual_proxy: str = "",
        manual_proxy_protocol_value: str = PROXY_PROTOCOL_HTTP,
        rotate_seconds: float | None = None,
    ):
        using_manual_proxy = bool(manual_proxy)
        if using_manual_proxy:
            runtime_protocol = infer_manual_proxy_protocol(manual_proxy, manual_proxy_protocol_value)
            pm = ProxyManager(manual_proxy, rotate_interval=rotate_seconds, proxy_protocol=runtime_protocol)
        else:
            pm = ProxyManager(proxy_key, rotate_interval=rotate_seconds, proxy_protocol=proxy_protocol) if proxy_key else None
        if pm is not None and not using_manual_proxy:
            self._kiotproxy_manager = pm
        processed = 0
        ok_count = 0

        for idx, (account, email) in enumerate(tasks, 1):
            if self.mail_stop_event.is_set():
                break

            proxy_url = None
            proxy_label = "Direct"
            if pm:
                try:
                    if pm.needs_rotation():
                        if using_manual_proxy:
                            status = "đang dùng proxy thủ công..."
                        else:
                            status = "đang lấy IP mới..." if auto_new_ip and pm.current_proxy is None and not pm.static_proxy else "đang đổi IP..."
                        self.root.after(0, lambda s=status: self.mail_proxy_status_var.set(f"Proxy Shopee AddMail: {s}"))
                        proxy_url = acquire_proxy_until_ready(
                            pm, self.mail_stop_event,
                            lambda text: self.root.after(0, self.mail_proxy_status_var.set, text),
                            force_new=(not using_manual_proxy) and auto_new_ip and pm.current_proxy is None,
                        )
                        if proxy_url is None:
                            break
                    else:
                        proxy_url = pm.current_proxy
                    proxy_label = self._format_mail_proxy(proxy_url)
                    if not using_manual_proxy:
                        self.root.after(0, self._sync_kiotproxy_manager_status, pm)
                    source = "Proxy tay" if using_manual_proxy else "KiotProxy"
                    self.root.after(0, lambda p=proxy_label, s=source: self.mail_proxy_status_var.set(f"Proxy Shopee AddMail: {s} • {p}"))
                except KiotProxyKeyExpiredError as exc:
                    self.root.after(0, self._record_mail_result, idx, account, email, TestResult(False, str(exc), None), "KEY HẾT HẠN")
                    processed += 1
                    self.mail_stop_event.set()
                    break
                except Exception as exc:
                    result = TestResult(False, f"Lỗi Proxy Shopee AddMail: {exc}", None)
                    actual_email = email
                    self.root.after(0, self._record_mail_result, idx, account, actual_email, result, "LỖI PROXY")
                    processed += 1
                    if idx < len(tasks) and self.mail_stop_event.wait(delay):
                        break
                    continue
            else:
                self.root.after(0, lambda: self.mail_proxy_status_var.set("Proxy Shopee AddMail: Direct (IP máy)"))

            self.root.after(0, self._mark_mail_running, idx, proxy_label, account)

            # QUAN TRỌNG: KHÔNG route request api.vubel.store qua proxy Shopee.
            # Theo tài liệu Vubel, proxy phải nằm trong body để backend Vubel dùng
            # khi thao tác Shopee. Route chính request Vubel qua proxy có thể làm
            # proxy CONNECT trả HTTP 500 trước khi nghiệp vụ addmail chạy.
            try:
                with build_httpx_client(proxy_url=None) as addmail_client:
                    def handle_mail_ivs(session_id, methods, ivs_proxy_url):
                        evt = threading.Event()
                        container: dict[str, Any] = {}

                        def _show_ivs():
                            self.mail_status_var.set("SPC_F cần xác minh IVS/OTP để lấy SPC_ST...")
                            IVSDialog(
                                self.root,
                                session_id,
                                methods,
                                api_key,
                                ivs_proxy_url,
                                container,
                                evt,
                            )

                        self.root.after(0, _show_ivs)
                        evt.wait()
                        return container.get("spc_st")

                    result, actual_email = vubel_add_mail_manual(
                        addmail_client,
                        account,
                        email,
                        api_key,
                        proxy_url,
                        handle_mail_ivs,
                    )
                    # Giữ email do vubel_add_mail đã kiểm chứng. Nếu backend trả
                    # mail khác, hàm trả THẤT BẠI và email đó để UI báo rõ thay
                    # vì che lỗi bằng cách ghi đè lại email người dùng dán.
            except Exception as exc:
                result = TestResult(False, f"Không tạo được HTTP client tới Vubel: {exc}", None)
                actual_email = email

            if result.success:
                ok_count += 1
            processed += 1
            self.root.after(0, self._record_mail_result, idx, account, actual_email, result, proxy_label)
            self.root.after(
                0,
                lambda p=processed, t=len(tasks), o=ok_count: self.mail_status_var.set(
                    f"Tiến độ: {p}/{t} | Thành công: {o} | Thất bại: {p-o}"
                ),
            )
            if idx < len(tasks) and self.mail_stop_event.wait(delay):
                break

        self.root.after(0, self._finish_mail_batch, self.mail_stop_event.is_set(), processed, len(tasks), ok_count)

    def _format_mail_proxy(self, proxy_url: str | None) -> str:
        """Hiện IP:port, không lộ user/pass proxy trên giao diện/log."""
        raw = str(proxy_url or "").strip()
        if not raw:
            return "Direct"
        value = re.sub(r"^[A-Za-z0-9+.-]+://", "", raw)
        if "@" in value:
            value = value.rsplit("@", 1)[1]
        return value[:80]

    def _mark_mail_running(self, idx: int, proxy_label: str = "Direct", account: str = ""):
        iid = f"addmail-{idx}"
        values = list(self.mail_tree.item(iid, "values"))
        if values:
            values[3] = proxy_label
            values[4] = "ĐANG XỬ LÝ"
            parsed = _parse_addmail_account(account)
            if parsed.get("kind") == "spc_f":
                values[5] = "SPC_F → REFRESH SPC_ST → ADDMAIL"
            else:
                values[5] = "SPC_ST → ADDMAIL"
            self.mail_tree.item(iid, values=values, tags=("running",))
            self.mail_tree.see(iid)

    def _record_mail_result(self, idx: int, account: str, email: str, result: TestResult, proxy_used: str = "Direct"):
        iid = f"addmail-{idx}"
        values = list(self.mail_tree.item(iid, "values"))
        if values:
            values[2] = email or values[2]
            values[3] = proxy_used
            values[4] = "THÀNH CÔNG" if result.success else "THẤT BẠI"
            values[5] = result.message
            self.mail_tree.item(iid, values=values, tags=(("ok",) if result.success else ("fail",)))
        source_account = _parse_addmail_account(account)
        resolved_account = _parse_addmail_account(result.resolved_cookie or account)
        username, password, spcf = (
            source_account["username"], source_account["password"],
            source_account["spc_f"] or resolved_account["spc_f"],
        )
        if result.success and email and "@" in str(email):
            # Tự lưu email đã AddMail thành công vào kho Đọc Mail. Chỉ lưu liên kết,
            # không ghi đè password/token mailbox đã có.
            try:
                meta = {"http_status": result.status_code, "proxy": proxy_used}
                if register_linked_email_global is not None:
                    register_linked_email_global(str(email), shopee_user=username, source="vubel_addmail", meta=meta)
                if self.native_mail_suite is not None and hasattr(self.native_mail_suite, "register_linked_email"):
                    self.native_mail_suite.register_linked_email(str(email), shopee_user=username, source="vubel_addmail", meta=meta)
                if values:
                    values[5] = str(values[5]) + " | ✓ Đã lưu vào Đọc Mail"
                    self.mail_tree.item(iid, values=values, tags=("ok",))
            except Exception as exc:
                if values:
                    values[5] = str(values[5]) + f" | Lưu Đọc Mail lỗi: {exc}"
                    self.mail_tree.item(iid, values=values, tags=("ok",))
        self.mail_results.append({
            "stt": idx,
            "raw_account": account,
            "resolved_cookie": result.resolved_cookie,
            "username": username,
            "password": password,
            "email": email,
            "proxy": proxy_used,
            "spc_f": spcf,
            "spc_st": resolved_account["spc_st"] or source_account["spc_st"],
            "success": result.success,
            "status": "THÀNH CÔNG" if result.success else "THẤT BẠI",
            "message": result.message,
            "http_status": result.status_code,
            "response": result.response_data,
        })

    def _finish_mail_batch(self, stopped: bool, processed: int, total: int, ok_count: int):
        self.mail_running = False
        self.mail_run_button.configure(state="normal")
        self.mail_stop_button.configure(state="disabled")
        self.mail_accounts_text.configure(state="normal")
        self.mail_emails_text.configure(state="normal")
        self.addmail_rotate_seconds_entry.configure(state="normal")
        prefix = "Đã dừng" if stopped else "Hoàn tất"
        self.mail_status_var.set(f"{prefix}: {processed}/{total} | Thành công: {ok_count} | Thất bại: {processed-ok_count}")
        if (not stopped) and processed >= total and total > 0 and self.completion_alarm_var.get():
            self._play_completion_alarm()

    def stop_add_mail(self):
        if self.mail_running:
            self.mail_stop_event.set()
            self.mail_stop_button.configure(state="disabled")
            self.mail_status_var.set("Đang dừng...")

    def clear_mail_results(self):
        if self.mail_running:
            return
        self.mail_tree.delete(*self.mail_tree.get_children())
        self.mail_results.clear()
        if hasattr(self, "mail_proxy_status_var"):
            self.mail_proxy_status_var.set("Proxy Shopee AddMail: Direct")
        self.mail_status_var.set("Sẵn sàng.")

    def start_mailfree(self):
        if self.mailfree_running:
            return
        if self.mail_running:
            messagebox.showwarning("AddMail đang chạy", "Hãy đợi hoặc dừng tab AddMail trước khi chạy MailFree.")
            return
        try:
            accounts = parse_line_list(
                self.mailfree_accounts_text.get("1.0", "end"),
                MAX_BATCH_COOKIES,
                "tài khoản MailFree",
            )
            delay = self._read_delay()
        except ValueError as exc:
            messagebox.showwarning("Dữ liệu không hợp lệ", str(exc))
            return
        if not accounts:
            messagebox.showwarning("Thiếu dữ liệu", "Cần nhập ít nhất 1 tài khoản trong tab MailFree.")
            return

        api_key = _VUBEL_KEY_RAM
        proxy_key = self.proxy_var.get().strip()
        proxy_protocol = normalize_proxy_protocol(self.proxy_protocol_var.get())
        manual_proxy = ""
        manual_proxy_protocol_value = proxy_protocol
        auto_new_ip = bool(self.kiotproxy_auto_new_ip_var.get())
        if not api_key:
            messagebox.showwarning("Chưa xác thực", "Vubel key chưa sẵn sàng. Hãy khởi động lại và kích hoạt license.")
            return

        self._save_config()
        self.clear_mailfree_results()
        for idx, account in enumerate(accounts, 1):
            preview = f"{account[:18]}...{account[-12:]}" if len(account) > 34 else account
            self.mailfree_tree.insert(
                "",
                "end",
                iid=f"mailfree-{idx}",
                values=(idx, preview, "<Vubel đang tạo>", "Direct", "CHỜ", "Chưa xử lý"),
            )

        self.mailfree_running = True
        self.mailfree_stop_event.clear()
        self.mailfree_run_button.configure(state="disabled")
        self.mailfree_stop_button.configure(state="normal")
        self.mailfree_accounts_text.configure(state="disabled")
        self.mailfree_status_var.set(f"Đang xử lý 0/{len(accounts)}...")
        threading.Thread(
            target=self._mailfree_worker,
            args=(accounts, delay, api_key, proxy_key, proxy_protocol, auto_new_ip, manual_proxy, manual_proxy_protocol_value),
            daemon=True,
        ).start()

    def _mailfree_worker(
        self,
        accounts: list[str],
        delay: float,
        api_key: str,
        proxy_key: str,
        proxy_protocol: str = PROXY_PROTOCOL_HTTP,
        auto_new_ip: bool = False,
        manual_proxy: str = "",
        manual_proxy_protocol_value: str = PROXY_PROTOCOL_HTTP,
    ):
        using_manual_proxy = bool(manual_proxy)
        if using_manual_proxy:
            runtime_protocol = infer_manual_proxy_protocol(manual_proxy, manual_proxy_protocol_value)
            pm = ProxyManager(manual_proxy, proxy_protocol=runtime_protocol)
        else:
            pm = ProxyManager(proxy_key, proxy_protocol=proxy_protocol) if proxy_key else None
        if pm is not None and not using_manual_proxy:
            self._kiotproxy_manager = pm
        processed = 0
        ok_count = 0

        for idx, account in enumerate(accounts, 1):
            if self.mailfree_stop_event.is_set():
                break

            proxy_url = None
            proxy_label = "Direct"
            if pm:
                try:
                    if pm.needs_rotation():
                        if using_manual_proxy:
                            status = "đang dùng proxy thủ công..."
                        else:
                            status = "đang lấy IP mới..." if auto_new_ip and pm.current_proxy is None and not pm.static_proxy else "đang đổi IP..."
                        self.root.after(0, lambda s=status: self.mailfree_proxy_status_var.set(f"Proxy MailFree: {s}"))
                        proxy_url = acquire_proxy_until_ready(
                            pm, self.mailfree_stop_event,
                            lambda text: self.root.after(0, self.mailfree_proxy_status_var.set, text),
                            force_new=(not using_manual_proxy) and auto_new_ip and pm.current_proxy is None,
                        )
                        if proxy_url is None:
                            break
                    else:
                        proxy_url = pm.current_proxy
                    proxy_label = self._format_mail_proxy(proxy_url)
                    if not using_manual_proxy:
                        self.root.after(0, self._sync_kiotproxy_manager_status, pm)
                    source = "Proxy tay" if using_manual_proxy else "KiotProxy"
                    self.root.after(0, lambda p=proxy_label, s=source: self.mailfree_proxy_status_var.set(f"Proxy MailFree: {s} • {p}"))
                except KiotProxyKeyExpiredError as exc:
                    self.root.after(0, self._record_mailfree_result, idx, account, "", TestResult(False, str(exc), None), "KEY HẾT HẠN")
                    processed += 1
                    self.mailfree_stop_event.set()
                    break
                except Exception as exc:
                    result = TestResult(False, f"Lỗi Proxy MailFree: {exc}", None)
                    self.root.after(0, self._record_mailfree_result, idx, account, "", result, "LỖI PROXY")
                    processed += 1
                    if idx < len(accounts) and self.mailfree_stop_event.wait(delay):
                        break
                    continue
            else:
                self.root.after(0, lambda: self.mailfree_proxy_status_var.set("Proxy MailFree: Direct (IP máy)"))

            self.root.after(0, self._mark_mailfree_running, idx, proxy_label, account)

            # Request tới Vubel luôn direct; proxy Shopee chỉ được truyền trong body.
            try:
                with build_httpx_client(proxy_url=None) as mailfree_client:
                    result, actual_email = vubel_mailfree(
                        mailfree_client,
                        account,
                        api_key,
                        proxy_url,
                    )
            except Exception as exc:
                result = TestResult(False, f"Không tạo được HTTP client tới Vubel: {exc}", None)
                actual_email = ""

            if result.success:
                ok_count += 1
            processed += 1
            self.root.after(0, self._record_mailfree_result, idx, account, actual_email, result, proxy_label)
            self.root.after(
                0,
                lambda p=processed, t=len(accounts), o=ok_count: self.mailfree_status_var.set(
                    f"Tiến độ: {p}/{t} | Thành công: {o} | Thất bại: {p-o}"
                ),
            )
            if idx < len(accounts) and self.mailfree_stop_event.wait(delay):
                break

        self.root.after(
            0,
            self._finish_mailfree_batch,
            self.mailfree_stop_event.is_set(),
            processed,
            len(accounts),
            ok_count,
        )

    def _mark_mailfree_running(self, idx: int, proxy_label: str = "Direct", account: str = ""):
        iid = f"mailfree-{idx}"
        values = list(self.mailfree_tree.item(iid, "values"))
        if values:
            values[3] = proxy_label
            values[4] = "ĐANG XỬ LÝ"
            values[5] = "MAILFREE: chỉ tạo email, không kiểm tra cookie Shopee"
            self.mailfree_tree.item(iid, values=values, tags=("running",))
            self.mailfree_tree.see(iid)

    def _record_mailfree_result(self, idx: int, account: str, email: str, result: TestResult, proxy_used: str = "Direct"):
        iid = f"mailfree-{idx}"
        values = list(self.mailfree_tree.item(iid, "values"))
        if values:
            values[2] = email or "<không nhận được email>"
            values[3] = proxy_used
            values[4] = "ĐÃ TẠO" if result.success else "THẤT BẠI"
            values[5] = result.message
            self.mailfree_tree.item(iid, values=values, tags=(("ok",) if result.success else ("fail",)))

        username, password, spcf = parse_combo(account) if "|" in account else ("", "", cookie_value(account, "SPC_F"))
        mail_password = extract_mailfree_password_from_data(result.response_data, email)
        if result.success and email and "@" in str(email):
            if mail_password:
                self._remember_mailfree_mailbox(str(email), mail_password)
                if values:
                    values[5] = str(values[5]) + " | ✓ Có thể đọc inbox Vubel"
                    self.mailfree_tree.item(iid, values=values, tags=("ok",))
            elif values:
                values[5] = str(values[5]) + " | Sẽ tìm mật khẩu từ lịch sử Vubel khi đọc"
                self.mailfree_tree.item(iid, values=values, tags=("ok",))

        self.mailfree_results.append({
            "stt": idx,
            "raw_account": account,
            "username": username,
            "password": password,
            "email": email,
            "mail_password": mail_password,
            "proxy": proxy_used,
            "spc_f": spcf,
            "spc_st": cookie_value(account, "SPC_ST"),
            "success": result.success,
            "status": "ĐÃ TẠO" if result.success else "THẤT BẠI",
            "message": result.message,
            "http_status": result.status_code,
            "response": result.response_data,
        })

    def _remember_mailfree_mailbox(
        self,
        email: str,
        password: str,
        *,
        created_at: str = "",
    ) -> None:
        email = _normalize_email_for_addmail(email)
        password = str(password or "").strip()
        if not EMAIL_ADDRESS_RE.fullmatch(email) or not password:
            return
        marker = email.casefold()
        existing = next(
            (
                item for item in self.mailfree_mailboxes
                if str(item.get("email") or "").casefold() == marker
            ),
            None,
        )
        timestamp = str(created_at or datetime.now().isoformat(timespec="seconds"))
        if existing is not None:
            existing.update({"email": email, "password": password, "created_at": timestamp})
            return
        self.mailfree_mailboxes.append({
            "email": email,
            "password": password,
            "created_at": timestamp,
        })
        if len(self.mailfree_mailboxes) > MAX_BATCH_COOKIES:
            del self.mailfree_mailboxes[:-MAX_BATCH_COOKIES]

    def _mailfree_mailbox_password(self, email: str) -> str:
        marker = _normalize_email_for_addmail(email).casefold()
        for item in reversed(self.mailfree_mailboxes):
            if str(item.get("email") or "").casefold() == marker:
                return str(item.get("password") or "").strip()
        return ""

    def transfer_mailfree_to_addmail(self):
        """Chuyển cặp tài khoản–mail đã tạo; không tự chạy AddMail."""
        if self.mailfree_running:
            messagebox.showwarning("MailFree đang chạy", "Hãy đợi MailFree hoàn tất trước khi chuyển.")
            return
        if self.mail_running:
            messagebox.showwarning("AddMail đang chạy", "Hãy đợi AddMail hoàn tất trước khi thay đổi dữ liệu.")
            return
        pairs = [
            (str(item.get("raw_account") or ""), str(item.get("email") or ""))
            for item in sorted(self.mailfree_results, key=lambda row: int(row.get("stt") or 0))
            if item.get("success") and item.get("email")
        ]
        if not pairs:
            messagebox.showwarning("Chưa có MailFree", "Chưa có email MailFree tạo thành công để chuyển.")
            return
        try:
            accounts_text, emails_text, added = merge_mailfree_pairs_into_addmail(
                self.mail_accounts_text.get("1.0", "end-1c"),
                self.mail_emails_text.get("1.0", "end-1c"),
                pairs,
            )
        except ValueError as exc:
            messagebox.showwarning("Chưa thể chuyển", str(exc))
            return

        if added:
            self.mail_accounts_text.delete("1.0", "end")
            self.mail_emails_text.delete("1.0", "end")
            self.mail_accounts_text.insert("1.0", accounts_text)
            self.mail_emails_text.insert("1.0", emails_text)
            self._save_config()
        self.mail_flow_notebook.select(self.addmail_tab)
        if added:
            self.mail_status_var.set(
                f"Đã nhận {added} cặp từ MailFree. Kiểm tra dữ liệu rồi bấm Bắt đầu thêm mail."
            )
            messagebox.showinfo(
                "Đã chuyển sang AddMail",
                f"Đã nối {added} cặp tài khoản–email theo đúng thứ tự.\n"
                "AddMail chưa tự chạy; hãy kiểm tra rồi bấm Bắt đầu thêm mail.",
            )
        else:
            self.mail_status_var.set("Các cặp MailFree đã có sẵn trong AddMail; không thêm trùng.")

    def _selected_mailfree_result(self) -> dict[str, Any] | None:
        selected = self.mailfree_tree.selection()
        if selected:
            values = self.mailfree_tree.item(selected[0], "values")
            try:
                stt = int(values[0])
            except (TypeError, ValueError, IndexError):
                stt = -1
            for item in reversed(self.mailfree_results):
                if int(item.get("stt") or -2) == stt:
                    return item
        for item in reversed(self.mailfree_results):
            if item.get("success") and item.get("email"):
                return item
        if self.mailfree_mailboxes:
            mailbox = self.mailfree_mailboxes[-1]
            return {
                "email": mailbox.get("email", ""),
                "mail_password": mailbox.get("password", ""),
                "response": None,
                "success": True,
            }
        return None

    def read_selected_mailfree(self):
        item = self._selected_mailfree_result()
        if not item:
            messagebox.showwarning(
                "Chưa chọn MailFree",
                "Hãy tạo MailFree hoặc chọn một dòng đã tạo trước khi đọc inbox.",
            )
            return
        email = _normalize_email_for_addmail(str(item.get("email") or ""))
        if not EMAIL_ADDRESS_RE.fullmatch(email):
            messagebox.showwarning("Email không hợp lệ", "Dòng đã chọn chưa có địa chỉ MailFree hợp lệ.")
            return
        api_key = _VUBEL_KEY_RAM
        if not api_key:
            messagebox.showwarning("Chưa xác thực", "Vubel key chưa sẵn sàng. Hãy khởi động lại và kích hoạt license.")
            return
        self._open_mailfree_inbox_window(email, item, api_key)

    def _open_mailfree_inbox_window(
        self,
        email: str,
        source_item: dict[str, Any],
        api_key: str,
    ) -> None:
        window = tk.Toplevel(self.root)
        window.title(f"Inbox MailFree • {email}")
        window.geometry("1080x680")
        window.minsize(780, 520)
        window.configure(bg="#f6f7f9")
        try:
            apply_ryan_window_icon(window)
        except Exception:
            pass

        outer = tk.Frame(window, bg="#f6f7f9")
        outer.pack(fill="both", expand=True, padx=16, pady=16)
        outer.grid_columnconfigure(0, weight=1)
        outer.grid_rowconfigure(1, weight=2)
        outer.grid_rowconfigure(3, weight=3)

        head = tk.Frame(outer, bg="#f6f7f9")
        head.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        tk.Label(
            head,
            text=email,
            bg="#f6f7f9",
            fg="#222222",
            font=("Segoe UI", 13, "bold"),
        ).pack(side="left")
        status_var = tk.StringVar(value="Đang lấy inbox từ Vubel...")
        tk.Label(
            head,
            textvariable=status_var,
            bg="#f6f7f9",
            fg="#3976c1",
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left", padx=16)

        mail_tree = ttk.Treeview(
            outer,
            columns=("stt", "time", "sender", "subject"),
            show="headings",
            height=9,
        )
        for column, title in (
            ("stt", "STT"), ("time", "Thời gian"),
            ("sender", "Người gửi"), ("subject", "Tiêu đề"),
        ):
            mail_tree.heading(column, text=title)
        mail_tree.column("stt", width=55, anchor="center", stretch=False)
        mail_tree.column("time", width=180, stretch=False)
        mail_tree.column("sender", width=240)
        mail_tree.column("subject", width=520)
        mail_tree.grid(row=1, column=0, sticky="nsew")

        detail = scrolledtext.ScrolledText(
            outer,
            font=("Segoe UI", 10),
            wrap="word",
            background="#ffffff",
            bd=0,
            relief="flat",
        )
        verify_bar = tk.Frame(outer, bg="#ffffff", bd=0, relief="flat")
        verify_bar.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        verify_bar.grid_columnconfigure(1, weight=1)
        tk.Label(
            verify_bar,
            text="🔐 Link xác thực Shopee",
            bg="#ffffff",
            fg="#333333",
            font=("Segoe UI", 9, "bold"),
        ).grid(row=0, column=0, padx=(10, 6), pady=8, sticky="w")
        verify_link_var = tk.StringVar(value="")
        verify_entry = ttk.Entry(verify_bar, textvariable=verify_link_var)
        verify_entry.grid(row=0, column=1, sticky="ew", padx=4, pady=8)

        detail.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
        detail.insert("1.0", "Đang tải thư...")
        detail.configure(state="disabled")
        loaded_messages: list[dict[str, Any]] = []
        loading = {"active": False}

        def set_detail(text: str):
            detail.configure(state="normal")
            detail.delete("1.0", "end")
            detail.insert("1.0", text)
            detail.configure(state="disabled")

        def message_verify_link(message: dict[str, Any]) -> str:
            if not callable(detect_shopee_verification_link):
                return ""
            _received, _sender, subject = mailfree_message_summary(message)
            raw_parts: list[str] = []
            html_parts: list[str] = []
            for key in ("text", "body", "content", "message", "bodyText"):
                candidate = message.get(key)
                if isinstance(candidate, dict):
                    candidate = candidate.get("text") or candidate.get("content") or candidate.get("html") or ""
                if candidate not in (None, "", [], {}):
                    raw_parts.append(str(candidate))
            for key in ("html", "bodyHtml"):
                candidate = message.get(key)
                if candidate not in (None, "", [], {}):
                    html_parts.append(str(candidate))
            # Fallback JSON giúp bắt link khi Vubel đổi tên field.
            raw_parts.append(json.dumps(message, ensure_ascii=False))
            return detect_shopee_verification_link(
                subject,
                "\n".join(raw_parts),
                "\n".join(html_parts),
            )

        def show_selected(_event=None):
            selection = mail_tree.selection()
            if not selection:
                return
            try:
                index = int(str(selection[0]).rsplit("-", 1)[-1])
                message = loaded_messages[index]
                set_detail(format_mailfree_message(message))
                verify_link_var.set(message_verify_link(message))
            except (ValueError, IndexError):
                return

        def filter_verify_link():
            if not loaded_messages:
                status_var.set("Inbox chưa có thư để lọc.")
                verify_link_var.set("")
                return
            for index, message in enumerate(loaded_messages):
                link = message_verify_link(message)
                if not link:
                    continue
                iid = f"mailfree-inbox-{index}"
                if mail_tree.exists(iid):
                    mail_tree.selection_set(iid)
                    mail_tree.focus(iid)
                    mail_tree.see(iid)
                verify_link_var.set(link)
                set_detail(format_mailfree_message(message))
                status_var.set("Đã lọc được link xác thực Shopee. Bấm Copy link.")
                return
            verify_link_var.set("")
            status_var.set("Chưa tìm thấy link xác thực Shopee trong các thư hiện có.")

        def copy_verify_link():
            link = verify_link_var.get().strip()
            if not link:
                filter_verify_link()
                link = verify_link_var.get().strip()
            if not link:
                messagebox.showwarning("Chưa có link", "Chưa tìm thấy link xác thực Shopee để copy.", parent=window)
                return
            window.clipboard_clear()
            window.clipboard_append(link)
            window.update_idletasks()
            status_var.set("Đã copy link xác thực Shopee vào clipboard.")

        ttk.Button(
            verify_bar,
            text="🔎  Lọc link",
            command=filter_verify_link,
            style="Soft.TButton",
        ).grid(row=0, column=2, padx=4, pady=8)
        ttk.Button(
            verify_bar,
            text="📋  Copy",
            command=copy_verify_link,
            style="Accent.TButton",
        ).grid(row=0, column=3, padx=(4, 10), pady=8)

        mail_tree.bind("<<TreeviewSelect>>", show_selected)

        def deliver(result: TestResult, messages: list[dict[str, Any]], password: str):
            if not window.winfo_exists():
                return
            loading["active"] = False
            reload_button.configure(state="normal")
            mail_tree.delete(*mail_tree.get_children())
            loaded_messages.clear()
            if not result.success:
                status_var.set("Không đọc được inbox")
                set_detail(result.message)
                return
            if password:
                source_item["mail_password"] = password
                self._remember_mailfree_mailbox(email, password)
                self._save_config()
            loaded_messages.extend(messages)
            status_var.set(result.message)
            if not loaded_messages:
                set_detail("Inbox hiện chưa có thư. Bấm Tải lại để kiểm tra lần nữa.")
                return
            for index, message in enumerate(loaded_messages):
                received, sender, subject = mailfree_message_summary(message)
                mail_tree.insert(
                    "",
                    "end",
                    iid=f"mailfree-inbox-{index}",
                    values=(index + 1, received, sender, subject),
                )
            first = mail_tree.get_children()[0]
            mail_tree.selection_set(first)
            mail_tree.focus(first)
            show_selected()

        def worker():
            password = str(source_item.get("mail_password") or "").strip()
            if not password:
                password = extract_mailfree_password_from_data(source_item.get("response"), email)
            if not password:
                password = self._mailfree_mailbox_password(email)
            try:
                with build_httpx_client(proxy_url=None) as inbox_client:
                    if not password:
                        password = vubel_find_mailfree_password(inbox_client, api_key, email)
                    if password:
                        result, messages = vubel_read_mailfree_inbox(
                            inbox_client,
                            api_key,
                            email,
                            password,
                        )
                    else:
                        result = TestResult(
                            False,
                            "Vubel chưa trả mật khẩu inbox cho email này và cũng không tìm thấy trong lịch sử. "
                            "Mail đã được tạo nhưng hiện chưa đủ thông tin để đọc.",
                            None,
                        )
                        messages = []
            except Exception as exc:
                result = TestResult(False, f"Không tạo được kết nối đọc MailFree: {exc}", None)
                messages = []
            self.root.after(0, deliver, result, messages, password)

        def reload_inbox():
            if loading["active"]:
                return
            loading["active"] = True
            reload_button.configure(state="disabled")
            status_var.set("Đang lấy inbox từ Vubel...")
            set_detail("Đang tải thư...")
            threading.Thread(target=worker, daemon=True).start()

        reload_button = ttk.Button(
            head,
            text="↻  Tải lại",
            command=reload_inbox,
            style="Soft.TButton",
        )
        reload_button.pack(side="right")
        reload_inbox()

    def _finish_mailfree_batch(self, stopped: bool, processed: int, total: int, ok_count: int):
        self.mailfree_running = False
        self.mailfree_run_button.configure(state="normal")
        self.mailfree_stop_button.configure(state="disabled")
        self.mailfree_accounts_text.configure(state="normal")
        prefix = "Đã dừng" if stopped else "Hoàn tất"
        self.mailfree_status_var.set(
            f"{prefix}: {processed}/{total} | Đã tạo: {ok_count} | Thất bại: {processed-ok_count} | "
            "Bấm Chuyển sang AddMail khi sẵn sàng."
        )
        if ok_count:
            self._save_config()
        if (not stopped) and processed >= total and total > 0 and self.completion_alarm_var.get():
            self._play_completion_alarm()

    def stop_mailfree(self):
        if self.mailfree_running:
            self.mailfree_stop_event.set()
            self.mailfree_stop_button.configure(state="disabled")
            self.mailfree_status_var.set("Đang dừng...")

    def clear_mailfree_results(self):
        if self.mailfree_running:
            return
        self.mailfree_tree.delete(*self.mailfree_tree.get_children())
        self.mailfree_results.clear()
        if hasattr(self, "mailfree_proxy_status_var"):
            self.mailfree_proxy_status_var.set("Proxy MailFree: Direct")
        self.mailfree_status_var.set("Sẵn sàng.")

    def export_mail_excel(self):
        self._export_mail_result_list(
            self.mail_results,
            dialog_title="Lưu kết quả AddMail",
            empty_message="Chưa có kết quả AddMail để xuất.",
            initialfile="Shopee_AddMail.xlsx",
            sheet_title="AddMail",
            proxy_header="Proxy Shopee AddMail",
        )

    def export_mailfree_excel(self):
        self._export_mail_result_list(
            self.mailfree_results,
            dialog_title="Lưu kết quả MailFree",
            empty_message="Chưa có kết quả MailFree để xuất.",
            initialfile="Shopee_MailFree.xlsx",
            sheet_title="MailFree",
            proxy_header="Proxy MailFree",
        )

    def _export_mail_result_list(
        self,
        results: list[dict[str, Any]],
        *,
        dialog_title: str,
        empty_message: str,
        initialfile: str,
        sheet_title: str,
        proxy_header: str,
    ):
        if not results:
            messagebox.showinfo("Trống", empty_message)
            return
        if Workbook is None:
            messagebox.showerror("Thiếu thư viện", "Cần cài openpyxl:\npython -m pip install openpyxl")
            return
        path = filedialog.asksaveasfilename(
            title=dialog_title,
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile=initialfile,
        )
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_title
        headers = [
            "STT", "Tài khoản", "Email", proxy_header, "SPC_F", "SPC_ST",
            "SPC_F / USER / MẬT KHẨU (3 dòng)", "FULL COOKIE", "Trạng thái", "HTTP", "Thông báo"
        ]
        ws.append(headers)
        if Font is not None:
            for cell in ws[1]:
                cell.font = Font(bold=True)
        for item in sorted(results, key=lambda x: x["stt"]):
            paste_3_lines = ""
            spcst = (
                _extract_named_token(item.get("resolved_cookie", ""), "SPC_ST", "SPCST")
                or str(item.get("spc_st") or "")
                or _extract_named_token(item.get("raw_account", ""), "SPC_ST", "SPCST")
            )
            spcst = _extract_named_token(spcst, "SPC_ST", "SPCST") or spcst
            if item.get("spc_f") and item.get("username") and item.get("password"):
                paste_3_lines = f"SPC_F={item['spc_f']}\n{item['username']}\n{item['password']}"
            ws.append([
                item["stt"], item["username"], item["email"], item.get("proxy", "Direct"), item["spc_f"], f"SPC_ST={spcst}" if spcst else "",
                paste_3_lines, item.get("raw_account", ""), item["status"], item["http_status"], item["message"]
            ])
        if Alignment is not None:
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row=row_idx, column=7).alignment = Alignment(wrap_text=True, vertical="top")
                ws.row_dimensions[row_idx].height = 45
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for idx, width in enumerate([7, 24, 30, 22, 40, 55, 42, 60, 16, 10, 60], 1):
            ws.column_dimensions[get_column_letter(idx)].width = width
        try:
            wb.save(path)
            try:
                open_file_in_default_app(path)
                open_note = "\n\nĐã tự mở file Excel."
            except Exception as open_exc:
                open_note = f"\n\n{open_exc}"
            messagebox.showinfo("Thành công", f"Đã xuất Excel:\n{path}{open_note}")
        except Exception as exc:
            messagebox.showerror("Lỗi", f"Không thể lưu Excel: {exc}")

    def _read_delay(self) -> float:
        try:
            delay = float(self.delay_var.get().strip().replace(",", "."))
        except ValueError:
            raise ValueError("Độ trễ phải là một số")
        if not math.isfinite(delay) or delay < MIN_DELAY_SECONDS:
            raise ValueError(f"Độ trễ phải là số hữu hạn từ {MIN_DELAY_SECONDS:g} giây, không giới hạn tối đa")
        return delay

    def _read_voucher_delay(self) -> float:
        return parse_voucher_request_delay(self.voucher_delay_var.get())

    def start_batch(self):
        if self.running:
            return

        try:
            refresh_only = self.refresh_only_var.get()
            codes = parse_text_list(self.codes_text.get("1.0", "end"), MAX_BATCH_CODES, "mã voucher")
            codes, duplicate_code_count = dedupe_voucher_entries(codes)
            if not refresh_only:
                codes = [c for c in codes if self.voucher_selected.get(c, True)]
            cookies = parse_line_list(self.cookies_text.get("1.0", "end"), MAX_BATCH_COOKIES, "cookie/tài khoản")
            delay = self._read_voucher_delay()
            thread_configs = normalize_voucher_thread_configs(self.voucher_thread_configs)
        except ValueError as exc:
            messagebox.showwarning("Dữ liệu không hợp lệ", str(exc))
            return

        if not cookies:
            messagebox.showwarning("Thiếu dữ liệu", "Cần nhập ít nhất 1 cookie/tài khoản.")
            return
        if refresh_only:
            invalid = [x for x in cookies if "|" not in x]
            if invalid:
                messagebox.showwarning("Sai định dạng", "Chế độ làm mới SPC_ST chỉ nhận mỗi dòng dạng User|Pass|SPC_F.")
                return
        else:
            if not codes:
                messagebox.showwarning(
                    "Thiếu dữ liệu",
                    "Cần nhập ít nhất 1 mã voucher và tích chọn (ô Chọn) trong bảng xem trước.",
                )
                return

        # Xóa kết quả CŨ trước khi tạo danh sách task mới.
        # clear_results() có self.tasks.clear(), nên nếu gọi sau khi gán self.tasks
        # thì nút Bắt đầu sẽ có 0 task và trông như không chạy.
        self._save_config()
        self.clear_results()
        self.batch_refresh_only = refresh_only
        self.batch_skip_locked = bool(self.skip_locked_var.get())
        if refresh_only:
            self.tasks = [(cookie, "") for cookie in cookies]
        else:
            self.tasks = [(cookie, code) for cookie in cookies for code in codes]
        self.total_count = len(self.tasks)

        for index, (cookie, code) in enumerate(self.tasks, 1):
            preview = f"{cookie[:15]}...{cookie[-15:]}" if len(cookie) > 30 else cookie
            code_display = code if code else "<Làm mới SPC_ST>"
            self.results_tree.insert("", "end", iid=str(index), values=(index, preview, "", code_display, "CHỜ", "Chưa xử lý"), tags=("pending",))

        self.running = True
        self.stop_event.clear()
        self.run_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.clear_button.configure(state="disabled")
        self.export_button.configure(state="disabled")
        self.codes_text.configure(state="disabled")
        self.cookies_text.configure(state="disabled")
        self._set_voucher_config_controls_state(False)
        if (not refresh_only) and duplicate_code_count:
            self.status_var.set(f"Đang xử lý 0/{self.total_count}... (Đã bỏ {duplicate_code_count} voucher trùng)")
        else:
            self.status_var.set(f"Đang xử lý 0/{self.total_count}...")

        self.proxy_managers = [
            ProxyManager(item.proxy_key, item.rotate_seconds) if item.proxy_key else None
            for item in thread_configs
        ]
        self.proxy_worker_status = {
            index: "Direct" if manager is None else "Chờ lấy IP"
            for index, manager in enumerate(self.proxy_managers, 1)
        }
        vubel_key = _VUBEL_KEY_RAM
        self._update_proxy_countdown()
        threading.Thread(
            target=self._batch_worker,
            args=(delay, vubel_key),
            name="VoucherCoordinator",
            daemon=True,
        ).start()

    def _update_proxy_countdown(self):
        if not self.running:
            return
        parts: list[tuple[str, bool]] = []
        for worker_number, manager in enumerate(self.proxy_managers, 1):
            message = self.proxy_worker_status.get(worker_number, "")
            if message.startswith(("Đang", "Lỗi")):
                detail = message
            elif manager is None:
                detail = "Direct"
            elif manager.current_proxy:
                proxy_name = proxy_display_name(manager.current_proxy)
                detail = f"{proxy_name} (tĩnh)" if manager.static_proxy else f"{proxy_name} ({manager.get_remaining_time()}s)"
            else:
                detail = message or "Chờ lấy IP"
            parts.append((f"L{worker_number}: {detail}", message.startswith("Lỗi")))
        errors = [text for text, is_error in parts if is_error]
        if errors:
            # Luôn đưa lỗi lên trước, kể cả lỗi nằm ở luồng thứ 4 trở đi.
            visible = errors[:2]
            visible_count = len(visible)
        else:
            visible = [text for text, _ in parts[:3]]
            visible_count = len(visible)
        if len(parts) > 3:
            visible.append(f"{len(parts)} luồng tổng cộng")
        elif errors and len(parts) > visible_count:
            visible.append(f"{len(parts)} luồng tổng cộng")
        self.proxy_status_var.set("  |  ".join(visible))
        self.root.after(1000, self._update_proxy_countdown)

    def _close_active_ivs_dialog(self) -> None:
        dialog = self.active_ivs_dialog
        self.active_ivs_dialog = None
        if dialog is None:
            return
        try:
            if dialog.winfo_exists():
                dialog.on_close()
        except Exception:
            pass

    def _clear_active_ivs_dialog(self, dialog: IVSDialog) -> None:
        if self.active_ivs_dialog is dialog:
            self.active_ivs_dialog = None

    def _batch_worker(self, delay: float, vubel_key: str):
        def set_proxy_status(worker_number: int, msg: str):
            def _update_ui():
                self.proxy_worker_status[worker_number] = msg
            self.root.after(0, _update_ui)

        def handle_ivs(session_id, methods, proxy_url):
            evt = threading.Event()
            container = {}
            def _show():
                if self.stop_event.is_set():
                    evt.set()
                    return
                dialog = IVSDialog(
                    self.root,
                    session_id,
                    methods,
                    vubel_key,
                    proxy_url,
                    container,
                    evt,
                )
                self.active_ivs_dialog = dialog
                container["_dialog"] = dialog
            self.root.after(0, _show)
            while not evt.wait(0.1):
                if self.stop_event.is_set():
                    self.root.after(0, self._close_active_ivs_dialog)
                    return None
            dialog = container.get("_dialog")
            if dialog is not None:
                self.root.after(0, self._clear_active_ivs_dialog, dialog)
            return container.get("spc_st")

        error_message = ""
        try:
            run_voucher_workers(
                self.tasks,
                delay,
                self.proxy_managers,
                vubel_key,
                self.stop_event,
                self.cookie_info_cache,
                lambda idx, c, cd: self.root.after(0, self._mark_running, str(idx)),
                lambda idx, c, cd, resolved, res: self.root.after(0, self._record_result, str(idx), c, resolved, res),
                set_proxy_status,
                handle_ivs,
                refresh_only=self.batch_refresh_only,
                skip_locked=getattr(self, "batch_skip_locked", True),
            )
        except Exception as exc:
            error_message = str(exc)
        self.root.after(
            0,
            self._finish_batch,
            self.stop_event.is_set(),
            error_message,
        )

    def _mark_running(self, iid: str):
        values = list(self.results_tree.item(iid, "values"))
        values[4] = "ĐANG XỬ LÝ"
        values[5] = "Đang kiểm tra/gửi request..."
        self.results_tree.item(iid, values=values, tags=("running",))
        self.results_tree.see(iid)

    def _record_result(self, iid: str, raw_cookie: str, resolved_cookie: str, result: TestResult):
        values = list(self.results_tree.item(iid, "values"))
        info = self.cookie_info_cache.get(raw_cookie, {})
        if info.get("status") == "success":
            values[2] = info.get("username", "")
        elif "|" in raw_cookie:
            values[2] = parse_combo(raw_cookie)[0]
        else:
            values[2] = "Lỗi/Die"

        if result.success:
            if self.batch_refresh_only:
                values[4] = "ĐÃ LÀM MỚI"
            else:
                already_saved = "đã có trong kho" in result.message.lower()
                values[4] = "ĐÃ ĐƯỢC LƯU RỒI" if already_saved else "LƯU THÀNH CÔNG"
                if already_saved:
                    self.already_saved_count += 1
                else:
                    self.new_saved_count += 1
            tag, self.ok_count = "ok", self.ok_count + 1
        elif getattr(result, "skipped", False):
            values[4] = "BỎ QUA"
            tag, self.skipped_count = "skipped", self.skipped_count + 1
        elif "bị khoá" in result.message.casefold() or "bị khóa" in result.message.casefold():
            values[4] = "BỊ KHOÁ"
            tag, self.fail_count = "fail", self.fail_count + 1
        else:
            invalid_code = None if self.batch_refresh_only else extract_invalid_message_code(result.response_data)
            specific_label = INVALID_MESSAGE_STATUS_LABELS.get(invalid_code) if invalid_code is not None else None
            if specific_label:
                values[4] = specific_label
            else:
                values[4] = "LỖI LÀM MỚI" if self.batch_refresh_only else "KHÔNG LƯU"
            tag, self.fail_count = "fail", self.fail_count + 1

        values[5] = result.message
        self.results_tree.item(iid, values=values, tags=(tag,))
        self.result_by_iid[iid] = IIDData(raw_cookie, resolved_cookie, result)
        self.processed_count += 1
        skipped_part = f" | Bỏ qua (khoá): {self.skipped_count}" if self.skipped_count else ""
        self.status_var.set(
            f"Tiến độ: {self.processed_count}/{self.total_count} | Thành công: {self.ok_count} "
            f"(Đã có: {self.already_saved_count} | Mới lưu: {self.new_saved_count}) | Thất bại: {self.fail_count}"
            f"{skipped_part}"
        )

    def stop_batch(self):
        if not self.running: return
        self.stop_event.set()
        self._close_active_ivs_dialog()
        self.stop_button.configure(state="disabled")
        self.status_var.set("Đang dừng...")

    def _finish_batch(self, stopped: bool, error_message: str = ""):
        if stopped:
            for iid in self.results_tree.get_children():
                if self.results_tree.item(iid, "values")[4] == "CHỜ":
                    v = list(self.results_tree.item(iid, "values"))
                    v[4], v[5] = "ĐÃ DỪNG", "Chưa gửi request"
                    self.results_tree.item(iid, values=v, tags=("pending",))

        self.running = False
        self.run_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.clear_button.configure(state="normal")
        self.export_button.configure(state="normal")
        self.codes_text.configure(state="normal")
        self.cookies_text.configure(state="normal")
        self._set_voucher_config_controls_state(True)
        self.proxy_status_var.set("")
        self.proxy_worker_status.clear()
        if error_message:
            self.status_var.set(f"Lỗi đa luồng: {error_message}")
            messagebox.showerror("Lỗi xử lý Voucher", error_message)
        else:
            prefix = "Đã dừng" if stopped else "Hoàn tất"
            skipped_part = f" | Bỏ qua (khoá): {self.skipped_count}" if self.skipped_count else ""
            self.status_var.set(
                f"{prefix}: {self.processed_count}/{self.total_count} | Thành công: {self.ok_count} "
                f"(Đã có: {self.already_saved_count} | Mới lưu: {self.new_saved_count}) | Thất bại: {self.fail_count}"
                f"{skipped_part}"
            )
        if (not stopped) and (not error_message) and self.processed_count >= self.total_count and self.total_count > 0 and self.completion_alarm_var.get():
            self._play_completion_alarm()

    def clear_results(self):
        if self.running: return
        self.results_tree.delete(*self.results_tree.get_children())
        self.result_by_iid.clear()
        self.cookie_info_cache.clear()
        self.tasks.clear()
        self.ok_count = self.fail_count = self.processed_count = self.total_count = 0
        self.already_saved_count = self.new_saved_count = self.skipped_count = 0
        self._set_details("")
        self.status_var.set("Sẵn sàng.")
        self.proxy_status_var.set("")

    def _update_copy_spcst_button(self):
        if self.refresh_only_var.get():
            self.copy_spcst_button.grid()
        else:
            self.copy_spcst_button.grid_remove()

    def copy_refreshed_spcst(self):
        """Copy các SPC_ST làm mới thành công, theo thứ tự bảng, mỗi dòng một cookie."""
        if not self.refresh_only_var.get() or not self.batch_refresh_only:
            messagebox.showinfo("Làm mới SPC_ST", "Hãy chạy chế độ Chỉ làm mới SPC_ST trước khi copy.")
            return
        cookies = []
        seen = set()
        for iid in self.results_tree.get_children():
            data = self.result_by_iid.get(iid)
            if data is None or not data.test_result.success:
                continue
            value = cookie_value(data.resolved_cookie or "", "SPC_ST")
            if value and value not in seen:
                seen.add(value)
                cookies.append(f"SPC_ST={value}")
        if not cookies:
            messagebox.showinfo("Chưa có SPC_ST", "Chưa có SPC_ST làm mới thành công để copy.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(cookies))
        self.status_var.set(f"Đã copy {len(cookies)} SPC_ST, mỗi dòng có tiền tố SPC_ST=.")

    def open_export_dialog(self):
        if not self.result_by_iid:
            messagebox.showinfo("Trống", "Chưa có kết quả nào để xuất!")
            return
        ExportDialog(self.root, self.result_by_iid, self.results_tree, self.cookie_info_cache)

    def _show_selected_result(self, _event):
        selected = self.results_tree.selection()
        if not selected: return
        iid = selected[0]
        data = self.result_by_iid.get(iid)
        if data is None:
            self._set_details("Chưa chạy...")
            return
            
        res = data.test_result
        text_resp = json.dumps({
            "success": res.success,
            "message": res.message,
            "http_status": res.status_code,
            "response": res.response_data,
        }, ensure_ascii=False, indent=2)
        
        info = self.cookie_info_cache.get(data.raw_cookie, {})
        text_info = json.dumps(info, ensure_ascii=False, indent=2)
        
        full_detail = (
            f"--- THÔNG TIN TÀI KHOẢN ---\n{text_info}\n\n"
            f"--- COOKIE/SPC_ST ĐÃ XỬ LÝ ---\n{data.resolved_cookie}\n\n"
            f"--- PHẢN HỒI API ---\n{text_resp}"
        )
        self._set_details(full_detail[:10000])

    def _set_details(self, text: str):
        self.details.configure(state="normal")
        self.details.delete("1.0", "end")
        if text: self.details.insert("1.0", text)
        self.details.configure(state="disabled")

def main():
    if httpx is None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Thiếu thư viện", 'Chạy lệnh: python -m pip install "httpx[http2]"')
        root.destroy()
        return

    if not run_license_gate():
        return

    root = tk.Tk()
    VoucherTesterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
