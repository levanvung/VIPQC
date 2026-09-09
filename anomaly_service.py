"""
VIPQC AI — Anomaly Report Cloud Service Module (Supabase Integration)
Provides full CRUD operations, image compression & cloud storage upload,
password authentication, and professional Excel reporting.
"""

import os
import sys
import io
import json
import uuid
import hashlib
from datetime import datetime
import urllib.request
import urllib.parse
import urllib.error
import ssl
from PIL import Image

def get_ssl_context():
    """Returns an SSL context that bypasses corporate proxy or missing root CA errors."""
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        return ctx
    except Exception:
        pass
    try:
        return ssl._create_unverified_context()
    except Exception:
        return None

CONFIG_FILE = "cloud_config.json"

def _get_config_path() -> str:
    """Returns absolute path to cloud_config.json."""
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.dirname(os.path.abspath(__file__))
    
    # Priority: current directory or workspace directory
    p_curr = os.path.join(os.getcwd(), CONFIG_FILE)
    if os.path.exists(p_curr):
        return p_curr
    p_base = os.path.join(base, CONFIG_FILE)
    if os.path.exists(p_base):
        return p_base
    return p_curr


def load_config() -> dict:
    """Loads cloud configuration from cloud_config.json."""
    cfg_path = _get_config_path()
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"[AnomalyService] Error reading config: {e}")
    # Default fallback
    return {
        "supabase_url": "https://tyeglupfonwvzcbpfcdh.supabase.co",
        "anon_key": "",
        "service_key": "",
        "bucket_name": "anomaly-images",
        "admin_password_hash": "69623787a01183df78e22b529263bd30f7e1f6e7178fe59603e856b83a37f5e1"
    }


def save_config(cfg: dict) -> bool:
    """Saves updated cloud configuration."""
    cfg_path = _get_config_path()
    try:
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"[AnomalyService] Error saving config: {e}")
        return False


# ── Password Authentication ──────────────────────────────────────────────────
def verify_admin_password(input_password: str) -> bool:
    """Verifies if input_password matches admin_password_hash."""
    if not input_password:
        return False
    cfg = load_config()
    target_hash = cfg.get("admin_password_hash", "")
    h = hashlib.sha256(input_password.encode("utf-8")).hexdigest()
    return h == target_hash


def change_admin_password(old_pass: str, new_pass: str) -> tuple[bool, str]:
    """Updates the admin password."""
    if not verify_admin_password(old_pass):
        return False, "Mật khẩu cũ không chính xác!"
    if not new_pass or len(new_pass.strip()) < 4:
        return False, "Mật khẩu mới phải có ít nhất 4 ký tự!"
    
    cfg = load_config()
    cfg["admin_password_hash"] = hashlib.sha256(new_pass.strip().encode("utf-8")).hexdigest()
    if save_config(cfg):
        return True, "Đổi mật khẩu thành công!"
    return False, "Không thể lưu tệp cấu hình."


# ── Image Compression & Upload ────────────────────────────────────────────────
def compress_image_bytes(image_data: bytes, max_dim: int = 1920, quality: int = 85) -> tuple[bytes, str]:
    """
    Compresses image bytes to optimal JPEG/WebP format while keeping high resolution.
    Returns (compressed_bytes, mime_type).
    """
    try:
        img = Image.open(io.BytesIO(image_data))
        
        # Handle orientation if present
        try:
            from PIL import ImageOps
            img = ImageOps.exif_transpose(img)
        except Exception:
            pass

        # Convert palette/transparency to RGB
        if img.mode in ("RGBA", "LA", "P"):
            bg = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "P":
                img = img.convert("RGBA")
            bg.paste(img, mask=img.split()[-1] if len(img.split()) > 3 else None)
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")

        # Resize if larger than max_dim
        w, h = img.size
        if max(w, h) > max_dim:
            if w > h:
                new_w = max_dim
                new_h = int(h * (max_dim / w))
            else:
                new_h = max_dim
                new_w = int(w * (max_dim / h))
            img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

        out_io = io.BytesIO()
        img.save(out_io, format="JPEG", quality=quality, optimize=True)
        return out_io.getvalue(), "image/jpeg"
    except Exception as e:
        print(f"[AnomalyService] Image compression warning: {e}")
        return image_data, "image/jpeg"


def upload_image(image_bytes: bytes, original_name: str = "defect.jpg") -> str:
    """
    Uploads compressed image to Supabase Storage bucket 'anomaly-images'.
    Returns public URL string on success, or raises Exception on failure.
    """
    cfg = load_config()
    supabase_url = cfg.get("supabase_url", "").rstrip("/")
    key = cfg.get("service_key") or cfg.get("anon_key", "")
    bucket = cfg.get("bucket_name", "anomaly-images")

    if not supabase_url or not key:
        raise ValueError("Chưa cấu hình Supabase URL hoặc API Key!")

    compressed_bytes, mime_type = compress_image_bytes(image_bytes)

    # Generate unique filename
    ext = ".jpg"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    uid = uuid.uuid4().hex[:6]
    clean_name = os.path.splitext(os.path.basename(original_name))[0]
    clean_name = "".join(c for c in clean_name if c.isalnum() or c in ("-", "_"))[:20] or "img"
    remote_filename = f"{ts}_{clean_name}_{uid}{ext}"

    upload_url = f"{supabase_url}/storage/v1/object/{bucket}/{remote_filename}"
    req = urllib.request.Request(
        upload_url,
        data=compressed_bytes,
        headers={
            "apikey": key,
            "Authorization": f"Bearer {key}",
            "Content-Type": mime_type
        },
        method="POST"
    )

    with urllib.request.urlopen(req, timeout=15, context=get_ssl_context()) as resp:
        if resp.status not in (200, 201):
            raise RuntimeError(f"Tải ảnh thất bại: HTTP {resp.status}")

    public_url = f"{supabase_url}/storage/v1/object/public/{bucket}/{remote_filename}"
    return public_url


def delete_image_from_storage(image_url: str) -> bool:
    """Deletes the remote image file from Supabase Storage given its public URL."""
    if not image_url or "anomaly-images" not in image_url:
        return False
    try:
        cfg = load_config()
        supabase_url = cfg.get("supabase_url", "").rstrip("/")
        key = cfg.get("service_key") or cfg.get("anon_key", "")
        bucket = cfg.get("bucket_name", "anomaly-images")

        # Extract filename from URL
        parts = image_url.split(f"/{bucket}/")
        if len(parts) < 2:
            return False
        filename = parts[1].split("?")[0]

        del_url = f"{supabase_url}/storage/v1/object/{bucket}"
        payload = json.dumps({"prefixes": [filename]}).encode("utf-8")
        req = urllib.request.Request(
            del_url,
            data=payload,
            headers={
                "apikey": key,
                "Authorization": f"Bearer {key}",
                "Content-Type": "application/json"
            },
            method="DELETE"
        )
        with urllib.request.urlopen(req, timeout=10, context=get_ssl_context()) as resp:
            return resp.status in (200, 204)
    except Exception as e:
        print(f"[AnomalyService] Delete image warning: {e}")
        return False


# ── Database Operations (CRUD) ────────────────────────────────────────────────
def _db_request(path: str, method: str = "GET", data: dict = None, params: dict = None) -> list | dict:
    """Sends authenticated HTTP request to Supabase PostgREST API."""
    cfg = load_config()
    supabase_url = cfg.get("supabase_url", "").rstrip("/")
    key = cfg.get("service_key") or cfg.get("anon_key", "")
    if not supabase_url or not key:
        raise ValueError("Chưa cấu hình Supabase URL hoặc API Key.")

    url = f"{supabase_url}/rest/v1/{path}"
    if params:
        q = urllib.parse.urlencode(params)
        url += f"{'&' if '?' in url else '?'}{q}"

    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Prefer": "return=representation"  # Returns created/updated row
    }

    body = json.dumps(data).encode("utf-8") if data is not None else None
    req = urllib.request.Request(url, data=body, headers=headers, method=method)

    try:
        with urllib.request.urlopen(req, timeout=15, context=get_ssl_context()) as resp:
            content = resp.read().decode("utf-8")
            if not content.strip():
                return []
            return json.loads(content)
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Database Error {e.code}: {err_body}")


def fetch_all_reports() -> list[dict]:
    """Fetches all anomaly reports ordered by ID descending."""
    res = _db_request("anomaly_reports?select=*&order=id.desc")
    return res if isinstance(res, list) else []


def create_report(report_data: dict, image_bytes: bytes = None, original_image_name: str = None) -> dict:
    """
    Creates a new Anomaly Report.
    Uploads image if provided, calculates defect_rate %, and inserts into DB.
    """
    data = dict(report_data)

    # Calculate defect rate %
    try:
        tot = float(data.get("total_qty") or 0)
        def_q = float(data.get("defect_qty") or 0)
        if tot > 0:
            data["defect_rate"] = round((def_q / tot) * 100.0, 2)
        else:
            data["defect_rate"] = 0.0
    except Exception:
        data["defect_rate"] = 0.0

    # Upload image if provided
    if image_bytes:
        img_name = original_image_name or "defect.jpg"
        img_url = upload_image(image_bytes, img_name)
        data["image_url"] = img_url

    # Ensure date
    if not data.get("report_date"):
        data["report_date"] = datetime.now().strftime("%Y-%m-%d")

    res = _db_request("anomaly_reports", method="POST", data=data)
    if isinstance(res, list) and res:
        return res[0]
    return data


def update_report(report_id: int, report_data: dict, new_image_bytes: bytes = None, new_image_name: str = None) -> dict:
    """
    Updates an existing Anomaly Report.
    Replaces image if new_image_bytes is provided.
    """
    data = dict(report_data)

    # Re-calculate defect rate
    try:
        tot = float(data.get("total_qty") or 0)
        def_q = float(data.get("defect_qty") or 0)
        if tot > 0:
            data["defect_rate"] = round((def_q / tot) * 100.0, 2)
        else:
            data["defect_rate"] = 0.0
    except Exception:
        pass

    # Remove id and created_at from update payload if present
    data.pop("id", None)
    data.pop("created_at", None)

    # Upload new image if provided
    if new_image_bytes:
        old_url = data.get("image_url")
        img_name = new_image_name or "defect_update.jpg"
        new_url = upload_image(new_image_bytes, img_name)
        data["image_url"] = new_url
        # Clean up old image
        if old_url:
            delete_image_from_storage(old_url)

    res = _db_request(f"anomaly_reports?id=eq.{report_id}", method="PATCH", data=data)
    if isinstance(res, list) and res:
        return res[0]
    return data


def delete_report(report_id: int, image_url: str = None) -> bool:
    """Deletes an anomaly report and its associated cloud image."""
    try:
        _db_request(f"anomaly_reports?id=eq.{report_id}", method="DELETE")
        if image_url:
            delete_image_from_storage(image_url)
        return True
    except Exception as e:
        print(f"[AnomalyService] Error deleting report: {e}")
        return False


# ── Excel Export ─────────────────────────────────────────────────────────────
def export_reports_to_excel(reports: list[dict], output_path: str) -> str:
    """
    Exports list of anomaly reports to a publication-ready Excel spreadsheet (.xlsx).
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Báo Cáo Bất Thường"
    ws.views.sheetView[0].showGridLines = True

    # 18 Columns matching user requirements
    headers = [
        ("STT", 6),
        ("Ngày Tháng", 13),
        ("Công Đoạn", 14),
        ("Sản Phẩm", 18),
        ("Máy Móc", 14),
        ("Số Lượng", 12),
        ("SL Lỗi", 10),
        ("Tỷ Lệ Lỗi (%)", 14),
        ("Người Chịu TN", 18),
        ("Người Phụ Trách", 18),
        ("Hình Ảnh Lỗi (URL)", 25),
        ("Mô Tả Lỗi", 35),
        ("Nguyên Nhân", 35),
        ("Biện Pháp Cải Tiến", 35),
        ("Tiêu Chuẩn SOP", 22),
        ("Tiến Độ", 16),
        ("Ghi Chú", 25)
    ]

    # Title block
    ws.merge_cells("A1:Q1")
    t_cell = ws["A1"]
    t_cell.value = "HỆ THỐNG QUẢN LÝ & BÁO CÁO BẤT THƯỜNG (ANOMALY REPORT) — VIPQC AI"
    t_cell.font = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    t_cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    # Subtitle
    ws.merge_cells("A2:Q2")
    sub_cell = ws["A2"]
    sub_cell.value = f"Thời gian xuất: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}   |   Tổng số bản ghi: {len(reports)}"
    sub_cell.font = Font(name="Segoe UI", size=10, italic=True, color="64748B")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 22

    # Header row
    hdr_row = 4
    ws.row_dimensions[hdr_row].height = 30
    hdr_fill = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid")
    hdr_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    for col_idx, (hdr_text, col_width) in enumerate(headers, 1):
        cell = ws.cell(row=hdr_row, column=col_idx, value=hdr_text)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = col_width

    # Data rows
    data_font = Font(name="Segoe UI", size=10)
    odd_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    even_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    for r_idx, rep in enumerate(reports, hdr_row + 1):
        ws.row_dimensions[r_idx].height = 26
        fill = even_fill if r_idx % 2 == 0 else odd_fill

        stt = r_idx - hdr_row
        d_date = str(rep.get("report_date") or "")
        proc = str(rep.get("process") or "")
        prod = str(rep.get("product_name") or "")
        mach = str(rep.get("machine") or "")
        tot_q = rep.get("total_qty", 0)
        def_q = rep.get("defect_qty", 0)
        rate = rep.get("defect_rate", 0.0)
        resp_p = str(rep.get("responsible_person") or "")
        pic = str(rep.get("pic") or "")
        img_url = str(rep.get("image_url") or "")
        desc = str(rep.get("description") or "")
        cause = str(rep.get("root_cause") or "")
        counter = str(rep.get("countermeasures") or "")
        sop = str(rep.get("sop_standard") or "")
        prog = str(rep.get("progress") or "Đang thực hiện")
        notes = str(rep.get("notes") or "")

        row_values = [
            (stt, "center"),
            (d_date, "center"),
            (proc, "center"),
            (prod, "left"),
            (mach, "center"),
            (tot_q, "right"),
            (def_q, "right"),
            (f"{rate:.2f}%" if rate is not None else "0.00%", "right"),
            (resp_p, "left"),
            (pic, "left"),
            (img_url, "left"),
            (desc, "left"),
            (cause, "left"),
            (counter, "left"),
            (sop, "center"),
            (prog, "center"),
            (notes, "left")
        ]

        for c_idx, (val, align) in enumerate(row_values, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = data_font
            cell.fill = fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal=align, vertical="center")

            # Clickable hyperlink for image
            if c_idx == 11 and val and str(val).startswith("http"):
                cell.hyperlink = val
                cell.font = Font(name="Segoe UI", size=10, color="0D9488", underline="single")

            # Style progress column
            if c_idx == 16:
                if "hoàn thành" in prog.lower() or "done" in prog.lower():
                    cell.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="166534")
                elif "đang" in prog.lower() or "progress" in prog.lower():
                    cell.fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="854D0E")
                elif "chưa" in prog.lower() or "open" in prog.lower():
                    cell.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
                    cell.font = Font(name="Segoe UI", size=10, bold=True, color="991B1B")

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    return output_path
