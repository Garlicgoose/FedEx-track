# -*- coding: utf-8 -*-
"""FedEx tracking via the official FedEx Tracking API (no browser).

Why API: FedEx gates the website behind Akamai - scripted browsers get
soft-blocked and plain requests get 403. The official API
(developer.fedex.com, free) returns structured status, exact delivery
timestamp, child-piece numbers and the signature proof-of-delivery PDF.

Child-piece rule (user requirement, 2026-09 修订):
- /track/v1/trackingnumbers 正常查主单（非 MPS）；无子单运单以官网状态为准
  （assoc 接口对无子单运单可能返回 unknown——用户实测 bug，已修复）。
- /track/v1/associatedshipments 探测子单：返回主单+子单，最多 40 条。
- 有子单且 <40：全部可见，全 Delivered 才算送达（主单自身状态不可单独信任）。
- 查到 40 条（可能有隐藏子单）：40 条全 Delivered 则默认送达，
  备注标注"子单超过40，需人工查询"；否则未送达并列出未送达子单。
- POD（signatureProofOfDelivery PDF）：仅送达时下载，命名 = 主单号.pdf。
"""
import base64
import time
from datetime import datetime
from pathlib import Path

import requests

TOKEN_URL = "https://apis.fedex.com/oauth/token"
TRACK_URL = "https://apis.fedex.com/track/v1/trackingnumbers"
ASSOC_URL = "https://apis.fedex.com/track/v1/associatedshipments"
DOC_URL = "https://apis.fedex.com/track/v1/trackingdocuments"

PDF_DIR = None  # 由调用方设置（输出目录/pdf/FedEx）
FEDEX_RELATED_LIMIT = 40   # API 返回件数上限（含主单自身）

_token_cache = {"token": "", "expires": 0.0}


def _get_token(api_key, api_secret):
    """OAuth client-credentials token, cached until ~60s before expiry."""
    now = time.time()
    if _token_cache["token"] and _token_cache["expires"] > now + 60:
        return _token_cache["token"]
    resp = requests.post(
        TOKEN_URL,
        data={
            "grant_type": "client_credentials",
            "client_id": api_key,
            "client_secret": api_secret,
        },
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        timeout=30)
    if resp.status_code != 200:
        raise RuntimeError("FedEx auth failed (HTTP {}): {}".format(
            resp.status_code, resp.text[:200]))
    data = resp.json()
    _token_cache["token"] = data.get("access_token", "")
    _token_cache["expires"] = now + float(data.get("expires_in", 3600))
    if not _token_cache["token"]:
        raise RuntimeError("FedEx auth returned no access token")
    return _token_cache["token"]


def _fmt_delivery(dt_str):
    """'2026-07-20T12:59:00-04:00' -> '2026-07-20 12:59'."""
    if not dt_str:
        return ""
    s = str(dt_str)
    try:
        return datetime.fromisoformat(s).strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return s[:16] if len(s) >= 16 else s


def _latest_status(piece):
    """返回 (code, description) 取 latestStatusDetail。"""
    latest = piece.get("latestStatusDetail") or {}
    code = str(latest.get("code") or "")
    desc = (latest.get("statusByLocale")
            or latest.get("description") or "")
    return code, desc


def _is_delivered(piece):
    """单个 piece 是否 Delivered。"""
    code, desc = _latest_status(piece)
    return (code.upper() == "DL"
            or str(desc).strip().casefold() == "delivered")


def _actual_delivery_times(piece):
    """取该 piece 的 ACTUAL_DELIVERY 时间列表。"""
    times = []
    for date_item in (piece.get("dateAndTimes") or []):
        if str(date_item.get("type") or "").upper() == "ACTUAL_DELIVERY":
            dt = date_item.get("dateTime")
            if dt:
                times.append(dt)
    if not times:
        latest = piece.get("latestStatusDetail") or {}
        dt = latest.get("dateTime") or ""
        if dt:
            times.append(dt)
    return times


def _query_trackingnumbers(token, tracking_number, timeout=40):
    """正常查找（非 MPS）：trackingnumbers 直接查主单。

    返回主单 piece dict；查不到返回 None。
    """
    payload = {
        "includeDetailedScans": False,
        "trackingInfo": [{
            "trackingNumberInfo": {"trackingNumber": tracking_number},
        }],
    }
    resp = requests.post(
        TRACK_URL,
        headers={
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json",
            "X-locale": "en_US",
        },
        json=payload,
        timeout=timeout)
    if resp.status_code != 200:
        return None
    data = resp.json()
    ctrs = ((data.get("output") or {}).get("completeTrackResults") or [])
    if not ctrs:
        return None
    results = ctrs[0].get("trackResults") or []
    return results[0] if results else None


def _query_assoc(token, tracking_number, timeout=40):
    """查询 associatedshipments，返回 pieces 列表（最多 40 条，含主单）。

    返回 (pieces, error)。pieces[0] 通常为主单。
    """
    payload = {
        "includeDetailedScans": False,
        "associatedType": "STANDARD_MPS",
        "masterTrackingNumberInfo": {
            "trackingNumberInfo": {"trackingNumber": tracking_number},
        },
    }
    resp = requests.post(
        ASSOC_URL,
        headers={
            "Authorization": "Bearer " + token,
            "Content-Type": "application/json",
            "X-locale": "en_US",
        },
        json=payload,
        timeout=timeout)
    if resp.status_code != 200:
        return [], "assoc API HTTP {}".format(resp.status_code)
    data = resp.json()
    ctrs = ((data.get("output") or {}).get("completeTrackResults") or [])
    if not ctrs:
        return [], "no completeTrackResults in response"
    return ctrs[0].get("trackResults") or [], ""


def _save_pod(tracking_number, token, master_piece, pdf_dir):
    """Signature proof of delivery as PDF named <tracking_number>.pdf."""
    if not pdf_dir:
        return ""
    out_dir = Path(pdf_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    headers = {
        "Authorization": "Bearer " + token,
        "Content-Type": "application/json",
        "X-locale": "en_US",
    }
    tcn = master_piece.get("trackControlNumber") or ""
    ts = master_piece.get("shipmentTimestamp") or ""
    if tcn and ts:
        try:
            resp = requests.post(
                DOC_URL,
                json={
                    "trackingNumberInfo": {"trackingNumber": tracking_number},
                    "trackingControlNumber": tcn,
                    "shipmentTimestamp": ts,
                    "documents": [{
                        "type": "signatureProofOfDelivery",
                        "imageType": "PDF",
                    }],
                },
                headers=headers, timeout=40)
            if resp.status_code == 200:
                docs = ((resp.json().get("output") or {})
                        .get("documents") or [])
                content = docs[0].get("content") if docs else ""
                if content:
                    p = out_dir / "{}.pdf".format(tracking_number)
                    p.write_bytes(base64.b64decode(content))
                    return str(p)
        except Exception:
            pass
    return ""


def _find_master(pieces, tracking_number):
    for piece in pieces:
        tn = (piece.get("trackingNumberInfo") or {}).get("trackingNumber", "")
        if tn == tracking_number:
            return piece
    return pieces[0] if pieces else None


def query_fedex_one(tracking_number, api_key="", api_secret="",
                    save_pdf=True, pdf_dir=None):
    """查询一个 FedEx 运单（官方 API，2026-09 修订逻辑）。

    流程：
      1. trackingnumbers 正常查主单（非 MPS）——无子单运单以官网状态为准；
      2. associatedshipments 探测子单：只有主单自己 → 无子单，用第 1 步状态；
      3. 有子单：<40 全查，子单全 DL 才算送达；
         查到 40 条（可能有隐藏子单）：40 条全 DL 则默认送达，
         备注标注"子单超过40，需人工查询"，否则未送达+明细；
      4. POD（签名 PDF）：送达才下载，命名 = 主单号.pdf。

    返回 dict:
        tracking_number / status / is_delivered ("Y"或"") / delivery_date /
        arrival_time / pdf_file / error / flag(备注)
    """
    result = {
        "tracking_number": tracking_number,
        "status": "",
        "is_delivered": "",
        "delivery_date": "",
        "arrival_time": "",
        "pdf_file": "",
        "error": "",
        "flag": "",
    }
    if not tracking_number:
        result["error"] = "Empty tracking number"
        return result
    if not api_key or not api_secret:
        result["error"] = ("FedEx API key/secret not set "
                           "(Settings > FedEx API)")
        return result

    try:
        token = _get_token(api_key, api_secret)

        # ---- 1. 正常查找（非 MPS）：trackingnumbers 查主单 ----
        main_piece = _query_trackingnumbers(token, tracking_number)

        # ---- 2. MPS 探测子单 ----
        pieces, err = _query_assoc(token, tracking_number)
        if err or not pieces:
            pieces = []
        master = _find_master(pieces, tracking_number) if pieces else None
        if master is None:
            master = main_piece
        piece_count = len(pieces)

        # 无子单（assoc 查不到 / 只返回主单自己 / assoc 失败）：
        # 以 trackingnumbers 主单状态为准（官网一致，2026-09 修复：
        # 无子单运单 assoc 可能返回 unknown）
        if piece_count <= 1:
            if main_piece is None:
                if master is None:
                    result["error"] = err or "no track results"
                    result["status"] = "Not Found"
                    return result
                main_piece = master
            code, desc = _latest_status(main_piece)
            result["status"] = desc or code or "Unknown"
            if _is_delivered(main_piece):
                result["is_delivered"] = "Y"
                times = _actual_delivery_times(main_piece)
                if times:
                    result["delivery_date"] = _fmt_delivery(times[-1])
                    result["arrival_time"] = result["delivery_date"]
                if save_pdf:
                    result["pdf_file"] = _save_pod(
                        tracking_number, token, main_piece, pdf_dir)
            return result

        if piece_count >= FEDEX_RELATED_LIMIT:
            # ---- 查到 40 条：可能有隐藏子单 ----
            # 40 条全部送达 → 默认送达 + 备注人工核查（2026-09 用户修订）
            undelivered = []
            for piece in pieces:
                if not _is_delivered(piece):
                    _c, dsc = _latest_status(piece)
                    tn = (piece.get("trackingNumberInfo") or {}) \
                        .get("trackingNumber", "")
                    undelivered.append((tn, dsc or _c or "Unknown"))
            result["flag"] = ("子单超过{}，需人工查询".format(
                FEDEX_RELATED_LIMIT - 1))
            if not undelivered:
                result["status"] = "Delivered"
                result["is_delivered"] = "Y"
                all_times = []
                for piece in pieces:
                    all_times.extend(_actual_delivery_times(piece))
                if all_times:
                    result["delivery_date"] = _fmt_delivery(all_times[-1])
                    result["arrival_time"] = result["delivery_date"]
                if save_pdf:
                    result["pdf_file"] = _save_pod(
                        tracking_number, token, master or main_piece, pdf_dir)
            else:
                result["status"] = undelivered[0][1]
                result["is_delivered"] = ""
                shown = undelivered[:5]
                result["flag"] += "；未送达子单 {} 件：{}".format(
                    len(undelivered),
                    "; ".join("{}={}".format(tn, st) for tn, st in shown)
                    + ("..." if len(undelivered) > 5 else ""))
            return result

        # ---- 件数 < 40 且有子单：全部可见，全部送达才算送达 ----
        undelivered = []
        for piece in pieces:
            if not _is_delivered(piece):
                _c, desc = _latest_status(piece)
                tn = (piece.get("trackingNumberInfo") or {}) \
                    .get("trackingNumber", "")
                undelivered.append((tn, desc or _c or "Unknown"))

        if not undelivered:
            result["status"] = "Delivered"
            result["is_delivered"] = "Y"
            # 取所有件 ACTUAL_DELIVERY 中最晚的时间
            all_times = []
            for piece in pieces:
                all_times.extend(_actual_delivery_times(piece))
            if all_times:
                result["delivery_date"] = _fmt_delivery(all_times[-1])
                result["arrival_time"] = result["delivery_date"]
            if save_pdf:
                result["pdf_file"] = _save_pod(
                    tracking_number, token, master or main_piece, pdf_dir)
        else:
            result["status"] = undelivered[0][1]
            result["is_delivered"] = ""
            shown = undelivered[:5]
            result["flag"] = ("未送达子单 {} 件：{}".format(
                len(undelivered),
                "; ".join("{}={}".format(tn, st) for tn, st in shown)
                + ("..." if len(undelivered) > 5 else "")))
        return result

    except requests.RequestException as exc:
        result["error"] = "network error: {}".format(exc)
        return result
    except Exception as exc:  # noqa: BLE001
        result["error"] = str(exc)
        return result


if __name__ == "__main__":
    # 手动测试：python modules/fedex_module.py <tracking_number>
    import sys
    num = sys.argv[1] if len(sys.argv) > 1 else "472887355535"
    KEY = "l726270ca1e3e14d2092057988b4aee3db"
    SECRET = "ea3249b6848d41ddae192a8111e4afd0"
    r = query_fedex_one(num, KEY, SECRET, save_pdf=False)
    for k, v in r.items():
        print(f"{k}: {v}")
