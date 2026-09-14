#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用持久化 Chromium/Edge 会话保存 FedEx 查询页和详情页 PDF。

每个已送达运单生成两个文件：
    <运单号>.pdf   查询结果主页面
    <运单号>+.pdf  点击详情后的页面

安装依赖：python -m pip install playwright
首次运行建议保持有界面模式，人工处理 FedEx 可能显示的验证码或提示。程序不会
尝试绕过网站安全校验；浏览器资料保存在 ``.fedex_browser``，以后会继续复用。
"""

from __future__ import annotations

import argparse
import html
import json
import os
import re
import sys
import time
import unittest
from dataclasses import asdict, dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Iterable, Sequence

import requests


TRACKING_RE = re.compile(r"^[A-Za-z0-9]{8,30}$")
DELIVERED_RE = re.compile(r"(?i)(?<![A-Za-z])delivered(?![A-Za-z])|已送达|递送完毕")
ERROR_MARKERS = (
    "we can’t find that tracking number",
    "we can't find that tracking number",
    "找不到该追踪号码",
    "无法找到该追踪号码",
    "tracking number is invalid",
    "system error",
    "temporarily unavailable",
)
DETAIL_LABEL_RE = re.compile(
    r"view\s+(?:more\s+)?details?|shipment\s+details?|"
    r"travel\s+history|查看(?:更多)?详情|货件详情|详细信息|行程历史",
    re.I,
)
DETAIL_CONTENT_RE = re.compile(
    r"(?i)travel\s+history|shipment\s+facts|shipment\s+details|"
    r"delivery\s+details|行程历史|物流记录|货件事实|货件详情|递送详情"
)
TRACKING_PAGE = "https://www.fedex.com/zh-cn/tracking.html"
TRACKING_RESULT_URL = "https://www.fedex.com/fedextrack/?trknbr={}"
FEDEX_BASE_URL = "https://apis.fedex.com"


class FedExPODError(RuntimeError):
    """单票处理失败。"""


@dataclass(frozen=True)
class DownloadResult:
    tracking_number: str
    ok: bool
    main_pdf: str = ""
    detail_pdf: str = ""
    message: str = ""


@dataclass(frozen=True)
class ShipmentSnapshot:
    tracking_number: str
    status: str
    delivered_at: str
    signed_by: str
    origin: str
    destination: str
    service: str
    weight: str
    pieces: tuple[tuple[str, str], ...]
    events: tuple[tuple[str, str, str], ...]


def normalize_tracking_number(value: str) -> str:
    number = re.sub(r"[\s-]+", "", str(value or "").strip())
    if not TRACKING_RE.fullmatch(number):
        raise ValueError(f"运单号格式不正确：{value!r}")
    return number


def output_paths(output_dir: Path, tracking_number: str) -> tuple[Path, Path]:
    number = normalize_tracking_number(tracking_number)
    return output_dir / f"{number}.pdf", output_dir / f"{number}+.pdf"


def is_valid_pdf(path: Path) -> bool:
    try:
        return path.is_file() and path.stat().st_size >= 1024 and path.read_bytes()[:5] == b"%PDF-"
    except OSError:
        return False


def find_browser(explicit: str = "") -> str | None:
    values = [
        explicit,
        os.getenv("FEDEX_BROWSER_PATH", ""),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    ]
    for value in values:
        if value and Path(value).is_file():
            return str(Path(value).resolve())
    if explicit:
        raise FileNotFoundError(f"找不到浏览器：{explicit}")
    return None


def load_shipmenttrack_credentials(project_dir: Path) -> None:
    """可选复用本机 ShipmentTrack 的 DPAPI 加密设置，不打印或写出密钥。"""
    root = project_dir.resolve()
    if not (root / "modules" / "settings_store.py").is_file():
        raise FedExPODError(f"不是有效的 ShipmentTrack 目录：{root}")
    sys.path.insert(0, str(root))
    try:
        from modules.settings_store import SettingsStore

        settings = SettingsStore().load_settings()
        os.environ["FEDEX_API_KEY"] = str(settings.get("fedex_api_key") or "").strip()
        os.environ["FEDEX_API_SECRET"] = str(settings.get("fedex_api_secret") or "").strip()
    except Exception as exc:
        raise FedExPODError(f"无法读取 ShipmentTrack 的 FedEx 设置：{exc}") from exc


def _body_text(page) -> str:
    try:
        return page.locator("body").inner_text(timeout=5_000)
    except Exception:
        return ""


def _dismiss_cookie_banner(page) -> None:
    labels = re.compile(
        r"reject optional cookies|accept all cookies|拒绝可选|仅使用必要|接受全部",
        re.I,
    )
    for role in ("button", "link"):
        locator = page.get_by_role(role, name=labels)
        if locator.count():
            try:
                locator.first.click(timeout=2_000)
                return
            except Exception:
                pass


def _submit_tracking(page, tracking_number: str, timeout_ms: int) -> None:
    # 主查询结果 URL 是 FedEx 官网公开入口；详情仍只点击网页按钮进入，
    # 不猜测依赖 session/trkqual 的详情地址。
    page.goto(
        TRACKING_RESULT_URL.format(tracking_number),
        wait_until="domcontentloaded",
        timeout=timeout_ms,
    )
    _dismiss_cookie_banner(page)
    page.wait_for_timeout(5_000)
    direct_text = _body_text(page).casefold()
    direct_failed = "/system-error" in page.url or any(
        marker in direct_text for marker in ERROR_MARKERS
    )
    if not direct_failed:
        return

    # 个别新浏览器资料会把直接入口导向 system-error，或在正文提示找不到运单；
    # 此时从中国区首页填写表单重试。
    page.goto(TRACKING_PAGE, wait_until="domcontentloaded", timeout=timeout_ms)
    _dismiss_cookie_banner(page)

    box = page.get_by_role("textbox", name=re.compile(r"追踪号码|tracking number", re.I))
    if not box.count() or not box.first.is_visible():
        # 页面同时存在桌面/移动版隐藏控件，只允许选择当前可见输入框。
        box = page.locator(
            "textarea[name*='tracking']:visible, input[name*='tracking']:visible, "
            "textarea[id*='tracking']:visible, input[id*='tracking']:visible"
        )
    if not box.count():
        # 首页访问本身会建立地区与功能开关 Cookie。某些新会话首次不渲染
        # 查询组件，再次打开公开结果入口即可正常加载。
        page.goto(
            TRACKING_RESULT_URL.format(tracking_number),
            wait_until="domcontentloaded",
            timeout=timeout_ms,
        )
        return

    box.first.fill(tracking_number)
    button = page.get_by_role(
        "button", name=re.compile(r"^(?:货件查询|追踪|track|track shipment)$", re.I)
    )
    if not button.count():
        button = page.locator("button[type='submit']:visible")
    if not button.count():
        raise FedExPODError("FedEx 查询按钮未出现。")
    button.first.click(timeout=10_000)


def _wait_for_delivered(page, timeout_ms: int) -> str:
    deadline = time.monotonic() + timeout_ms / 1000
    last_text = ""
    while time.monotonic() < deadline:
        last_text = _body_text(page)
        folded = last_text.casefold()
        if any(marker in folded for marker in ERROR_MARKERS):
            raise FedExPODError("FedEx 返回“找不到运单/系统错误”；未生成 PDF。")
        if DELIVERED_RE.search(last_text):
            return last_text
        page.wait_for_timeout(750)

    excerpt = re.sub(r"\s+", " ", last_text).strip()[:180]
    if excerpt:
        raise FedExPODError(f"等待送达状态超时；当前页面摘要：{excerpt}")
    raise FedExPODError("等待送达状态超时；FedEx 页面没有返回可识别内容。")


def _print_pdf(page, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_name(f".{destination.name}.{os.getpid()}.tmp")
    try:
        page.emulate_media(media="screen")
        page.pdf(
            path=str(temporary),
            format="A4",
            print_background=True,
            margin={"top": "8mm", "right": "8mm", "bottom": "8mm", "left": "8mm"},
        )
        if not is_valid_pdf(temporary):
            raise FedExPODError(f"浏览器生成的 PDF 无效：{destination.name}")
        os.replace(temporary, destination)
    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass


def _open_details(page, previous_text: str, timeout_ms: int) -> None:
    candidates = []
    for role in ("button", "link"):
        locator = page.get_by_role(role, name=DETAIL_LABEL_RE)
        candidates.extend(locator.nth(index) for index in range(min(locator.count(), 12)))

    if not candidates:
        text_locator = page.get_by_text(DETAIL_LABEL_RE)
        candidates.extend(text_locator.nth(index) for index in range(min(text_locator.count(), 12)))

    deadline = time.monotonic() + timeout_ms / 1000
    for candidate in candidates:
        try:
            if not candidate.is_visible():
                continue
            candidate.click(timeout=5_000)
        except Exception:
            continue

        while time.monotonic() < deadline:
            current = _body_text(page)
            if DETAIL_CONTENT_RE.search(current) and current != previous_text:
                return
            page.wait_for_timeout(500)
        break
    raise FedExPODError("已保存主页面，但未能打开 FedEx 详情页；未生成“运单号+.pdf”。")


class FedExAPIClient:
    """只读调用 FedEx Track API，用于网页自动化受限时生成归档快照。"""

    def __init__(self, api_key: str, api_secret: str, timeout_seconds: int = 45) -> None:
        if not api_key or not api_secret:
            raise FedExPODError("API 模式需要 FEDEX_API_KEY 和 FEDEX_API_SECRET 环境变量。")
        self.api_key = api_key
        self.api_secret = api_secret
        self.timeout = timeout_seconds
        self.session = requests.Session()
        self.token = self._get_token()

    def _post(self, path: str, **kwargs):
        last_error: Exception | None = None
        for attempt in range(4):
            try:
                response = self.session.post(
                    f"{FEDEX_BASE_URL}{path}", timeout=self.timeout, **kwargs
                )
                if response.status_code not in {429, 500, 502, 503, 504}:
                    return response
            except (requests.Timeout, requests.ConnectionError) as exc:
                last_error = exc
            if attempt < 3:
                time.sleep(1.5 * (2**attempt))
        if last_error:
            raise FedExPODError(f"FedEx API 网络错误：{last_error}") from last_error
        return response

    def _get_token(self) -> str:
        response = self._post(
            "/oauth/token",
            data={
                "grant_type": "client_credentials",
                "client_id": self.api_key,
                "client_secret": self.api_secret,
            },
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        if response.status_code != 200:
            raise FedExPODError(f"FedEx API 授权失败（HTTP {response.status_code}）。")
        token = str(response.json().get("access_token") or "").strip()
        if not token:
            raise FedExPODError("FedEx API 授权响应中没有 access_token。")
        return token

    @property
    def headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
            "X-locale": "zh_CN",
        }

    @staticmethod
    def _results(response) -> list[dict]:
        if response.status_code != 200:
            raise FedExPODError(f"FedEx Track API 查询失败（HTTP {response.status_code}）。")
        try:
            complete = (response.json().get("output") or {}).get("completeTrackResults") or []
            return list((complete[0] or {}).get("trackResults") or []) if complete else []
        except (TypeError, ValueError) as exc:
            raise FedExPODError("FedEx Track API 返回了无法解析的数据。") from exc

    def _tracking_piece(self, number: str) -> dict:
        response = self._post(
            "/track/v1/trackingnumbers",
            headers=self.headers,
            json={
                "includeDetailedScans": True,
                "trackingInfo": [{"trackingNumberInfo": {"trackingNumber": number}}],
            },
        )
        pieces = self._results(response)
        for piece in pieces:
            returned = str(
                ((piece or {}).get("trackingNumberInfo") or {}).get("trackingNumber") or ""
            )
            if returned == number:
                return piece
        if pieces:
            return pieces[0]
        raise FedExPODError("FedEx Track API 找不到该运单。")

    def _associated_pieces(self, number: str) -> list[dict]:
        response = self._post(
            "/track/v1/associatedshipments",
            headers=self.headers,
            json={
                "includeDetailedScans": False,
                "associatedType": "STANDARD_MPS",
                "masterTrackingNumberInfo": {
                    "trackingNumberInfo": {"trackingNumber": number}
                },
            },
        )
        return self._results(response) if response.status_code == 200 else []

    @staticmethod
    def _status(piece: dict) -> str:
        latest = (piece or {}).get("latestStatusDetail") or {}
        return str(latest.get("statusByLocale") or latest.get("description") or "").strip()

    @staticmethod
    def _location(value: dict) -> str:
        location = value or {}
        address = location.get("locationContactAndAddress") or location
        address = address.get("address") or address
        return ", ".join(
            str(address.get(key) or "").strip()
            for key in ("city", "stateOrProvinceCode", "countryCode")
            if str(address.get(key) or "").strip()
        )

    def fetch(self, number: str) -> ShipmentSnapshot:
        piece = self._tracking_piece(number)
        status = self._status(piece)
        if not DELIVERED_RE.search(status):
            raise FedExPODError(f"运单尚未送达（当前状态：{status or '未知'}），未生成 PDF。")

        associated = self._associated_pieces(number)
        if len(associated) > 1:
            pending = [self._status(item) or "未知" for item in associated if not DELIVERED_RE.search(self._status(item))]
            if pending:
                raise FedExPODError(f"关联货件尚未全部送达（{len(pending)} 件），未生成 PDF。")
        pieces_source = associated or [piece]
        pieces = tuple(
            (
                str(((item.get("trackingNumberInfo") or {}).get("trackingNumber")) or ""),
                self._status(item),
            )
            for item in pieces_source
        )

        delivered_at = ""
        for item in piece.get("dateAndTimes") or []:
            if str(item.get("type") or "").upper() == "ACTUAL_DELIVERY":
                delivered_at = str(item.get("dateTime") or "")
                break
        delivery = piece.get("deliveryDetails") or {}
        signed_by = str(
            delivery.get("receivedByName")
            or delivery.get("signedByName")
            or delivery.get("locationDescription")
            or ""
        ).strip()
        origin = self._location(piece.get("shipperInformation") or {})
        destination = self._location(piece.get("recipientInformation") or {})
        service = str((piece.get("serviceDetail") or {}).get("description") or "").strip()
        weight_data = ((piece.get("packageDetails") or {}).get("weightAndDimensions") or {}).get("weight") or []
        weight = " / ".join(
            f"{item.get('value')} {item.get('unit')}" for item in weight_data if item.get("value")
        )
        events = tuple(
            (
                str(event.get("date") or ""),
                str(event.get("eventDescription") or event.get("derivedStatus") or "").strip(),
                self._location(event.get("scanLocation") or {}),
            )
            for event in reversed(piece.get("scanEvents") or [])
        )
        return ShipmentSnapshot(
            tracking_number=number,
            status=status,
            delivered_at=delivered_at,
            signed_by=signed_by,
            origin=origin,
            destination=destination,
            service=service,
            weight=weight,
            pieces=pieces,
            events=events,
        )


def _snapshot_html(snapshot: ShipmentSnapshot, *, detailed: bool) -> str:
    esc = lambda value: html.escape(str(value or ""))
    pieces = "".join(
        f"<tr><td>{esc(number)}</td><td>{esc(status)}</td></tr>"
        for number, status in snapshot.pieces
    )
    events = "".join(
        f"<tr><td>{esc(date)}</td><td>{esc(status)}</td><td>{esc(location)}</td></tr>"
        for date, status, location in snapshot.events
    )
    detail_section = ""
    if detailed:
        detail_section = f"""
        <h2>物流记录</h2>
        <table><thead><tr><th>日期时间</th><th>状态</th><th>地点</th></tr></thead>
        <tbody>{events or '<tr><td colspan="3">API 未返回扫描记录</td></tr>'}</tbody></table>
        <h2>关联货件</h2>
        <table><thead><tr><th>追踪号码</th><th>状态</th></tr></thead><tbody>{pieces}</tbody></table>
        """
    return f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
    <style>
      @page {{ size:A4; margin:14mm; }}
      body {{ font-family:'Microsoft YaHei','Segoe UI',sans-serif; color:#242424; font-size:12px; }}
      header {{ border-bottom:5px solid #4d148c; padding-bottom:12px; margin-bottom:22px; }}
      .brand {{ color:#4d148c; font-size:28px; font-weight:800; }}
      .tag {{ color:#666; margin-top:5px; }}
      .status {{ font-size:34px; color:#238636; font-weight:800; margin:18px 0 6px; }}
      .number {{ font-size:19px; letter-spacing:1px; }}
      .grid {{ display:grid; grid-template-columns:1fr 1fr; gap:10px 30px; margin:22px 0; }}
      .item {{ border-bottom:1px solid #ddd; padding:7px 0; }}
      .label {{ color:#666; display:block; font-size:10px; }}
      h2 {{ color:#4d148c; margin-top:24px; }}
      table {{ width:100%; border-collapse:collapse; page-break-inside:auto; }}
      th,td {{ text-align:left; border-bottom:1px solid #ddd; padding:7px 6px; vertical-align:top; }}
      th {{ background:#f4f0f8; }} tr {{ page-break-inside:avoid; }}
      footer {{ margin-top:28px; color:#777; font-size:9px; }}
    </style></head><body>
    <header><div class="brand">FedEx Tracking Snapshot</div>
    <div class="tag">数据来源：FedEx Track API · 这不是 FedEx 官方签名 POD</div></header>
    <div class="number">追踪号码：{esc(snapshot.tracking_number)}</div>
    <div class="status">{esc(snapshot.status)}</div>
    <div class="grid">
      <div class="item"><span class="label">送达时间</span>{esc(snapshot.delivered_at)}</div>
      <div class="item"><span class="label">签收人</span>{esc(snapshot.signed_by)}</div>
      <div class="item"><span class="label">寄件地</span>{esc(snapshot.origin)}</div>
      <div class="item"><span class="label">目的地</span>{esc(snapshot.destination)}</div>
      <div class="item"><span class="label">服务</span>{esc(snapshot.service)}</div>
      <div class="item"><span class="label">重量</span>{esc(snapshot.weight)}</div>
      <div class="item"><span class="label">关联件数</span>{len(snapshot.pieces)}</div>
    </div>{detail_section}
    <footer>生成时间：{time.strftime('%Y-%m-%d %H:%M:%S')} · 请按业务需要与 FedEx 官网复核。</footer>
    </body></html>"""


def archive_with_api(
    tracking_numbers: Iterable[str],
    output_dir: Path,
    *,
    browser_path: str = "",
    timeout_seconds: int = 45,
    overwrite: bool = False,
) -> list[DownloadResult]:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise FedExPODError("缺少 Playwright。请运行：python -m pip install playwright") from exc
    client = FedExAPIClient(
        os.getenv("FEDEX_API_KEY", "").strip(),
        os.getenv("FEDEX_API_SECRET", "").strip(),
        timeout_seconds,
    )
    directory = output_dir.resolve()
    directory.mkdir(parents=True, exist_ok=True)
    results: list[DownloadResult] = []
    with sync_playwright() as playwright:
        options = {"headless": True}
        executable = find_browser(browser_path)
        if executable:
            options["executable_path"] = executable
        browser = playwright.chromium.launch(**options)
        page = browser.new_page(viewport={"width": 1440, "height": 1000})
        try:
            for raw_number in tracking_numbers:
                number = normalize_tracking_number(raw_number)
                main_pdf, detail_pdf = output_paths(directory, number)
                if not overwrite and is_valid_pdf(main_pdf) and is_valid_pdf(detail_pdf):
                    results.append(DownloadResult(number, True, str(main_pdf), str(detail_pdf), "两个 PDF 已存在，已跳过。"))
                    continue
                try:
                    snapshot = client.fetch(number)
                    page.set_content(_snapshot_html(snapshot, detailed=False), wait_until="load")
                    _print_pdf(page, main_pdf)
                    page.set_content(_snapshot_html(snapshot, detailed=True), wait_until="load")
                    _print_pdf(page, detail_pdf)
                    results.append(DownloadResult(number, True, str(main_pdf), str(detail_pdf), "已用 FedEx API 数据生成主状态页和详细物流记录页。"))
                except Exception as exc:
                    results.append(DownloadResult(number, False, str(main_pdf) if is_valid_pdf(main_pdf) else "", str(detail_pdf) if is_valid_pdf(detail_pdf) else "", str(exc)))
        finally:
            browser.close()
            client.session.close()
    return results


class FedExPODDownloader:
    def __init__(
        self,
        output_dir: Path,
        profile_dir: Path,
        *,
        browser_path: str = "",
        headless: bool = False,
        timeout_seconds: int = 90,
        overwrite: bool = False,
    ) -> None:
        self.output_dir = output_dir.resolve()
        self.profile_dir = profile_dir.resolve()
        self.browser_path = find_browser(browser_path)
        self.headless = headless
        self.timeout_ms = timeout_seconds * 1000
        self.overwrite = overwrite

    def run(self, tracking_numbers: Iterable[str]) -> list[DownloadResult]:
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as exc:
            raise FedExPODError(
                "缺少 Playwright。请运行：python -m pip install playwright"
            ) from exc

        numbers = list(dict.fromkeys(normalize_tracking_number(n) for n in tracking_numbers))
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.profile_dir.mkdir(parents=True, exist_ok=True)
        results: list[DownloadResult] = []

        with sync_playwright() as playwright:
            launch_options = {
                "headless": self.headless,
                "locale": "zh-CN",
                "accept_downloads": True,
                "viewport": {"width": 1440, "height": 1000},
            }
            if self.browser_path:
                launch_options["executable_path"] = self.browser_path
            context = playwright.chromium.launch_persistent_context(
                str(self.profile_dir), **launch_options
            )
            context.set_default_timeout(min(self.timeout_ms, 30_000))
            page = context.pages[0] if context.pages else context.new_page()
            try:
                for number in numbers:
                    results.append(self._download_one(page, number))
            finally:
                context.close()
        return results

    def _download_one(self, page, number: str) -> DownloadResult:
        main_pdf, detail_pdf = output_paths(self.output_dir, number)
        if (
            not self.overwrite
            and is_valid_pdf(main_pdf)
            and is_valid_pdf(detail_pdf)
        ):
            return DownloadResult(
                number, True, str(main_pdf), str(detail_pdf), "两个 PDF 已存在，已跳过。"
            )

        try:
            _submit_tracking(page, number, self.timeout_ms)
            main_text = _wait_for_delivered(page, self.timeout_ms)
            _print_pdf(page, main_pdf)
            _open_details(page, main_text, self.timeout_ms)
            _print_pdf(page, detail_pdf)
            if not (is_valid_pdf(main_pdf) and is_valid_pdf(detail_pdf)):
                raise FedExPODError("生成后校验失败。")
            return DownloadResult(
                number, True, str(main_pdf), str(detail_pdf), "主页面和详情页均已保存。"
            )
        except Exception as exc:
            return DownloadResult(
                number,
                False,
                str(main_pdf) if is_valid_pdf(main_pdf) else "",
                str(detail_pdf) if is_valid_pdf(detail_pdf) else "",
                str(exc),
            )


class _SelfTests(unittest.TestCase):
    def test_normalize_tracking_number(self) -> None:
        self.assertEqual("123456789012", normalize_tracking_number("1234 5678-9012"))
        with self.assertRaises(ValueError):
            normalize_tracking_number("../../bad")

    def test_output_names(self) -> None:
        main, detail = output_paths(Path("POD"), "123456789012")
        self.assertEqual("123456789012.pdf", main.name)
        self.assertEqual("123456789012+.pdf", detail.name)

    def test_status_detection_is_exact_enough(self) -> None:
        self.assertTrue(DELIVERED_RE.search("DELIVERED Tuesday at 09:30"))
        self.assertTrue(DELIVERED_RE.search("货件已送达"))
        self.assertFalse(DELIVERED_RE.search("Delivery updated"))

    def test_pdf_validation(self) -> None:
        with TemporaryDirectory() as directory:
            path = Path(directory) / "test.pdf"
            path.write_bytes(b"%PDF-1.7\n" + b"x" * 1100)
            self.assertTrue(is_valid_pdf(path))
            path.write_bytes(b"not a pdf")
            self.assertFalse(is_valid_pdf(path))

    def test_snapshot_html_escapes_api_text(self) -> None:
        snapshot = ShipmentSnapshot(
            "123456789012", "Delivered", "2026-09-09", "A&B <C>", "CN", "US",
            "Priority", "1 KG", (("123456789012", "Delivered"),), ()
        )
        document = _snapshot_html(snapshot, detailed=True)
        self.assertIn("A&amp;B &lt;C&gt;", document)
        self.assertIn("物流记录", document)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="保存已送达 FedEx 运单的主查询页和详情页 PDF。"
    )
    parser.add_argument("tracking_numbers", nargs="*", help="一个或多个 FedEx 运单号")
    parser.add_argument("-o", "--output-dir", type=Path, default=Path("POD"))
    parser.add_argument("--profile-dir", type=Path, default=Path(".fedex_browser"))
    parser.add_argument("--browser", default="", help="chrome.exe 或 msedge.exe 的完整路径")
    parser.add_argument(
        "--shipmenttrack-dir",
        type=Path,
        help="可选：复用本机 ShipmentTrack 中已加密保存的 FedEx API 设置",
    )
    parser.add_argument("--headless", action="store_true", help="无界面运行（首次使用不建议）")
    parser.add_argument("--timeout", type=int, default=90, help="每个页面等待秒数")
    parser.add_argument("--overwrite", action="store_true", help="覆盖已存在的两个 PDF")
    parser.add_argument(
        "--mode",
        choices=("auto", "browser", "api"),
        default="auto",
        help="auto 有 API 环境变量时优先 API；否则使用 FedEx 网页（默认）",
    )
    parser.add_argument("--self-test", action="store_true", help="运行内置单元测试")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.self_test:
        suite = unittest.defaultTestLoader.loadTestsFromTestCase(_SelfTests)
        return 0 if unittest.TextTestRunner(verbosity=2).run(suite).wasSuccessful() else 1
    if not args.tracking_numbers:
        print("错误：请提供至少一个 FedEx 运单号。", file=sys.stderr)
        return 2
    if args.timeout < 10:
        print("错误：--timeout 不能小于 10 秒。", file=sys.stderr)
        return 2

    try:
        if args.shipmenttrack_dir:
            load_shipmenttrack_credentials(args.shipmenttrack_dir)
        has_api = bool(os.getenv("FEDEX_API_KEY", "").strip()) and bool(
            os.getenv("FEDEX_API_SECRET", "").strip()
        )
        use_api = args.mode == "api" or (args.mode == "auto" and has_api)
        if use_api:
            results = archive_with_api(
                args.tracking_numbers,
                args.output_dir,
                browser_path=args.browser,
                timeout_seconds=args.timeout,
                overwrite=args.overwrite,
            )
        else:
            downloader = FedExPODDownloader(
                args.output_dir,
                args.profile_dir,
                browser_path=args.browser,
                headless=args.headless,
                timeout_seconds=args.timeout,
                overwrite=args.overwrite,
            )
            results = downloader.run(args.tracking_numbers)
    except Exception as exc:
        print(f"启动失败：{exc}", file=sys.stderr)
        return 1

    for result in results:
        print(json.dumps(asdict(result), ensure_ascii=False))
    return 0 if results and all(result.ok for result in results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
