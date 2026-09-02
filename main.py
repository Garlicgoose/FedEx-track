# -*- coding: utf-8 -*-
"""FedExMailAlert —— FedEx 邮件跟踪 demo（2026-09）。

作用：批量监控 FedEx 运单，未送达且超过 N 天（或处于清关/海关状态）时
自动发邮件提醒（最多 6 个收件邮箱），避免运单卡住无人跟进。

- 查询：官方 FedEx Tracking API（modules/fedex_module.py，与 ShipmentTrack 同逻辑）
- 发送：Outlook COM（需本机桌面版 Outlook；也可换 SMTP）
- 同一运单同一触发条件只提醒一次（避免重复轰炸）
"""
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

BASE = Path(__file__).resolve().parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                               QLabel, QLineEdit, QPlainTextEdit, QPushButton,
                               QSpinBox, QCheckBox, QMessageBox, QGroupBox,
                               QFormLayout, QFileDialog)

import json

sys.path.insert(0, str(BASE / "modules"))
from fedex_module import query_fedex_one  # noqa: E402

APP_TITLE = "FedEx Mail Alert"
SETTINGS_FILE = BASE / "settings.json"

# 触发后标记：记录 (运单号, 原因) -> 已提醒，避免重复
TRIGGER_STATES = {
    "clearance": ["Clearance", "Customs", "International shipment release",
                  "In transit to destination country"],
}


def load_settings():
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(s):
    SETTINGS_FILE.write_text(json.dumps(s, ensure_ascii=False, indent=1),
                             encoding="utf-8")


# --------------------------------------------------------------- mail send

def send_alert_email(subject, body, to_emails):
    """Outlook COM 发送。失败抛异常（由调用方提示）。"""
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.Body = body
    mail.To = "; ".join(e for e in to_emails if e)
    mail.Send()
    return True


# ------------------------------------------------------------ monitor worker

class MonitorWorker(QThread):
    log = Signal(str)
    alert = Signal(str, str)      # (subject, body)
    finished_ok = Signal(bool, str)

    def __init__(self, numbers, to_emails, api_key, api_secret,
                 stale_days, alert_clearance, interval_min, parent=None):
        super().__init__(parent)
        self.numbers = [n.strip() for n in numbers if n.strip()]
        self.to_emails = to_emails
        self.api_key = api_key
        self.api_secret = api_secret
        self.stale_days = stale_days
        self.alert_clearance = alert_clearance
        self.interval_min = max(1, interval_min)
        self._stop = False
        self._notified = {}   # (number, reason) -> date str

    def stop(self):
        self._stop = True

    def _fire(self, number, reason, body):
        subject = "[FedEx Alert] {} {}".format(number, reason)
        self.alert.emit(subject, body + "\n\n— FedEx Mail Alert")

    def run(self):
        # 简化天数逻辑：用首次监控日期，之后每次检查对比
        first_seen = {n: datetime.now() for n in self.numbers}
        while not self._stop:
            self.log.emit("== 检查 {} 个运单 @ {} ==".format(
                len(self.numbers), datetime.now().strftime("%H:%M:%S")))
            for num in self.numbers:
                if self._stop:
                    return
                try:
                    r = query_fedex_one(num, self.api_key, self.api_secret,
                                        save_pdf=False)
                except Exception as exc:  # noqa: BLE001
                    self.log.emit("ERR {}: {}".format(num, exc))
                    continue
                if r.get("error"):
                    self.log.emit("{}: {}".format(num, r["error"]))
                    continue
                st = r.get("status") or "Unknown"
                delivered = r["is_delivered"] == "Y"
                self.log.emit("{}: {}".format(
                    num, "Delivered" if delivered else st))
                if delivered:
                    continue
                # 清关/海关
                if self.alert_clearance and any(
                        k in st.lower() for k in ("clearance", "customs")):
                    key = (num, "clearance")
                    if key not in self._notified:
                        self._notified[key] = True
                        self._fire(num, "清关/海关",
                                   "运单 {} 处于 {}，可能卡在清关，请跟进。".format(num, st))
                        continue
                # 超期未送达（按首次发现计时）
                days = (datetime.now() - first_seen[num]).days
                if days >= self.stale_days:
                    key = (num, "stale")
                    if key not in self._notified:
                        self._notified[key] = True
                        self._fire(num, "超过 {} 天未送达".format(self.stale_days),
                                   "运单 {} 已监控 {} 天仍未送达（当前状态：{}），请跟进。".format(
                                       num, days, st))
            # 等下一轮
            for _ in range(self.interval_min * 60):
                if self._stop:
                    return
                time.sleep(1)
        self.finished_ok.emit(True, "")


# ------------------------------------------------------------------ UI

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(720, 640)
        self.settings = load_settings()
        self.worker = None
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)

        # ---- FedEx API ----
        api = QGroupBox("FedEx API")
        af = QFormLayout(api)
        self.key_edit = QLineEdit(self.settings.get("fedex_key", ""))
        self.secret_edit = QLineEdit(self.settings.get("fedex_secret", ""))
        self.secret_edit.setEchoMode(QLineEdit.EchoMode.Password)
        af.addRow("API Key", self.key_edit)
        af.addRow("API Secret", self.secret_edit)
        lay.addWidget(api)

        # ---- 监控设置 ----
        mon = QGroupBox("监控设置")
        mf = QFormLayout(mon)
        self.numbers_edit = QPlainTextEdit()
        self.numbers_edit.setPlaceholderText(
            "每行一个 FedEx 运单号\n（批量跟踪：子单会一并判断送达）")
        self.numbers_edit.setFixedHeight(90)
        mf.addRow("运单号", self.numbers_edit)
        self.emails_edit = QLineEdit(self.settings.get("emails", ""))
        self.emails_edit.setPlaceholderText("最多 6 个，逗号分隔")
        mf.addRow("收件邮箱", self.emails_edit)
        self.stale_spin = QSpinBox()
        self.stale_spin.setRange(1, 365)
        self.stale_spin.setValue(int(self.settings.get("stale_days", 7)))
        mf.addRow("超过 N 天未送达提醒", self.stale_spin)
        self.clearance_check = QCheckBox("清关/海关（Customs/Clearance）状态立即提醒")
        self.clearance_check.setChecked(bool(self.settings.get("clearance", True)))
        mf.addRow("", self.clearance_check)
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(5, 1440)
        self.interval_spin.setValue(int(self.settings.get("interval_min", 30)))
        mf.addRow("检查间隔（分钟）", self.interval_spin)
        lay.addWidget(mon)

        # ---- 操作 ----
        ops = QHBoxLayout()
        self.start_btn = QPushButton("开始监控")
        self.start_btn.clicked.connect(self._start)
        self.stop_btn = QPushButton("停止")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._stop)
        self.test_btn = QPushButton("发送测试邮件")
        self.test_btn.clicked.connect(self._test_mail)
        ops.addWidget(self.start_btn)
        ops.addWidget(self.stop_btn)
        ops.addWidget(self.test_btn)
        ops.addStretch(1)
        lay.addLayout(ops)

        # ---- 日志 ----
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        lay.addWidget(self.log_view, 1)

        self.log("FedEx Mail Alert 就绪。填好 API Key/Secret + 运单号 + 邮箱后点开始。")

    # ------------------------------------------------------------ helpers

    def log(self, msg):
        self.log_view.appendPlainText(
            "[{}] {}".format(datetime.now().strftime("%H:%M:%S"), msg))

    def _collect(self):
        key = self.key_edit.text().strip()
        secret = self.secret_edit.text().strip()
        nums = [n.strip() for n in self.numbers_edit.toPlainText().splitlines()
                if n.strip()]
        emails = [e.strip() for e in
                  self.emails_edit.text().replace("；", ";").replace(";", ",")
                  .replace("，", ",").split(",") if e.strip()][:6]
        if not key or not secret:
            QMessageBox.warning(self, APP_TITLE, "请填写 FedEx API Key / Secret")
            return None
        if not nums:
            QMessageBox.warning(self, APP_TITLE, "请至少输入一个运单号")
            return None
        if not emails:
            QMessageBox.warning(self, APP_TITLE, "请至少填一个收件邮箱")
            return None
        self.settings.update({
            "fedex_key": key, "fedex_secret": secret,
            "emails": self.emails_edit.text().strip(),
            "stale_days": self.stale_spin.value(),
            "clearance": self.clearance_check.isChecked(),
            "interval_min": self.interval_spin.value(),
        })
        save_settings(self.settings)
        return key, secret, nums, emails

    def _start(self):
        got = self._collect()
        if got is None:
            return
        key, secret, nums, emails = got
        self.worker = MonitorWorker(
            nums, emails, key, secret,
            self.stale_spin.value(), self.clearance_check.isChecked(),
            self.interval_spin.value())
        self.worker.log.connect(self.log)
        self.worker.alert.connect(self._on_alert)
        self.worker.start()
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.log("监控开始：{} 个运单，间隔 {} 分钟，提醒到 {} 个邮箱".format(
            len(nums), self.interval_spin.value(), len(emails)))

    def _stop(self):
        if self.worker is not None and self.worker.isRunning():
            self.worker.stop()
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.log("已请求停止（当前轮结束后生效）")

    def _on_alert(self, subject, body):
        self.log("触发邮件: {}".format(subject))
        emails = [e.strip() for e in
                  self.settings.get("emails", "").replace("；", ";")
                  .replace(";", ",").replace("，", ",").split(",") if e.strip()][:6]
        try:
            send_alert_email(subject, body, emails)
            self.log("邮件已发送到: {}".format(", ".join(emails)))
        except Exception as exc:  # noqa: BLE001
            self.log("发送失败（需本机装 Outlook）: {}".format(exc))
            QMessageBox.critical(
                self, APP_TITLE,
                "邮件发送失败：{}\n\n需要本机安装并登录桌面版 Outlook。".format(exc))

    def _test_mail(self):
        got = self._collect()
        if got is None:
            return
        _k, _s, _n, emails = got
        try:
            send_alert_email(
                "[FedEx Alert] 测试邮件",
                "这是一封测试邮件。FedEx Mail Alert 配置正常。", emails)
            self.log("测试邮件已发送到: {}".format(", ".join(emails)))
            QMessageBox.information(self, APP_TITLE, "测试邮件已发送")
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(
                self, APP_TITLE,
                "发送失败：{}\n\n需要本机安装并登录桌面版 Outlook。".format(exc))

    def closeEvent(self, event):
        if self.worker is not None and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait(3000)
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
