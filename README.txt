FedExMailAlert（2026-09-01 demo）
================================

作用：批量监控 FedEx 运单，以下情况自动发邮件提醒（最多 6 个邮箱）：
  1) 运单处于 清关/海关（Clearance/Customs）状态
  2) 未送达且监控超过 N 天（默认 7 天，可调）
同一运单同一原因只提醒一次（避免重复轰炸）。

运行（源码）：
  pip install PySide6 requests pywin32
  python main.py

打包：
  powershell -ExecutionPolicy Bypass -File build.ps1

配置说明：
  - FedEx API Key/Secret：developer.fedex.com 注册（与 ShipmentTrack 相同凭据可复用）
  - 运单号：每行一个；批量跟踪会走子单全送达判断（fedex_module 与 ShipmentTrack 同逻辑）
  - 收件邮箱：逗号分隔，最多 6 个
  - 发送：走本机桌面 Outlook（Outlook COM）。本机没有 Outlook 会发送失败并提示，
    可在公司装有 Outlook 的电脑上运行。

查询逻辑（与 ShipmentTrack v0.6 一致）：
  1. trackingnumbers 正常查主单；无子单运单以官网状态为准
  2. 有子单：<40 全查全送达才算送达；=40 全送达默认送达（备注人工核查）
