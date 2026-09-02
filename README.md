# FedExMailAlert（FedEx 运单邮件提醒 demo）

批量监控 FedEx 运单：**清关/海关状态**或**超过 N 天未送达**时自动发邮件提醒
（最多 6 个收件邮箱，同一运单同一原因只提醒一次）。

## 功能
- 查询：官方 FedEx Tracking API（复用 ShipmentTrack fedex_module 逻辑：
  无子单运单以官网状态为准；有子单全 DL 才算送达；=40 全 DL 默认送达+备注）
- 发送：Outlook COM（需本机桌面版 Outlook）
- API Key/Secret 界面填写，存 settings.json（不入库）
- 检查间隔可调（默认 30 分钟）

## 运行（源码）
```
pip install PySide6 requests pywin32
python main.py
```

## 打包
```
powershell -ExecutionPolicy Bypass -File build.ps1
# 产物: dist\FedExMailAlert\FedExMailAlert.exe
```

## 目录
```
main.py / modules/fedex_module.py（FedEx API 查询，与 ShipmentTrack 同步）
settings.json（FedEx API 凭据+邮箱，运行时生成，不入库）
```

## 说明
- FedEx API key：developer.fedex.com 免费注册（与 ShipmentTrack/CargoMate 可共用）
- settings.json 不入库（含 API 密钥）
