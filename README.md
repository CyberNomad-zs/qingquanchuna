# 清泉出纳 · 考勤整理

钉钉考勤报表 → 整理结果 Excel → 飞书多维表格。

## 用法

双击 `考勤整理启动器.bat`，或看 `使用说明.html`。

```text
python kaoqin.py 7月考勤报表_20260701_20260731.xlsx
python feishu-attendance-monthly.py 7月考勤报表_xxx.xlsx 7月考勤整理结果_校验.xlsx          # 干跑
python feishu-attendance-monthly.py 7月考勤报表_xxx.xlsx 7月考勤整理结果_校验.xlsx --apply  # 上传
```

## 飞书凭据

复制 `.env.example` 为 `.env`，填入飞书开放平台的 App ID / Secret。不要把 `.env` 提交进仓库。
