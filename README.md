# 清泉出纳 · 考勤整理

钉钉考勤报表 → 整理结果 Excel → 飞书多维表格。

整理结果 Sheet1 为报表全员（含「纳入汇总」标记）。上传飞书时：共用表「考勤报表 / 打卡时间明细 / 考勤结果总表」每次清空重写；当月「{月}打卡异常说明 / {月}其他员工」按前缀建或复用，历史月表不动。

## 用法

双击 `考勤整理启动器.bat`，或看 `使用说明.html`。

```text
python kaoqin.py 7月考勤报表_20260701_20260731.xlsx
python feishu-attendance-monthly.py 7月考勤报表_xxx.xlsx 7月考勤整理结果_校验.xlsx          # 干跑
python feishu-attendance-monthly.py 7月考勤报表_xxx.xlsx 7月考勤整理结果_校验.xlsx --apply  # 上传
```

## 飞书凭据

复制 `.env.example` 为 `.env`，填入飞书开放平台的 App ID / Secret。不要把 `.env` 提交进仓库。
