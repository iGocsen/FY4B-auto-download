# 定时任务配置参考

## 目标 cron 任务

- **ID**：`cfceaff2-848d-4de5-8188-8e2b7b4c0e13`
- **名称**：FY4B云图每日下载
- **schedule**：`30 9,15,21 * * *`（每天 09:30 / 15:30 / 21:30 北京时间）
- **sessionTarget**：`isolated`
- **delivery.mode**：`none`

## payload.message 模板

```text
步骤1：执行 FY4B 云图下载脚本获取输出。

运行以下命令并捕获完整输出：
python %UserProfile%\.qclaw\skills\FY4B-auto-download\scripts\fy4b_download.py

步骤2：从输出中提取 Markdown 报告。

如果输出包含「📡 FY4B 云图下载报告」则为成功，直接使用该 Markdown 内容。
如果输出包含「⚠️ 解析失败」则为失败，使用错误原文作为消息。

步骤3：读取配置文件 %UserProfile%\.qclaw\skills\FY4B-auto-download\scripts\skill-config.json，遍历 push_targets 数组，对每个目标调用 message 工具（channel/to/accountId 均来自配置字段）。

注意：不要回复 HEARTBEAT_OK，不要有多余文字，直接调用 message 工具推送报告内容即可并在对话窗口输出捕获的 markdown 内容
```

## 更新推送目标步骤

1. 修改 `scripts/skill-config.json` 中的 `push_targets` 数组
2. 执行 `cron update` 同步修改任务 payload（推送目标从配置读取，payload 无需变更）
2. payload 指示 agent 从配置文件读取推送目标，因此**只需改配置文件即可，无需更新 cron**

## 查看当前任务

```python
cron(action="list")
```

## 手动触发测试

```python
cron(action="run", jobId="cfceaff2-848d-4de5-8188-8e2b7b4c0e13")
```