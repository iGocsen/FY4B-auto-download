# 定时任务配置参考

## 目标 cron 任务

- **ID**：`cfceaff2-848d-4de5-8188-8e2b7b4c0e13`
- **名称**：FY4B云图每日下载
- **schedule**：`30 9,15,21 * * *`（每天 09:30 / 15:30 / 21:30 北京时间）
- **sessionTarget**：`isolated`
- **delivery.mode**：`none`

## payload.message 模板

```text
执行 Python 脚本：`python %UserProfile%\.qclaw\skills\FY4B-auto-download\scripts\fy4b_download.py`。

脚本执行后会输出 Markdown 格式报告，请等待完成后，直接调用 message 工具向配置中所有 push_targets 推送相同内容。

读取配置文件 `%UserProfile%\.qclaw\skills\FY4B-auto-download\scripts\config.json` 获取 push_targets 列表，遍历每个目标调用 message（channel/to/accountId 均来自配置）。

不要回复 HEARTBEAT_OK，不要有多余文字，直接调用 message 工具即可。
```

## 更新推送目标步骤

1. 修改 `config.json` 中的 `push_targets` 数组
2. 执行 `cron update` 同步修改任务 payload（推送目标从配置读取，payload 无需变更）
3. 实际上 payload 指示 agent 读取配置文件，因此只需改配置文件即可

## 查看当前任务

```python
cron(action="list")
```

## 手动触发测试

```python
cron(action="run", jobId="cfceaff2-848d-4de5-8188-8e2b7b4c0e13")
```