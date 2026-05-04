---
name: FY4B-auto-download
description: FY4B 卫星云图自动下载技能。根据定时触发，自动检查待导入批次、通过 Python 脚本下载 FY4B 云图、更新 Excel 过期时间并生成新的下载链接列表。触发词包括：FY4B云图、定时下载卫星云图、FY4B自动下载、卫星云图监控。当用户提到 FY4B 下载任务、定时任务配置、云图自动导入时使用此技能。
---

# FY4B 云图自动下载

## 配置读取

所有关键配置从 `scripts/config.json` 读取（相对于 skill 目录）。**不要硬编码任何路径或账号信息**，一律从配置文件读取。

> ⚠️ 配置文件路径：`~\.qclaw\skills\FY4B-auto-download\scripts\config.json`

读取配置示例：
```python
import json, pathlib
_config_path = pathlib.Path(__file__).parent / "config.json"
_config = json.loads(_config_path.read_text(encoding="utf-8"))
```

### 配置字段说明

| 字段 | 说明 |
|------|------|
| `excel_path` | Excel 文件路径（含链接过期时间） |
| `excel_sheet` | 工作表名称 |
| `url_template_file` | 下载链接模板文件路径 |
| `staging_dir` | 临时下载目录（下载完成后移动至此） |
| `target_dir` | 最终图片保存目录 |
| `push_targets` | 推送目标列表，含 `channel`、`to`、`accountId` |
| `cron_schedule` | cron 表达式，固定 `30 9,15,21 * * *` |

## 工作流程

### 定时触发（agentTurn）

由 OpenClaw cron 定时触发，payload 指示 agent 执行以下步骤：

1. **执行下载脚本**：`python {skill_dir}/scripts/fy4b_download.py`
2. **解析输出**：从脚本日志中提取 `[DONE] 新下载: X | 已存在: Y | 未来: Z | 失败: W`
3. **推送结果**：读取 `push_targets`，向每个目标发送 Markdown 格式报告

### 手动触发

用户说"测试 FY4B 下载"时，直接执行 `python {skill_dir}/scripts/fy4b_download.py`，解析日志，推送结果。

### 定时任务配置

首次配置时，创建 cron 任务：

- schedule: `30 9,15,21 * * *`（每天 09:30 / 15:30 / 21:30，北京时间）
- sessionTarget: `isolated`
- payload.kind: `agentTurn`
- delivery.mode: `none`（推送由 agentTurn 内的 message 工具完成）

参考 `references/cron-setup.md` 获取详细步骤。

## 推送消息格式

```markdown
📡 **FY4B 云图下载报告**

✅ 新下载: **X**
⏭️ 已存在: **Y**
🔮 未来图片: **Z**
❌ 失败: **W**
```

## 关键脚本

- `scripts/fy4b_download.py` — 主下载脚本，读取配置、执行下载、更新 Excel、生成 txt
- `scripts/config.json` — 统一配置文件，所有路径和账号信息
- `scripts/create_download_links.py` — 生成链接 txt 工具（被主脚本调用）
- `scripts/check_excel.py` — Excel/txt 状态检查模块（被主脚本调用）

## Excel 结构参考

- B1：当前批次过期时间（北京时间）
- B162：下一批次过期时间
- 更新逻辑：B162 值 → 覆盖 B2 → Excel 内部公式链自动重算（B2→B3→...→B162）

详见 `references/excel-structure.md`。

## 更新定时任务

当用户修改推送目标（如更换微信 ID）时：
1. 读取当前 `config.json` 中的 `push_targets`
2. 用 `cron update` 命令更新任务 payload 中的 message 内容
3. 保持 schedule 不变