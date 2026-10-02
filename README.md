# 内测批量续卡 Skill

这是一个跨 Agent 的签到记录清洗与批量续卡工具。它会先列出原始 Excel 中实际出现的卡项，等待用户确认哪些卡项需要续卡以及每项增加的天数和次数，再生成最终续卡 Excel。

## 适用场景

- 门店内测、试营业或活动结束后的批量续期与加次。
- 原始表中卡项复杂，需要先盘点再决定续卡范围。
- 不同卡项需要设置不同的新增天数或次数。
- 需要补充用户卡 ID、成员姓名、手机号码和卡项类别。
- 需要对用户卡 ID 去重、检查冲突并按卡项分组。

不适用于普通签到排序、标色或与续卡无关的会员资产清洗。

## 兼容性

| 环境 | 支持情况 | 使用方式 |
| --- | --- | --- |
| Codex Desktop | 支持 | Python 便携后端；也可使用 Node 增强后端生成预览 |
| CodeBuddy | 支持 | 需要 Python 3.8+ 和本地文件读写权限 |
| Claude Code | 支持 | 需要 Python 3.8+ 和本地文件读写权限 |
| Cursor / 普通终端 Agent | 支持 | 需要 Python 3.8+ 和本地文件读写权限 |
| 无网络沙箱 | 支持 | Python 后端不下载依赖 |
| 仅有 Node、没有 Python、没有 `@oai/artifact-tool` | 不支持 | 需要提供 Python 3 或 Codex 私有依赖 |

便携后端要求 Python 3.8 或更高版本，只使用标准库：`argparse`、`json`、`zipfile` 和 `xml.etree.ElementTree` 等。无需运行 `pip install`。

## 工作流程

### 1. 用户提供原始签到记录

```text
使用 $internal-test-batch-renewal 清洗这份签到记录，并告诉我表里有哪些卡项。
```

### 2. Agent 只读盘点

Agent 返回每个卡项的出现行数、唯一用户卡数量，以及空白行、占位行、重复数据、缺失字段和冲突 ID。

### 3. 用户确认规则

```text
月卡：新增天数 1，新增次数 0
季卡：新增天数 2，新增次数 0
10次卡：新增天数 -1，新增次数 1
续卡原因：杭州来福士内测续卡
```

没有选择的卡项不会进入最终 Excel。`新增天数 = -1` 表示永久有效。

### 4. Agent 生成并校验 Excel

最终表格按用户确认的卡项顺序分组，包含：

1. `用户卡ID`
2. `成员姓名`
3. `手机号码`
4. `卡项类别`
5. `新增天数`
6. `新增次数`
7. `续卡原因`

## 直接运行：跨 Agent 便携后端

macOS/Linux 使用 `python3`；Windows 或没有 `python3` 命令时使用 `python`。

第一阶段只读盘点：

```bash
python3 scripts/build_batch_renewal.py \
  --mode inspect \
  --source "/absolute/path/签到记录.xlsx"
```

用户确认后建立规则文件：

```json
{
  "rules": [
    { "card": "月卡", "days": 1, "times": 0 },
    { "card": "10次卡", "days": -1, "times": 1 }
  ]
}
```

第二阶段生成：

```bash
python3 scripts/build_batch_renewal.py \
  --mode build \
  --source "/absolute/path/签到记录.xlsx" \
  --output "/absolute/path/内测批量续卡.xlsx" \
  --reason "10.1悠方店内测续卡" \
  --rules-json "/absolute/path/renewal-rules.json"
```

使用自定义模板时追加：

```bash
--template "/absolute/path/批量续卡模板.xlsx"
```

### 便携后端参数

| 参数 | 是否必填 | 说明 |
| --- | --- | --- |
| `--mode` | 否 | `inspect` 盘点；`build` 生成，默认为 `build` |
| `--source` | 是 | 原始签到记录 `.xlsx` |
| `--output` | build 时必填 | 最终输出 `.xlsx` |
| `--reason` | build 时必填 | 续卡原因 |
| `--rules-json` | 推荐 | 用户确认后的逐卡规则 |
| `--template` | 否 | 自定义模板；默认使用仓库内模板 |
| `--duration-cards` | 兼容模式 | 旧版期限卡列表，固定使用 `1 / 0` |
| `--count-cards` | 兼容模式 | 旧版次卡列表，固定使用 `-1 / 1` |

## Codex 增强后端

仓库仍保留 `scripts/build_batch_renewal.mjs`。它使用 Codex 的 `@oai/artifact-tool`，可以生成 PNG 预览，但不是跨 Agent 运行的必要条件。

只有确认当前环境提供该依赖时才使用：

```bash
"<node_path>" scripts/build_batch_renewal.mjs \
  --mode inspect \
  --node-modules "<node_modules_path>" \
  --source "/absolute/path/签到记录.xlsx"
```

普通 Agent 应直接使用 Python 便携后端，不要尝试从公共 npm 安装 `@oai/artifact-tool`。

## 校验与安全规则

- 卡项名称使用完全匹配，不使用模糊匹配。
- 空白卡项及 `-`、`—`、`－` 占位卡项不会进入候选列表。
- 目标记录缺少用户卡 ID、姓名或手机号时停止。
- 完全重复的用户卡记录只保留第一条。
- 同一用户卡 ID 信息不一致时停止并报告。
- 用户确认的某个卡项在源表中不存在时停止。
- 输出后重新打开文件，验证表头、记录数、逐卡规则、重复 ID 和续卡原因。
- 模板主工作表之外的所有文件部件保持不变。

## 文件结构

```text
internal-test-batch-renewal/
├── SKILL.md
├── README.md
├── agents/
│   └── openai.yaml
├── assets/
│   └── batch-renewal-template.xlsx
└── scripts/
    ├── build_batch_renewal.py    # 跨 Agent 便携后端
    └── build_batch_renewal.mjs   # Codex 增强后端
```

## 处理失败时

脚本会返回非零退出码并输出中文错误，不会生成可能错误的续卡文件。常见原因包括字段缺失、规则无效、卡项未匹配、用户卡 ID 冲突或模板损坏。
