---
name: internal-test-batch-renewal
description: 将香蕉攀岩门店的原始签到记录 Excel 清洗并列出实际卡项，向用户确认逐卡新增天数和次数后生成批量续卡模板。适用于 Codex、CodeBuddy、Claude Code、Cursor 等能运行 Python 3 的 Agent 环境；不用于普通签到排序或其他会员资产清洗。
---

# 内测批量续卡

采用“盘点卡项 → 用户确认 → 生成 Excel”的两阶段流程。默认运行无需第三方包的 Python 便携后端，不要假设环境中存在 Codex 私有依赖。

## 环境选择

优先使用 [scripts/build_batch_renewal.py](scripts/build_batch_renewal.py)：

- 仅要求 Python 3.8 或更高版本。
- 只使用 Python 标准库，不需要联网、`pip install`、Node.js 或 `@oai/artifact-tool`。
- macOS/Linux 优先尝试 `python3`，Windows 或没有 `python3` 时尝试 `python`。

Codex 环境可以选用 [scripts/build_batch_renewal.mjs](scripts/build_batch_renewal.mjs) 获得 PNG 预览，但它需要 Codex 工作区提供的 `@oai/artifact-tool`。不要在 CodeBuddy 或普通沙箱中选择 Node 后端，除非已经确认该依赖存在。

如果 Python 和 Codex 增强后端都不可用，停止并报告缺失的运行环境，不要静默改用手工复制。

## 第一阶段：清洗并盘点

运行只读命令：

```bash
python3 scripts/build_batch_renewal.py \
  --mode inspect \
  --source "/absolute/path/签到记录.xlsx"
```

向用户展示：

- 每个真实卡项的完整名称
- 出现行数
- 唯一用户卡 ID 数
- 空白行、`-` 等占位行、完全重复行、缺失字段行和冲突 ID 数

占位行不作为真实卡项。只列出源表实际出现的卡项，卡名必须完整展示。

如果用户首次请求没有给出完整规则，暂停生成并询问：

```text
请确认需要续卡的卡项：
月卡：新增天数 1，新增次数 0
10次卡：新增天数 -1，新增次数 1
续卡原因：10.1悠方店内测续卡
```

每个卡项都要确认 `新增天数` 和 `新增次数`。如果用户只说“次卡加次”但没有说明是否永久有效，必须追问；用户要求永久有效时使用 `新增天数 = -1`。

## 第二阶段：按确认规则生成

把用户确认的规则写入临时 JSON：

```json
{
  "rules": [
    { "card": "月卡", "days": 1, "times": 0 },
    { "card": "10次卡", "days": -1, "times": 1 }
  ]
}
```

规则要求：

- `card` 必须与盘点结果中的完整名称精确一致。
- `days` 是大于等于 `-1` 的整数。
- `times` 是大于等于 `0` 的整数。
- 两者不能同时为 `0`。

生成：

```bash
python3 scripts/build_batch_renewal.py \
  --mode build \
  --source "/absolute/path/签到记录.xlsx" \
  --output "/absolute/path/内测批量续卡.xlsx" \
  --reason "10.1悠方店内测续卡" \
  --rules-json "/absolute/path/renewal-rules.json"
```

脚本默认使用 [assets/batch-renewal-template.xlsx](assets/batch-renewal-template.xlsx)。只有用户明确提供其他模板时才添加 `--template`。

若所在 Agent 对文件创建有自己的标记、审批或交付规范，遵守宿主环境规范；这些规范不是便携脚本自身的运行依赖。Codex 中生成 Excel 时仍应遵守可用的 Spreadsheets skill。

## 输出和排序

主表固定包含：

1. `用户卡ID`
2. `成员姓名`
3. `手机号码`
4. `卡项类别`
5. `新增天数`
6. `新增次数`
7. `续卡原因`

用户卡 ID、姓名、手机号和卡项按文本写入；天数和次数按整数写入。忽略用户未选择的卡项。按规则 JSON 的卡项顺序分组，同一卡项内保留源表顺序。

## 数据完整性

- 目标记录缺少用户卡 ID、姓名或手机号：停止并报告源行。
- 同一用户卡 ID 完全重复：只保留第一条并报告跳过数量。
- 同一用户卡 ID 对应不同卡项、姓名或手机号：停止并报告冲突。
- 用户选择了源表不存在的卡项：停止并报告。
- 输出后重新打开文件，校验表头、记录数、逐卡规则、重复 ID、续卡原因和公式错误。
- 模板主工作表以外的 ZIP 部件必须保持逐字节不变。

## 旧版规则兼容

用户明确要求直接使用旧规则时，可以不传 `--rules-json`：

- `新人月卡、月卡、季卡、年卡`：新增天数 `1`，新增次数 `0`
- `10次卡、20次卡`：新增天数 `-1`，新增次数 `1`

可以通过 `--duration-cards` 和 `--count-cards` 覆盖旧版卡项。交互式流程仍优先使用 `--rules-json`。

## Codex 增强后端

只有确认存在 Codex 工作区 Node.js 与 `@oai/artifact-tool` 时才使用：

```bash
"<node_path>" scripts/build_batch_renewal.mjs \
  --mode inspect \
  --node-modules "<node_modules_path>" \
  --source "/absolute/path/签到记录.xlsx"
```

生成阶段参数与 Python 后端相同，另外需要 `--node-modules`，可选 `--preview-dir`。Node 后端不是跨 Agent 的必需依赖。

## 交付

最终只交付生成的 `.xlsx`，并报告每个卡项的输出数量、去重数量和被阻止的异常。不要交付临时规则 JSON、预览或审计文件。
