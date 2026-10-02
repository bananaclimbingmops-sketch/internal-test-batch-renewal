#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const defaultTemplatePath = path.resolve(scriptDir, "..", "assets", "batch-renewal-template.xlsx");
const defaultDurationCards = ["新人月卡", "月卡", "季卡", "年卡"];
const defaultCountCards = ["10次卡", "20次卡"];
const expectedHeaders = ["用户卡ID", "成员姓名", "手机号码", "卡项类别", "新增天数", "新增次数", "续卡原因"];

function parseArgs(argv) {
  const parsed = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) {
      throw new Error(`无法识别的参数：${token}`);
    }
    const key = token.slice(2);
    const value = argv[index + 1];
    if (value === undefined || value.startsWith("--")) {
      throw new Error(`参数 --${key} 缺少值`);
    }
    parsed[key] = value;
    index += 1;
  }
  return parsed;
}

function required(args, key) {
  const value = String(args[key] ?? "").trim();
  if (!value) throw new Error(`缺少必填参数 --${key}`);
  return value;
}

function parseCardList(raw, fallback) {
  if (raw === undefined) return [...fallback];
  const cards = String(raw)
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);
  if (!cards.length) throw new Error("卡项列表不能为空");
  if (new Set(cards).size !== cards.length) throw new Error(`卡项列表包含重复值：${cards.join(", ")}`);
  return cards;
}

function validateRule(rawRule, index) {
  if (!rawRule || typeof rawRule !== "object" || Array.isArray(rawRule)) {
    throw new Error(`规则第 ${index + 1} 项必须是对象`);
  }
  const card = normalize(rawRule.card);
  const days = Number(rawRule.days);
  const times = Number(rawRule.times);
  if (!card) throw new Error(`规则第 ${index + 1} 项缺少 card`);
  if (!Number.isInteger(days) || days < -1) {
    throw new Error(`卡项“${card}”的 days 必须是大于等于 -1 的整数`);
  }
  if (!Number.isInteger(times) || times < 0) {
    throw new Error(`卡项“${card}”的 times 必须是大于等于 0 的整数`);
  }
  if (days === 0 && times === 0) {
    throw new Error(`卡项“${card}”的 days 和 times 不能同时为 0`);
  }
  return { card, days, times };
}

async function loadRules(args) {
  if (args["rules-json"]) {
    if (args["duration-cards"] !== undefined || args["count-cards"] !== undefined) {
      throw new Error("--rules-json 不能与 --duration-cards 或 --count-cards 同时使用");
    }
    const rulesPath = path.resolve(args["rules-json"]);
    const parsed = JSON.parse(await fs.readFile(rulesPath, "utf8"));
    const rawRules = Array.isArray(parsed) ? parsed : parsed.rules;
    if (!Array.isArray(rawRules) || !rawRules.length) {
      throw new Error("规则文件必须是非空数组，或包含非空 rules 数组");
    }
    const rules = rawRules.map(validateRule);
    if (new Set(rules.map((rule) => rule.card)).size !== rules.length) {
      throw new Error("规则文件包含重复卡项");
    }
    return { rules, rulesPath, mode: "custom" };
  }

  const durationCards = parseCardList(args["duration-cards"], defaultDurationCards);
  const countCards = parseCardList(args["count-cards"], defaultCountCards);
  const overlap = durationCards.filter((card) => countCards.includes(card));
  if (overlap.length) throw new Error(`期限卡与次卡列表重叠：${overlap.join(", ")}`);
  return {
    rules: [
      ...durationCards.map((card) => ({ card, days: 1, times: 0 })),
      ...countCards.map((card) => ({ card, days: -1, times: 1 })),
    ],
    rulesPath: null,
    mode: "legacy",
  };
}

function columnName(index) {
  let value = index + 1;
  let result = "";
  while (value > 0) {
    const digit = (value - 1) % 26;
    result = String.fromCharCode(65 + digit) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function normalize(value) {
  return String(value ?? "").trim();
}

function isPlaceholder(value) {
  return ["-", "—", "－"].includes(normalize(value));
}

function findHeaderIndex(header, aliases) {
  for (const alias of aliases) {
    const index = header.indexOf(alias);
    if (index >= 0) return index;
  }
  return -1;
}

function sameEntry(left, right) {
  return left.card === right.card
    && left.name === right.name
    && left.phone === right.phone;
}

function findSourceData(workbook) {
  for (const sheet of workbook.worksheets.items) {
    const used = sheet.getUsedRange();
    if (!used) continue;
    const values = used.values;
    if (!values.length) continue;
    const header = values[0].map(normalize);
    const indexes = {
      id: findHeaderIndex(header, ["用户卡ID"]),
      name: findHeaderIndex(header, ["成员姓名"]),
      phone: findHeaderIndex(header, ["手机号码"]),
      card: findHeaderIndex(header, ["签到卡", "签到卡类"]),
    };
    if (Object.values(indexes).every((index) => index >= 0)) {
      return { sheet, values, header, indexes };
    }
  }
  throw new Error("没有找到同时包含“用户卡ID、成员姓名、手机号码、签到卡/签到卡类”的工作表");
}

function inspectSourceData(sourceData) {
  const cardStats = new Map();
  const seen = new Map();
  const missingFields = [];
  const conflicts = [];
  let blankRowsSkipped = 0;
  let placeholderRowsSkipped = 0;
  let duplicatesSkipped = 0;

  for (const [offset, row] of sourceData.values.slice(1).entries()) {
    const sourceRow = offset + 2;
    const entry = {
      sourceRow,
      sourceOrder: offset,
      id: normalize(row[sourceData.indexes.id]),
      name: normalize(row[sourceData.indexes.name]),
      phone: normalize(row[sourceData.indexes.phone]),
      card: normalize(row[sourceData.indexes.card]),
    };
    if (!entry.id && !entry.name && !entry.phone && !entry.card) {
      blankRowsSkipped += 1;
      continue;
    }
    if (isPlaceholder(entry.card)) {
      placeholderRowsSkipped += 1;
      continue;
    }

    const missing = [
      !entry.id || isPlaceholder(entry.id) ? "用户卡ID" : null,
      !entry.name || isPlaceholder(entry.name) ? "成员姓名" : null,
      !entry.phone || isPlaceholder(entry.phone) ? "手机号码" : null,
      !entry.card ? "签到卡/签到卡类" : null,
    ].filter(Boolean);
    if (missing.length) {
      missingFields.push({ row: sourceRow, missing });
    }
    if (!entry.card) continue;
    if (!cardStats.has(entry.card)) {
      cardStats.set(entry.card, {
        card: entry.card,
        firstSourceRow: sourceRow,
        occurrences: 0,
        uniqueIds: new Set(),
      });
    }
    const stats = cardStats.get(entry.card);
    stats.occurrences += 1;
    if (entry.id && !isPlaceholder(entry.id)) stats.uniqueIds.add(entry.id);

    if (!entry.id || isPlaceholder(entry.id)) continue;
    if (seen.has(entry.id)) {
      const prior = seen.get(entry.id);
      if (sameEntry(prior, entry)) {
        duplicatesSkipped += 1;
      } else {
        conflicts.push({ id: entry.id, first: prior, duplicate: entry });
      }
      continue;
    }
    seen.set(entry.id, entry);
  }

  return {
    mode: "inspect",
    sourceSheet: sourceData.sheet.name,
    sourceRows: Math.max(0, sourceData.values.length - 1),
    nonblankUniqueIds: seen.size,
    blankRowsSkipped,
    placeholderRowsSkipped,
    duplicatesSkipped,
    cardItems: [...cardStats.values()].map((item) => ({
      card: item.card,
      occurrences: item.occurrences,
      uniqueUserCardIds: item.uniqueIds.size,
      firstSourceRow: item.firstSourceRow,
    })),
    issues: {
      missingFieldRows: missingFields.length,
      missingFields: missingFields.slice(0, 20),
      conflictingIds: conflicts.length,
      conflicts: conflicts.slice(0, 20).map((item) => ({
        id: item.id,
        firstRow: item.first.sourceRow,
        duplicateRow: item.duplicate.sourceRow,
        firstCard: item.first.card,
        duplicateCard: item.duplicate.card,
        firstName: item.first.name,
        duplicateName: item.duplicate.name,
        firstPhone: item.first.phone,
        duplicatePhone: item.duplicate.phone,
      })),
    },
  };
}

function collectRows(sourceData, rules) {
  const rulesByCard = new Map(rules.map((rule) => [rule.card, rule]));
  const categoryOrder = new Map(rules.map((rule, index) => [rule.card, index]));
  const seen = new Map();
  const selected = [];
  const conflicts = [];
  let duplicatesSkipped = 0;

  for (const [offset, row] of sourceData.values.slice(1).entries()) {
    const sourceRow = offset + 2;
    const card = normalize(row[sourceData.indexes.card]);
    const rule = rulesByCard.get(card);
    if (!rule) continue;

    const entry = {
      sourceRow,
      sourceOrder: offset,
      id: normalize(row[sourceData.indexes.id]),
      name: normalize(row[sourceData.indexes.name]),
      phone: normalize(row[sourceData.indexes.phone]),
      card,
      rule,
    };

    if (!entry.id) throw new Error(`源表“${sourceData.sheet.name}”第 ${sourceRow} 行缺少用户卡ID`);
    if (!entry.name) throw new Error(`源表“${sourceData.sheet.name}”第 ${sourceRow} 行缺少成员姓名`);
    if (!entry.phone) throw new Error(`源表“${sourceData.sheet.name}”第 ${sourceRow} 行缺少手机号码`);

    if (seen.has(entry.id)) {
      const prior = seen.get(entry.id);
      if (sameEntry(prior, entry)) {
        duplicatesSkipped += 1;
      } else {
        conflicts.push({ id: entry.id, first: prior, duplicate: entry });
      }
      continue;
    }

    seen.set(entry.id, entry);
    selected.push(entry);
  }

  if (conflicts.length) {
    const preview = conflicts.slice(0, 10).map((item) => ({
      id: item.id,
      firstRow: item.first.sourceRow,
      duplicateRow: item.duplicate.sourceRow,
      firstCard: item.first.card,
      duplicateCard: item.duplicate.card,
      firstName: item.first.name,
      duplicateName: item.duplicate.name,
      firstPhone: item.first.phone,
      duplicatePhone: item.duplicate.phone,
    }));
    throw new Error(`发现 ${conflicts.length} 个用户卡ID冲突：${JSON.stringify(preview)}`);
  }

  if (!selected.length) {
    throw new Error("没有任何记录精确命中用户确认的续卡卡项");
  }

  const matchedCards = new Set(selected.map((entry) => entry.card));
  const unmatchedCards = rules
    .map((rule) => rule.card)
    .filter((card) => !matchedCards.has(card));
  if (unmatchedCards.length) {
    throw new Error(`用户确认的卡项在源表中没有匹配记录：${unmatchedCards.join("、")}`);
  }

  selected.sort((left, right) => {
    const categoryDifference = categoryOrder.get(left.card) - categoryOrder.get(right.card);
    return categoryDifference || left.sourceOrder - right.sourceOrder;
  });

  return { selected, duplicatesSkipped };
}

function valuesEqual(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function snapshotOtherSheets(workbook, mainSheetName) {
  return workbook.worksheets.items
    .filter((sheet) => sheet.name !== mainSheetName)
    .map((sheet) => ({
      name: sheet.name,
      values: sheet.getUsedRange()?.values ?? [],
    }));
}

function scanFormulaErrors(values) {
  const errors = [];
  const pattern = /^#(REF|DIV\/0|VALUE|NAME\?|N\/A|NUM|NULL|SPILL|CALC)!?$/;
  for (const [rowIndex, row] of values.entries()) {
    for (const [colIndex, value] of row.entries()) {
      if (typeof value === "string" && pattern.test(value.trim())) {
        errors.push({ cell: `${columnName(colIndex)}${rowIndex + 1}`, value });
      }
    }
  }
  return errors;
}

async function loadArtifactTool(nodeModulesPath) {
  const modulePath = path.join(
    nodeModulesPath,
    "@oai",
    "artifact-tool",
    "dist",
    "artifact_tool.mjs",
  );
  await fs.access(modulePath);
  return import(pathToFileURL(modulePath).href);
}

async function renderPreviews(workbook, previewDir, sheetName, rowCount, boundaryRow) {
  if (!previewDir) return [];
  await fs.mkdir(previewDir, { recursive: true });
  const ranges = [
    ["top.png", "A1:G20"],
    [
      "boundary.png",
      `A${Math.max(1, boundaryRow - 3)}:G${Math.min(rowCount, boundaryRow + 7)}`,
    ],
    ["bottom.png", `A${Math.max(1, rowCount - 14)}:G${rowCount}`],
  ];
  const paths = [];
  for (const [filename, range] of ranges) {
    const blob = await workbook.render({
      sheetName,
      range,
      scale: 2,
      format: "png",
    });
    const outputPath = path.join(previewDir, filename);
    await fs.writeFile(outputPath, new Uint8Array(await blob.arrayBuffer()));
    paths.push(outputPath);
  }
  return paths;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const mode = normalize(args.mode || "build").toLowerCase();
  if (!["inspect", "build"].includes(mode)) {
    throw new Error("--mode 只能是 inspect 或 build");
  }
  const nodeModulesPath = path.resolve(required(args, "node-modules"));
  const sourcePath = path.resolve(required(args, "source"));
  if (path.extname(sourcePath).toLowerCase() !== ".xlsx") {
    throw new Error("原始签到记录必须是 .xlsx 文件");
  }
  await fs.access(sourcePath);
  const { FileBlob, SpreadsheetFile } = await loadArtifactTool(nodeModulesPath);
  const sourceWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
  const sourceData = findSourceData(sourceWorkbook);
  if (mode === "inspect") {
    console.log(JSON.stringify(inspectSourceData(sourceData), null, 2));
    return;
  }

  const outputPath = path.resolve(required(args, "output"));
  const reason = required(args, "reason");
  const templatePath = path.resolve(args.template ?? defaultTemplatePath);
  const previewDir = args["preview-dir"] ? path.resolve(args["preview-dir"]) : null;
  if (path.extname(outputPath).toLowerCase() !== ".xlsx") {
    throw new Error("输出路径必须以 .xlsx 结尾");
  }
  await fs.access(templatePath);
  const { rules, rulesPath, mode: rulesMode } = await loadRules(args);
  const rulesByCard = new Map(rules.map((rule) => [rule.card, rule]));
  const { selected, duplicatesSkipped } = collectRows(sourceData, rules);

  const templateWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(templatePath));
  const mainSheet = templateWorkbook.worksheets.items.find(
    (sheet) => sheet.name === "批量续卡模板",
  ) ?? templateWorkbook.worksheets.items[0];
  if (!mainSheet) throw new Error("模板没有可写入的工作表");

  const preservedSheets = snapshotOtherSheets(templateWorkbook, mainSheet.name);
  const outputRows = [
    expectedHeaders,
    ...selected.map((entry) => [
      entry.id,
      entry.name,
      entry.phone,
      entry.card,
      entry.rule.days,
      entry.rule.times,
      reason,
    ]),
  ];

  const existing = mainSheet.getUsedRange();
  if (existing) existing.clear({ applyTo: "contents" });
  mainSheet.getRange(`A1:G${outputRows.length}`).values = outputRows;
  mainSheet.getRange(`A2:D${outputRows.length}`).format.numberFormat = "@";
  mainSheet.getRange(`E2:F${outputRows.length}`).format.numberFormat = "0";
  mainSheet.getRange(`G2:G${outputRows.length}`).format.numberFormat = "@";
  mainSheet.getRange(`A1:G${outputRows.length}`).format.verticalAlignment = "center";
  mainSheet.getRange(`A1:A${outputRows.length}`).format.columnWidthPx = 320;
  mainSheet.getRange(`B1:B${outputRows.length}`).format.columnWidthPx = 100;
  mainSheet.getRange(`C1:C${outputRows.length}`).format.columnWidthPx = 120;
  mainSheet.getRange(`D1:D${outputRows.length}`).format.columnWidthPx = 100;
  mainSheet.getRange(`E1:F${outputRows.length}`).format.columnWidthPx = 90;
  mainSheet.getRange(`G1:G${outputRows.length}`).format.columnWidthPx = 210;

  templateWorkbook.recalculate();
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const exported = await SpreadsheetFile.exportXlsx(templateWorkbook);
  await exported.save(outputPath);

  const verifiedWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
  const verifiedSheet = verifiedWorkbook.worksheets.getItem(mainSheet.name);
  if (!verifiedSheet) throw new Error("导出后找不到主工作表");
  const verifiedValues = verifiedSheet.getUsedRange()?.values ?? [];
  const verifiedHeader = (verifiedValues[0] ?? []).map(normalize);
  if (!valuesEqual(verifiedHeader, expectedHeaders)) {
    throw new Error(`导出表头不正确：${JSON.stringify(verifiedHeader)}`);
  }

  const dataRows = verifiedValues.slice(1).filter((row) => normalize(row[0]));
  const seenIds = new Set();
  const invalidRows = [];
  for (const [index, row] of dataRows.entries()) {
    const id = normalize(row[0]);
    const name = normalize(row[1]);
    const phone = normalize(row[2]);
    const card = normalize(row[3]);
    const days = Number(row[4]);
    const times = Number(row[5]);
    const rowReason = normalize(row[6]);
    const duplicate = seenIds.has(id);
    seenIds.add(id);
    const rule = rulesByCard.get(card);
    const validAction = Boolean(rule) && days === rule.days && times === rule.times;
    if (!id || !name || !phone || !validAction || rowReason !== reason || duplicate) {
      invalidRows.push({
        row: index + 2,
        id,
        name,
        phone,
        card,
        days,
        times,
        reason: rowReason,
        duplicate,
      });
    }
  }

  if (dataRows.length !== selected.length) {
    throw new Error(`导出记录数不一致：预期 ${selected.length}，实际 ${dataRows.length}`);
  }
  if (invalidRows.length) {
    throw new Error(`导出校验失败：${JSON.stringify(invalidRows.slice(0, 20))}`);
  }

  const formulaErrors = scanFormulaErrors(verifiedValues);
  if (formulaErrors.length) {
    throw new Error(`发现公式错误：${JSON.stringify(formulaErrors)}`);
  }

  const verifiedOtherSheets = snapshotOtherSheets(verifiedWorkbook, mainSheet.name);
  if (!valuesEqual(verifiedOtherSheets, preservedSheets)) {
    throw new Error("模板的其他工作表在导出后发生了变化");
  }

  const firstBoundaryIndex = selected.findIndex(
    (entry, index) => index > 0 && entry.card !== selected[0].card,
  );
  const boundaryRow = firstBoundaryIndex >= 0 ? firstBoundaryIndex + 2 : outputRows.length;
  const previews = await renderPreviews(
    verifiedWorkbook,
    previewDir,
    mainSheet.name,
    outputRows.length,
    boundaryRow,
  );

  const summary = {
    outputPath,
    sourceSheet: sourceData.sheet.name,
    reason,
    outputRows: selected.length,
    duplicatesSkipped,
    rulesMode,
    rulesPath,
    rules,
    cardCounts: Object.fromEntries(
      rules.map((rule) => [
        rule.card,
        selected.filter((entry) => entry.card === rule.card).length,
      ]),
    ),
    rowsWithDayChange: selected.filter((entry) => entry.rule.days !== 0).length,
    rowsWithTimeChange: selected.filter((entry) => entry.rule.times !== 0).length,
    formulaErrors: [],
    otherSheetsPreserved: true,
    previews,
  };
  console.log(JSON.stringify(summary, null, 2));
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack : String(error));
  process.exitCode = 1;
});
