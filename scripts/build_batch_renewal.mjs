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

function findHeaderIndex(header, aliases) {
  for (const alias of aliases) {
    const index = header.indexOf(alias);
    if (index >= 0) return index;
  }
  return -1;
}

function sameEntry(left, right) {
  return left.card === right.card
    && left.action === right.action
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

function collectRows(sourceData, durationCards, countCards) {
  const durationSet = new Set(durationCards);
  const countSet = new Set(countCards);
  const overlap = durationCards.filter((card) => countSet.has(card));
  if (overlap.length) throw new Error(`期限卡与次卡列表重叠：${overlap.join(", ")}`);

  const categoryOrder = new Map(
    [...durationCards, ...countCards].map((card, index) => [card, index]),
  );
  const seen = new Map();
  const selected = [];
  const conflicts = [];
  let duplicatesSkipped = 0;

  for (const [offset, row] of sourceData.values.slice(1).entries()) {
    const sourceRow = offset + 2;
    const card = normalize(row[sourceData.indexes.card]);
    const action = durationSet.has(card) ? "duration" : countSet.has(card) ? "count" : null;
    if (!action) continue;

    const entry = {
      sourceRow,
      sourceOrder: offset,
      id: normalize(row[sourceData.indexes.id]),
      name: normalize(row[sourceData.indexes.name]),
      phone: normalize(row[sourceData.indexes.phone]),
      card,
      action,
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
    throw new Error("没有任何记录精确命中目标期限卡或次卡");
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

async function renderPreviews(workbook, previewDir, sheetName, rowCount, firstCountRow) {
  if (!previewDir) return [];
  await fs.mkdir(previewDir, { recursive: true });
  const ranges = [
    ["top.png", "A1:G20"],
    [
      "boundary.png",
      `A${Math.max(1, firstCountRow - 3)}:G${Math.min(rowCount, firstCountRow + 7)}`,
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
  const nodeModulesPath = path.resolve(required(args, "node-modules"));
  const sourcePath = path.resolve(required(args, "source"));
  const outputPath = path.resolve(required(args, "output"));
  const reason = required(args, "reason");
  const templatePath = path.resolve(args.template ?? defaultTemplatePath);
  const previewDir = args["preview-dir"] ? path.resolve(args["preview-dir"]) : null;
  const durationCards = parseCardList(args["duration-cards"], defaultDurationCards);
  const countCards = parseCardList(args["count-cards"], defaultCountCards);

  if (path.extname(sourcePath).toLowerCase() !== ".xlsx") {
    throw new Error("原始签到记录必须是 .xlsx 文件");
  }
  if (path.extname(outputPath).toLowerCase() !== ".xlsx") {
    throw new Error("输出路径必须以 .xlsx 结尾");
  }

  await fs.access(sourcePath);
  await fs.access(templatePath);
  const { FileBlob, SpreadsheetFile } = await loadArtifactTool(nodeModulesPath);

  const sourceWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sourcePath));
  const sourceData = findSourceData(sourceWorkbook);
  const { selected, duplicatesSkipped } = collectRows(sourceData, durationCards, countCards);

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
      entry.action === "duration" ? 1 : -1,
      entry.action === "count" ? 1 : 0,
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
  const durationSet = new Set(durationCards);
  const countSet = new Set(countCards);
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
    const validAction = durationSet.has(card)
      ? days === 1 && times === 0
      : countSet.has(card) && days === -1 && times === 1;
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

  const firstCountIndex = selected.findIndex((entry) => entry.action === "count");
  const firstCountRow = firstCountIndex >= 0 ? firstCountIndex + 2 : outputRows.length;
  const previews = await renderPreviews(
    verifiedWorkbook,
    previewDir,
    mainSheet.name,
    outputRows.length,
    firstCountRow,
  );

  const summary = {
    outputPath,
    sourceSheet: sourceData.sheet.name,
    reason,
    outputRows: selected.length,
    durationRows: selected.filter((entry) => entry.action === "duration").length,
    countRows: selected.filter((entry) => entry.action === "count").length,
    duplicatesSkipped,
    cardCounts: Object.fromEntries(
      [...durationCards, ...countCards].map((card) => [
        card,
        selected.filter((entry) => entry.card === card).length,
      ]),
    ),
    countCardRule: { days: -1, times: 1, permanent: true },
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
