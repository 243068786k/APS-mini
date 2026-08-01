import assert from "node:assert/strict";
import { readFile, rm, writeFile } from "node:fs/promises";
import test, { after } from "node:test";
import ts from "typescript";

const sourceUrl = new URL("../app/page.tsx", import.meta.url);
const compiledUrl = new URL("../.audit-rules.test-runtime.mjs", import.meta.url);
const source = await readFile(sourceUrl, "utf8");
const { outputText } = ts.transpileModule(source, {
  compilerOptions: {
    jsx: ts.JsxEmit.ReactJSX,
    module: ts.ModuleKind.ESNext,
    target: ts.ScriptTarget.ES2022,
  },
  fileName: "page.tsx",
});
await writeFile(compiledUrl, outputText);
const { evaluate } = await import(`${compiledUrl.href}?run=${Date.now()}`);

after(async () => {
  await rm(compiledUrl, { force: true });
});

const rules = {
  warningDays: 7,
  defaultClearanceHours: 0,
  processClearanceRules: [],
  testingDays: 7,
  standardTimeTolerancePercent: 10,
};

function plan(batch, product = "产品A", specification = "10 mg", status = "待生产") {
  return {
    id: `plan-${batch}`,
    product,
    specification,
    batch,
    originalShipDate: "2099-12-20",
    permittedShipDate: "2099-12-20",
    requiredFinishDate: "2099-12-10",
    status,
  };
}

function schedule({
  id,
  batch,
  process,
  start,
  end,
  product = "产品A",
  specification = "10 mg",
  toolingCode = "M-10",
}) {
  return {
    id,
    workshop: "一车间",
    equipment: "设备1",
    product,
    batch,
    process,
    activityType: "生产",
    specification,
    toolingCode,
    start,
    end,
    status: "已排产",
  };
}

const route = [
  { id: "route-1", product: "产品A", workshop: "一车间", sequence: 1, process: "制粒", required: true, note: "" },
  { id: "route-2", product: "产品A", workshop: "一车间", sequence: 2, process: "压片", required: true, note: "" },
];

function alertTypes(planRows, scheduleRows, routes = [], switchRules = []) {
  return evaluate(planRows, scheduleRows, rules, [], routes, switchRules).map((item) => item.type);
}

test("只排到前序时不误报后续工序缺失", () => {
  const types = alertTypes(
    [plan("BATCH001")],
    [schedule({ id: "s1", batch: "BATCH001", process: "制粒", start: "2099-10-01T08:00", end: "2099-10-01T12:00" })],
    route,
  );
  assert.equal(types.includes("缺失工序"), false);
});

test("跳过必需前序直接排后序时报告缺失工序", () => {
  const types = alertTypes(
    [plan("BATCH002")],
    [schedule({ id: "s2", batch: "BATCH002", process: "压片", start: "2099-10-01T13:00", end: "2099-10-01T17:00" })],
    route,
  );
  assert.equal(types.includes("缺失工序"), true);
});

test("同设备换产品且未安排清场时报告切换活动缺失", () => {
  const switchRules = [{
    id: "switch-product",
    workshop: "一车间",
    process: "压片",
    equipment: "设备1",
    triggerType: "换产品",
    requiredActivity: "清场",
    applicability: "必须",
    note: "",
  }];
  const rows = [
    schedule({ id: "s3", batch: "BATCH003", product: "产品A", process: "压片", start: "2099-10-01T08:00", end: "2099-10-01T12:00" }),
    schedule({ id: "s4", batch: "BATCH004", product: "产品B", process: "压片", start: "2099-10-01T13:00", end: "2099-10-01T17:00" }),
  ];
  const types = alertTypes([plan("BATCH003", "产品A"), plan("BATCH004", "产品B")], rows, [], switchRules);
  assert.equal(types.includes("切换活动缺失"), true);
});

test("模具编码缺失时转为人工确认而非强判异常", () => {
  const switchRules = [{
    id: "switch-tooling",
    workshop: "一车间",
    process: "压片",
    equipment: "设备1",
    triggerType: "换模具",
    requiredActivity: "装机",
    applicability: "必须",
    note: "",
  }];
  const rows = [
    schedule({ id: "s5", batch: "BATCH005", process: "压片", toolingCode: "", start: "2099-10-01T08:00", end: "2099-10-01T12:00" }),
    schedule({ id: "s6", batch: "BATCH006", process: "压片", toolingCode: "", start: "2099-10-01T13:00", end: "2099-10-01T17:00" }),
  ];
  const types = alertTypes([plan("BATCH005"), plan("BATCH006")], rows, [], switchRules);
  assert.equal(types.includes("切换条件待人工确认"), true);
  assert.equal(types.includes("切换活动缺失"), false);
});
