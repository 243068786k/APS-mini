"use client";

import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type PlanRow = {
  id: string;
  product: string;
  specification: string;
  batch: string;
  originalShipDate: string;
  permittedShipDate: string;
  requiredFinishDate: string;
  status: string;
};

type ScheduleRow = {
  id: string;
  workshop: string;
  equipment: string;
  product: string;
  batch: string;
  process: string;
  start: string;
  end: string;
  status: string;
};

type StandardTimeRow = {
  id: string;
  product: string;
  workshop: string;
  process: string;
  standardHours: number;
  tolerancePercent: number;
  note: string;
};

type Rules = {
  warningDays: number;
  defaultClearanceHours: number;
  processClearanceRules: ProcessClearanceRule[];
  testingDays: number;
  standardTimeTolerancePercent: number;
};

type ProcessClearanceRule = {
  id: string;
  process: string;
  hours: number;
};

type ImportState = {
  plan: boolean;
  schedule: boolean;
  standardTime: boolean;
};

type MergeResult<T> = {
  rows: T[];
  added: number;
  updated: number;
};

type AlertLevel = "高" | "中" | "低";
type Alert = {
  id: string;
  type: string;
  level: AlertLevel;
  batch: string;
  product: string;
  time: string;
  detail: string;
  action: string;
};

type Tab =
  | "overview"
  | "alerts"
  | "schedule"
  | "plan"
  | "standardTime"
  | "rules";

const DEMO_PLAN: PlanRow[] = [
  {
    id: "p1",
    product: "YF013片",
    specification: "50 mg",
    batch: "32607094",
    originalShipDate: "2026-08-16",
    permittedShipDate: "2026-08-20",
    requiredFinishDate: "2026-08-13",
    status: "待生产",
  },
  {
    id: "p2",
    product: "YF013片",
    specification: "50 mg",
    batch: "32607095",
    originalShipDate: "2026-08-16",
    permittedShipDate: "2026-08-20",
    requiredFinishDate: "2026-08-13",
    status: "待生产",
  },
  {
    id: "p3",
    product: "YF029片",
    specification: "100 mg",
    batch: "32607108",
    originalShipDate: "2026-08-05",
    permittedShipDate: "2026-08-08",
    requiredFinishDate: "2026-08-01",
    status: "生产中",
  },
  {
    id: "p4",
    product: "YF060片",
    specification: "25 mg",
    batch: "32607116",
    originalShipDate: "2026-08-28",
    permittedShipDate: "2026-08-28",
    requiredFinishDate: "2026-08-21",
    status: "待生产",
  },
];

const DEMO_SCHEDULE: ScheduleRow[] = [
  {
    id: "s1",
    workshop: "二车间",
    equipment: "压片机-01",
    product: "YF013片",
    batch: "32607094",
    process: "压片",
    start: "2026-07-29T08:00",
    end: "2026-07-29T16:00",
    status: "已排产",
  },
  {
    id: "s2",
    workshop: "二车间",
    equipment: "压片机-01",
    product: "YF029片",
    batch: "32607108",
    process: "压片",
    start: "2026-07-29T15:00",
    end: "2026-07-30T02:00",
    status: "已排产",
  },
  {
    id: "s3",
    workshop: "三车间",
    equipment: "湿法制粒-01",
    product: "YF013片",
    batch: "32607095",
    process: "制粒",
    start: "2026-08-11T08:00",
    end: "2026-08-11T18:00",
    status: "已排产",
  },
  {
    id: "s4",
    workshop: "三车间",
    equipment: "湿法制粒-01",
    product: "临时插单A",
    batch: "TMP26001",
    process: "制粒",
    start: "2026-08-11T19:00",
    end: "2026-08-12T04:00",
    status: "已排产",
  },
];

const DEMO_STANDARD_TIMES: StandardTimeRow[] = [
  {
    id: "st1",
    product: "YF013片",
    workshop: "二车间",
    process: "压片",
    standardHours: 8,
    tolerancePercent: 10,
    note: "示例：每批标准工时",
  },
  {
    id: "st2",
    product: "YF029片",
    workshop: "二车间",
    process: "压片",
    standardHours: 10,
    tolerancePercent: 10,
    note: "示例：每批标准工时",
  },
  {
    id: "st3",
    product: "YF013片",
    workshop: "三车间",
    process: "制粒",
    standardHours: 10,
    tolerancePercent: 10,
    note: "示例：每批标准工时",
  },
];

const DEFAULT_RULES: Rules = {
  warningDays: 7,
  defaultClearanceHours: 4,
  processClearanceRules: [
    { id: "clearance-weighing", process: "称量", hours: 1 },
  ],
  testingDays: 7,
  standardTimeTolerancePercent: 10,
};

const NAV: Array<{ id: Tab; label: string; hint: string }> = [
  { id: "overview", label: "审核总览", hint: "风险与优先事项" },
  { id: "alerts", label: "异常清单", hint: "自动判断结果" },
  { id: "schedule", label: "排产明细", hint: "一批一工序一行" },
  { id: "plan", label: "生产计划", hint: "交期审核底表" },
  { id: "standardTime", label: "标准工时", hint: "产品与车间基准" },
  { id: "rules", label: "规则参数", hint: "预警与清场设置" },
];

const HEADER_ALIASES = {
  plan: {
    product: ["产品名称", "品名", "产品"],
    specification: ["规格", "产品规格"],
    batch: ["批号", "生产批号", "批次"],
    originalShipDate: ["原预计发货日期", "预计发货日期", "原发货日期"],
    permittedShipDate: ["许可发货日期", "发货日期"],
    requiredFinishDate: ["要求生产完成日期", "生产完成日期"],
    status: ["状态", "生产状态"],
  },
  schedule: {
    workshop: ["车间", "生产车间"],
    equipment: ["设备", "设备名称"],
    product: ["产品名称", "品名", "产品"],
    batch: ["批号", "生产批号", "批次"],
    process: ["工序", "生产工序"],
    start: ["开始时间", "计划开始时间", "开始日期"],
    end: ["结束时间", "计划结束时间", "结束日期"],
    status: ["状态", "排产状态"],
  },
  standardTime: {
    product: ["产品名称", "品名", "产品"],
    workshop: ["车间", "生产车间"],
    process: ["工序", "生产工序"],
    standardHours: [
      "标准工时（小时）",
      "标准工时(小时)",
      "标准工时",
      "工时",
    ],
    tolerancePercent: [
      "容许偏差（%）",
      "容许偏差(%)",
      "容许偏差",
      "偏差比例",
    ],
    note: ["备注", "说明"],
  },
};

function uid(prefix: string) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
}

function asDate(value: unknown, withTime = false) {
  if (!value && value !== 0) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const local = new Date(value.getTime() - value.getTimezoneOffset() * 60000);
    return local.toISOString().slice(0, withTime ? 16 : 10);
  }
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      const date = new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M);
      return asDate(date, withTime);
    }
  }
  const text = String(value).trim().replace(/\//g, "-");
  const date = new Date(text);
  if (!Number.isNaN(date.getTime())) return asDate(date, withTime);
  return text;
}

function getValue(row: Record<string, unknown>, aliases: string[]) {
  const key = Object.keys(row).find((item) =>
    aliases.some((alias) => item.trim().replace(/\s/g, "") === alias)
  );
  return key ? row[key] : "";
}

function expandBatch(batch: string) {
  const value = String(batch ?? "").trim();
  const match = value.match(/^(\d{7})(\d{1,3})\s*[-~—至]\s*(\d{1,3})$/);
  if (!match) return [value];
  const [, prefix, startText, endText] = match;
  const start = Number(startText);
  const end = Number(endText);
  if (end < start || end - start > 50) return [value];
  return Array.from({ length: end - start + 1 }, (_, index) =>
    `${prefix}${String(start + index).padStart(startText.length, "0")}`
  );
}

function batchKey(batch: string) {
  const normalized = String(batch ?? "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  const digits = normalized.match(/\d{8}/)?.[0];
  return digits ?? normalized;
}

function normalizedText(value: string) {
  return String(value ?? "").trim().replace(/\s+/g, "").toUpperCase();
}

function dateTime(value: string) {
  const parts = value?.match(
    /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/
  );
  if (parts) {
    const [, year, month, day, hour = "0", minute = "0"] = parts;
    return new Date(
      Date.UTC(
        Number(year),
        Number(month) - 1,
        Number(day),
        Number(hour),
        Number(minute)
      )
    );
  }
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function formatDate(value: string, withTime = false) {
  if (!value) return "—";
  const parts = value.match(
    /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/
  );
  if (parts) {
    const [, , month, day, hour, minute] = parts;
    return withTime && hour && minute
      ? `${month}/${day} ${hour}:${minute}`
      : `${month}/${day}`;
  }
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value;
  return new Intl.DateTimeFormat("zh-CN", {
    month: "2-digit",
    day: "2-digit",
    ...(withTime ? { hour: "2-digit", minute: "2-digit", hour12: false } : {}),
    timeZone: "Asia/Shanghai",
  }).format(parsed);
}

function clearanceHoursForProcess(process: string, rules: Rules) {
  const normalized = process.trim().replace(/\s+/g, "");
  const matched = rules.processClearanceRules.find((rule) => {
    const ruleName = rule.process.trim().replace(/\s+/g, "");
    return (
      Boolean(ruleName) &&
      (normalized === ruleName ||
        normalized.includes(ruleName) ||
        ruleName.includes(normalized))
    );
  });
  return matched?.hours ?? rules.defaultClearanceHours;
}

function evaluate(
  plan: PlanRow[],
  schedule: ScheduleRow[],
  rules: Rules,
  standardTimes: StandardTimeRow[]
): Alert[] {
  const alerts: Alert[] = [];
  const sorted = [...schedule].sort(
    (a, b) => (dateTime(a.start)?.getTime() ?? 0) - (dateTime(b.start)?.getTime() ?? 0)
  );

  for (let i = 0; i < sorted.length; i += 1) {
    const current = sorted[i];
    const currentStart = dateTime(current.start);
    const currentEnd = dateTime(current.end);
    if (!currentStart || !currentEnd) continue;

    for (let j = i + 1; j < sorted.length; j += 1) {
      const next = sorted[j];
      const nextStart = dateTime(next.start);
      const nextEnd = dateTime(next.end);
      if (!nextStart || !nextEnd) continue;

      if (
        current.workshop === next.workshop &&
        current.equipment === next.equipment &&
        currentStart < nextEnd &&
        nextStart < currentEnd
      ) {
        alerts.push({
          id: `equipment-${current.id}-${next.id}`,
          type: "设备时间冲突",
          level: "高",
          batch: `${current.batch} / ${next.batch}`,
          product: `${current.product} / ${next.product}`,
          time: `${formatDate(current.start, true)}–${formatDate(next.end, true)}`,
          detail: `${current.workshop} ${current.equipment}在两个任务中重复占用。`,
          action: "调整其中一个任务的开始或结束时间",
        });
      }

      if (
        batchKey(current.batch) === batchKey(next.batch) &&
        current.process === next.process &&
        current.equipment !== next.equipment &&
        currentStart < nextEnd &&
        nextStart < currentEnd
      ) {
        alerts.push({
          id: `cross-device-${current.id}-${next.id}`,
          type: "同批同工序跨设备",
          level: "高",
          batch: current.batch,
          product: current.product,
          time: `${formatDate(current.start, true)}–${formatDate(next.end, true)}`,
          detail: `同一批次的${current.process}同时安排在不同设备。`,
          action: "确认是否重复排产或需要拆分任务",
        });
      }
    }
  }

  const deviceGroups = new Map<string, ScheduleRow[]>();
  sorted.forEach((row) => {
    const key = `${row.workshop}|${row.equipment}`;
    deviceGroups.set(key, [...(deviceGroups.get(key) ?? []), row]);
  });
  deviceGroups.forEach((rows) => {
    rows.forEach((row, index) => {
      const next = rows[index + 1];
      if (!next) return;
      const end = dateTime(row.end);
      const start = dateTime(next.start);
      if (!end || !start || start < end) return;
      const gapHours = (start.getTime() - end.getTime()) / 3600000;
      const process = row.process || next.process || "未填写工序";
      const minimumHours = clearanceHoursForProcess(process, rules);
      if (gapHours < minimumHours) {
        alerts.push({
          id: `clearance-${row.id}-${next.id}`,
          type: "清场间隔不足",
          level: "中",
          batch: `${row.batch} → ${next.batch}`,
          product: `${row.product} → ${next.product}`,
          time: `${gapHours.toFixed(1)} 小时`,
          detail: `${row.equipment}的${process}工序，两项任务之间仅预留${gapHours.toFixed(1)}小时。`,
          action: `该工序至少预留${minimumHours}小时并确认清场方式`,
        });
      }
    });
  });

  const planKeys = new Set(plan.map((row) => batchKey(row.batch)));
  const scheduleKeys = new Set(schedule.map((row) => batchKey(row.batch)));
  schedule.forEach((row) => {
    if (!planKeys.has(batchKey(row.batch))) {
      alerts.push({
        id: `unmatched-${row.id}`,
        type: "未匹配生产计划",
        level: "中",
        batch: row.batch,
        product: row.product,
        time: formatDate(row.start, true),
        detail: "该排产任务未在生产计划底表中找到对应批次。",
        action: "确认是否为临时插单、批号录入差异或漏登计划",
      });
    }
  });

  if (standardTimes.length) {
    const workGroups = new Map<
      string,
      { row: ScheduleRow; hours: number }
    >();
    schedule.forEach((row) => {
      const start = dateTime(row.start);
      const end = dateTime(row.end);
      if (!start || !end || end <= start) return;
      const key = [
        batchKey(row.batch),
        normalizedText(row.product),
        normalizedText(row.workshop),
        normalizedText(row.process),
      ].join("|");
      const current = workGroups.get(key);
      workGroups.set(key, {
        row: current?.row ?? row,
        hours:
          (current?.hours ?? 0) +
          (end.getTime() - start.getTime()) / 3600000,
      });
    });

    workGroups.forEach(({ row, hours }) => {
      const product = normalizedText(row.product);
      const workshop = normalizedText(row.workshop);
      const process = normalizedText(row.process);
      const candidates = standardTimes.filter(
        (standard) =>
          normalizedText(standard.product) === product &&
          normalizedText(standard.workshop) === workshop
      );
      const standard =
        candidates.find(
          (item) => normalizedText(item.process) === process
        ) ?? candidates.find((item) => !normalizedText(item.process));

      if (!standard) {
        alerts.push({
          id: `standard-missing-${row.id}`,
          type: "未维护标准工时",
          level: "低",
          batch: row.batch,
          product: row.product,
          time: `${hours.toFixed(1)} 小时`,
          detail: `${row.workshop}的${row.process || "未填写"}工序未找到该产品的标准工时。`,
          action: "在“标准工时”sheet中补充产品、车间和工序基准",
        });
        return;
      }

      const tolerance =
        Number.isFinite(standard.tolerancePercent) &&
        standard.tolerancePercent >= 0
          ? standard.tolerancePercent
          : rules.standardTimeTolerancePercent;
      const lower = standard.standardHours * (1 - tolerance / 100);
      const upper = standard.standardHours * (1 + tolerance / 100);
      const deviation =
        ((hours - standard.standardHours) / standard.standardHours) * 100;

      if (hours > upper) {
        alerts.push({
          id: `standard-over-${row.id}`,
          type: "排产工时超出标准",
          level: "中",
          batch: row.batch,
          product: row.product,
          time: `${hours.toFixed(1)} / ${standard.standardHours.toFixed(1)} 小时`,
          detail: `${row.workshop}${row.process}排产总时长较标准工时高${Math.abs(
            deviation
          ).toFixed(1)}%，超出±${tolerance}%容许范围。`,
          action: "核对是否包含等待、清场或停机时间，并确认排产时长",
        });
      } else if (hours < lower) {
        alerts.push({
          id: `standard-under-${row.id}`,
          type: "排产工时低于标准",
          level: "中",
          batch: row.batch,
          product: row.product,
          time: `${hours.toFixed(1)} / ${standard.standardHours.toFixed(1)} 小时`,
          detail: `${row.workshop}${row.process}排产总时长较标准工时低${Math.abs(
            deviation
          ).toFixed(1)}%，超出±${tolerance}%容许范围。`,
          action: "核对是否漏排时段、批量不同或标准工时需要更新",
        });
      }
    });
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  plan.forEach((row) => {
    const key = batchKey(row.batch);
    const related = schedule.filter((item) => batchKey(item.batch) === key);
    const permitted = dateTime(row.permittedShipDate || row.originalShipDate);
    let required = dateTime(row.requiredFinishDate);
    if (!required && permitted) {
      required = new Date(permitted);
      required.setDate(required.getDate() - rules.testingDays);
    }

    if (!scheduleKeys.has(key)) {
      alerts.push({
        id: `unscheduled-${row.id}`,
        type: "计划未排产",
        level: required && required < today ? "高" : "中",
        batch: row.batch,
        product: row.product,
        time: required ? formatDate(required.toISOString()) : "未设置",
        detail: "生产计划中已有该批次，但排产明细尚无对应任务。",
        action: "结合要求完成日期、许可发货日期和前序工序状态安排生产",
      });
      return;
    }

    const latestEnd = related
      .map((item) => dateTime(item.end))
      .filter((item): item is Date => Boolean(item))
      .sort((a, b) => b.getTime() - a.getTime())[0];
    if (!required || !latestEnd) return;
    const remainingDays = Math.ceil((required.getTime() - latestEnd.getTime()) / 86400000);
    if (latestEnd > required) {
      alerts.push({
        id: `overdue-${row.id}`,
        type: "交期超期",
        level: "高",
        batch: row.batch,
        product: row.product,
        time: formatDate(required.toISOString()),
        detail: `当前排产结束时间晚于要求生产完成日期${Math.abs(remainingDays)}天。`,
        action: "评估前移排产、调整资源或尽快反馈交付风险",
      });
    } else if (remainingDays <= rules.warningDays) {
      alerts.push({
        id: `near-${row.id}`,
        type: "临近交期",
        level: "低",
        batch: row.batch,
        product: row.product,
        time: formatDate(required.toISOString()),
        detail: `排产完成后距要求生产完成日期仅剩${remainingDays}天。`,
        action: "确认检验周期与放行资源是否可满足",
      });
    }
  });

  const order: Record<AlertLevel, number> = { 高: 0, 中: 1, 低: 2 };
  return alerts.sort((a, b) => order[a.level] - order[b.level]);
}

function parsePlanRows(rows: Record<string, unknown>[], rules: Rules) {
  const result: PlanRow[] = [];
  rows.forEach((row) => {
    const batchRaw = String(getValue(row, HEADER_ALIASES.plan.batch) ?? "").trim();
    if (!batchRaw) return;
    expandBatch(batchRaw).forEach((batch) => {
      const permittedShipDate = asDate(
        getValue(row, HEADER_ALIASES.plan.permittedShipDate)
      );
      let requiredFinishDate = asDate(
        getValue(row, HEADER_ALIASES.plan.requiredFinishDate)
      );
      if (!requiredFinishDate && permittedShipDate) {
        const date = new Date(permittedShipDate);
        date.setDate(date.getDate() - rules.testingDays);
        requiredFinishDate = asDate(date);
      }
      result.push({
        id: uid("p"),
        product: String(getValue(row, HEADER_ALIASES.plan.product) ?? "").trim(),
        specification: String(
          getValue(row, HEADER_ALIASES.plan.specification) ?? ""
        ).trim(),
        batch,
        originalShipDate: asDate(
          getValue(row, HEADER_ALIASES.plan.originalShipDate)
        ),
        permittedShipDate,
        requiredFinishDate,
        status: String(getValue(row, HEADER_ALIASES.plan.status) ?? "").trim(),
      });
    });
  });
  return result;
}

function parseScheduleRows(rows: Record<string, unknown>[]) {
  const result: ScheduleRow[] = [];
  rows.forEach((row) => {
    const batchRaw = String(getValue(row, HEADER_ALIASES.schedule.batch) ?? "").trim();
    if (!batchRaw) return;
    expandBatch(batchRaw).forEach((batch) => {
      result.push({
        id: uid("s"),
        workshop: String(
          getValue(row, HEADER_ALIASES.schedule.workshop) ?? ""
        ).trim(),
        equipment: String(
          getValue(row, HEADER_ALIASES.schedule.equipment) ?? ""
        ).trim(),
        product: String(
          getValue(row, HEADER_ALIASES.schedule.product) ?? ""
        ).trim(),
        batch,
        process: String(
          getValue(row, HEADER_ALIASES.schedule.process) ?? ""
        ).trim(),
        start: asDate(getValue(row, HEADER_ALIASES.schedule.start), true),
        end: asDate(getValue(row, HEADER_ALIASES.schedule.end), true),
        status: String(
          getValue(row, HEADER_ALIASES.schedule.status) ?? ""
        ).trim(),
      });
    });
  });
  return result;
}

function parseStandardTimeRows(rows: Record<string, unknown>[]) {
  const result: StandardTimeRow[] = [];
  rows.forEach((row) => {
    const product = String(
      getValue(row, HEADER_ALIASES.standardTime.product) ?? ""
    ).trim();
    const workshop = String(
      getValue(row, HEADER_ALIASES.standardTime.workshop) ?? ""
    ).trim();
    const standardHours = Number(
      getValue(row, HEADER_ALIASES.standardTime.standardHours)
    );
    if (
      !product ||
      !workshop ||
      !Number.isFinite(standardHours) ||
      standardHours <= 0
    ) {
      return;
    }
    const toleranceRaw = getValue(
      row,
      HEADER_ALIASES.standardTime.tolerancePercent
    );
    result.push({
      id: uid("st"),
      product,
      workshop,
      process: String(
        getValue(row, HEADER_ALIASES.standardTime.process) ?? ""
      ).trim(),
      standardHours,
      tolerancePercent:
        toleranceRaw === "" || toleranceRaw === null
          ? Number.NaN
          : Number(toleranceRaw),
      note: String(
        getValue(row, HEADER_ALIASES.standardTime.note) ?? ""
      ).trim(),
    });
  });
  return result;
}

function planMergeKey(row: PlanRow) {
  return batchKey(row.batch);
}

function scheduleBaseKey(row: ScheduleRow) {
  return `${batchKey(row.batch)}|${normalizedText(row.process)}`;
}

function scheduleMergeKey(row: ScheduleRow) {
  return `${scheduleBaseKey(row)}|${normalizedText(row.equipment)}`;
}

function mergePlanRecord(current: PlanRow, incoming: PlanRow): PlanRow {
  return {
    id: current.id,
    product: incoming.product || current.product,
    specification: incoming.specification || current.specification,
    batch: incoming.batch || current.batch,
    originalShipDate: incoming.originalShipDate || current.originalShipDate,
    permittedShipDate: incoming.permittedShipDate || current.permittedShipDate,
    requiredFinishDate:
      incoming.requiredFinishDate || current.requiredFinishDate,
    status: incoming.status || current.status || "待生产",
  };
}

function mergeScheduleRecord(
  current: ScheduleRow,
  incoming: ScheduleRow
): ScheduleRow {
  return {
    id: current.id,
    workshop: incoming.workshop || current.workshop,
    equipment: incoming.equipment || current.equipment,
    product: incoming.product || current.product,
    batch: incoming.batch || current.batch,
    process: incoming.process || current.process,
    start: incoming.start || current.start,
    end: incoming.end || current.end,
    status: incoming.status || current.status || "已排产",
  };
}

function mergePlanRows(
  current: PlanRow[],
  incoming: PlanRow[]
): MergeResult<PlanRow> {
  const rows = current.map((row) => ({ ...row }));
  const indexByKey = new Map(
    rows.map((row, index) => [planMergeKey(row), index])
  );
  let added = 0;
  let updated = 0;

  incoming.forEach((row) => {
    const key = planMergeKey(row);
    const existingIndex = indexByKey.get(key);
    if (existingIndex === undefined) {
      rows.push({ ...row, status: row.status || "待生产" });
      indexByKey.set(key, rows.length - 1);
      added += 1;
      return;
    }
    rows[existingIndex] = mergePlanRecord(rows[existingIndex], row);
    updated += 1;
  });

  return { rows, added, updated };
}

function mergeScheduleRows(
  current: ScheduleRow[],
  incoming: ScheduleRow[]
): MergeResult<ScheduleRow> {
  const rows = current.map((row) => ({ ...row }));
  const exactIndex = new Map(
    rows.map((row, index) => [scheduleMergeKey(row), index])
  );
  const currentBaseIndexes = new Map<string, number[]>();
  rows.forEach((row, index) => {
    const key = scheduleBaseKey(row);
    currentBaseIndexes.set(key, [
      ...(currentBaseIndexes.get(key) ?? []),
      index,
    ]);
  });
  const incomingBaseCounts = new Map<string, number>();
  incoming.forEach((row) => {
    const key = scheduleBaseKey(row);
    incomingBaseCounts.set(key, (incomingBaseCounts.get(key) ?? 0) + 1);
  });

  let added = 0;
  let updated = 0;
  incoming.forEach((row) => {
    const exactKey = scheduleMergeKey(row);
    const baseKey = scheduleBaseKey(row);
    let existingIndex = exactIndex.get(exactKey);

    if (
      existingIndex === undefined &&
      incomingBaseCounts.get(baseKey) === 1 &&
      currentBaseIndexes.get(baseKey)?.length === 1
    ) {
      existingIndex = currentBaseIndexes.get(baseKey)?.[0];
    }

    if (existingIndex === undefined) {
      rows.push({ ...row, status: row.status || "已排产" });
      const newIndex = rows.length - 1;
      exactIndex.set(exactKey, newIndex);
      currentBaseIndexes.set(baseKey, [
        ...(currentBaseIndexes.get(baseKey) ?? []),
        newIndex,
      ]);
      added += 1;
      return;
    }

    rows[existingIndex] = mergeScheduleRecord(rows[existingIndex], row);
    exactIndex.set(scheduleMergeKey(rows[existingIndex]), existingIndex);
    updated += 1;
  });

  return { rows, added, updated };
}

function standardTimeMergeKey(row: StandardTimeRow) {
  return [
    normalizedText(row.product),
    normalizedText(row.workshop),
    normalizedText(row.process),
  ].join("|");
}

function mergeStandardTimeRows(
  current: StandardTimeRow[],
  incoming: StandardTimeRow[]
): MergeResult<StandardTimeRow> {
  const rows = current.map((row) => ({ ...row }));
  const indexByKey = new Map(
    rows.map((row, index) => [standardTimeMergeKey(row), index])
  );
  let added = 0;
  let updated = 0;
  incoming.forEach((row) => {
    const key = standardTimeMergeKey(row);
    const existingIndex = indexByKey.get(key);
    if (existingIndex === undefined) {
      rows.push({ ...row });
      indexByKey.set(key, rows.length - 1);
      added += 1;
      return;
    }
    rows[existingIndex] = { ...row, id: rows[existingIndex].id };
    updated += 1;
  });
  return { rows, added, updated };
}

function isUnchangedDemo<T extends { id: string }>(rows: T[], demo: T[]) {
  return (
    rows.length === demo.length &&
    rows.every((row, index) => row.id === demo[index]?.id)
  );
}

function StatusPill({ level }: { level: AlertLevel }) {
  return <span className={`status status-${level}`}>{level}风险</span>;
}

export default function Home() {
  const [tab, setTab] = useState<Tab>("overview");
  const [plan, setPlan] = useState<PlanRow[]>(DEMO_PLAN);
  const [schedule, setSchedule] = useState<ScheduleRow[]>(DEMO_SCHEDULE);
  const [standardTimes, setStandardTimes] =
    useState<StandardTimeRow[]>(DEMO_STANDARD_TIMES);
  const [rules, setRules] = useState<Rules>(DEFAULT_RULES);
  const [importState, setImportState] = useState<ImportState>({
    plan: false,
    schedule: false,
    standardTime: false,
  });
  const [loaded, setLoaded] = useState(false);
  const [query, setQuery] = useState("");
  const [level, setLevel] = useState<"全部" | AlertLevel>("全部");
  const [alertType, setAlertType] = useState("全部类型");
  const [notice, setNotice] = useState(
    "首次上传会替换对应示例数据；后续上传将保留历史记录并自动合并重复信息"
  );
  const fileRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    queueMicrotask(() => {
      try {
        const saved = localStorage.getItem("aps-mini-data");
        if (saved) {
          const data = JSON.parse(saved);
          if (Array.isArray(data.plan)) setPlan(data.plan);
          if (Array.isArray(data.schedule)) setSchedule(data.schedule);
          setStandardTimes(
            Array.isArray(data.standardTimes) ? data.standardTimes : []
          );
          setImportState({
            plan:
              data.importState?.plan ??
              (Array.isArray(data.plan) &&
                !isUnchangedDemo(data.plan, DEMO_PLAN)),
            schedule:
              data.importState?.schedule ??
              (Array.isArray(data.schedule) &&
                !isUnchangedDemo(data.schedule, DEMO_SCHEDULE)),
            standardTime:
              data.importState?.standardTime ??
              (Array.isArray(data.standardTimes) &&
                data.standardTimes.length > 0),
          });
          if (data.rules) {
            const savedRules = data.rules as Partial<Rules> & {
              clearanceHours?: number;
            };
            setRules({
              warningDays:
                savedRules.warningDays ?? DEFAULT_RULES.warningDays,
              testingDays:
                savedRules.testingDays ?? DEFAULT_RULES.testingDays,
              defaultClearanceHours:
                savedRules.defaultClearanceHours ??
                savedRules.clearanceHours ??
                DEFAULT_RULES.defaultClearanceHours,
              processClearanceRules: Array.isArray(
                savedRules.processClearanceRules
              )
                ? savedRules.processClearanceRules
                : DEFAULT_RULES.processClearanceRules,
              standardTimeTolerancePercent:
                savedRules.standardTimeTolerancePercent ??
                DEFAULT_RULES.standardTimeTolerancePercent,
            });
          }
        }
      } catch {
        setNotice("本地缓存读取失败，已使用示例数据");
      }
      setLoaded(true);
    });
  }, []);

  useEffect(() => {
    if (!loaded) return;
    localStorage.setItem(
      "aps-mini-data",
      JSON.stringify({ plan, schedule, standardTimes, rules, importState })
    );
  }, [loaded, plan, schedule, standardTimes, rules, importState]);

  const alerts = useMemo(
    () => evaluate(plan, schedule, rules, standardTimes),
    [plan, schedule, rules, standardTimes]
  );
  const alertTypes = useMemo(
    () =>
      Array.from(new Set(alerts.map((item) => item.type))).map((type) => ({
        type,
        count: alerts.filter((item) => item.type === type).length,
      })),
    [alerts]
  );
  const filteredAlerts = useMemo(
    () =>
      alerts.filter((item) => {
        const matchesLevel = level === "全部" || item.level === level;
        const matchesType =
          alertType === "全部类型" || item.type === alertType;
        const haystack = `${item.type}${item.batch}${item.product}${item.detail}`;
        return (
          matchesLevel &&
          matchesType &&
          haystack.toLowerCase().includes(query.toLowerCase())
        );
      }),
    [alerts, level, alertType, query]
  );
  const summary = useMemo(
    () => ({
      high: alerts.filter((item) => item.level === "高").length,
      medium: alerts.filter((item) => item.level === "中").length,
      low: alerts.filter((item) => item.level === "低").length,
      unscheduled: alerts.filter((item) => item.type === "计划未排产").length,
    }),
    [alerts]
  );

  async function importWorkbook(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const planSheet =
        workbook.Sheets["生产计划"] ??
        workbook.Sheets[workbook.SheetNames.find((name) => name.includes("生产计划")) ?? ""];
      const scheduleSheet =
        workbook.Sheets["排产明细"] ??
        workbook.Sheets[workbook.SheetNames.find((name) => name.includes("排产")) ?? ""];
      const standardTimeSheet =
        workbook.Sheets["标准工时"] ??
        workbook.Sheets[
          workbook.SheetNames.find((name) => name.includes("标准工时")) ?? ""
        ];
      const nextPlan = planSheet
        ? parsePlanRows(
            XLSX.utils.sheet_to_json<Record<string, unknown>>(planSheet, { defval: "" }),
            rules
          )
        : [];
      const nextSchedule = scheduleSheet
        ? parseScheduleRows(
            XLSX.utils.sheet_to_json<Record<string, unknown>>(scheduleSheet, {
              defval: "",
            })
          )
        : [];
      const nextStandardTimes = standardTimeSheet
        ? parseStandardTimeRows(
            XLSX.utils.sheet_to_json<Record<string, unknown>>(
              standardTimeSheet,
              { defval: "" }
            )
          )
        : [];
      if (
        !nextPlan.length &&
        !nextSchedule.length &&
        !nextStandardTimes.length
      ) {
        setNotice(
          "未识别到“生产计划”“排产明细”或“标准工时”工作表，请先下载模板"
        );
        return;
      }
      const messages: string[] = [];
      const nextImportState = { ...importState };

      if (nextPlan.length) {
        if (importState.plan) {
          const merged = mergePlanRows(plan, nextPlan);
          setPlan(merged.rows);
          messages.push(
            `生产计划新增${merged.added}条、合并更新${merged.updated}条`
          );
        } else {
          setPlan(
            nextPlan.map((row) => ({
              ...row,
              status: row.status || "待生产",
            }))
          );
          messages.push(`生产计划首次导入${nextPlan.length}条`);
        }
        nextImportState.plan = true;
      }

      if (nextSchedule.length) {
        if (importState.schedule) {
          const merged = mergeScheduleRows(schedule, nextSchedule);
          setSchedule(merged.rows);
          messages.push(
            `排产记录新增${merged.added}条、合并更新${merged.updated}条`
          );
        } else {
          setSchedule(
            nextSchedule.map((row) => ({
              ...row,
              status: row.status || "已排产",
            }))
          );
          messages.push(`排产记录首次导入${nextSchedule.length}条`);
        }
        nextImportState.schedule = true;
      }

      if (nextStandardTimes.length) {
        if (importState.standardTime) {
          const merged = mergeStandardTimeRows(
            standardTimes,
            nextStandardTimes
          );
          setStandardTimes(merged.rows);
          messages.push(
            `标准工时新增${merged.added}条、合并更新${merged.updated}条`
          );
        } else {
          setStandardTimes(nextStandardTimes);
          messages.push(`标准工时首次导入${nextStandardTimes.length}条`);
        }
        nextImportState.standardTime = true;
      }

      setImportState(nextImportState);
      setNotice(
        `${messages.join("；")}，有效历史数据已保留并重新完成异常判断`
      );
    } catch {
      setNotice("表格读取失败，请确认文件为有效的 Excel 或 CSV 文件");
    } finally {
      event.target.value = "";
    }
  }

  function downloadTemplate() {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(
        DEMO_PLAN.map((row) => ({
          产品名称: row.product,
          规格: row.specification,
          批号: row.batch,
          原预计发货日期: row.originalShipDate,
          许可发货日期: row.permittedShipDate,
          要求生产完成日期: row.requiredFinishDate,
          状态: row.status,
        }))
      ),
      "生产计划"
    );
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(
        DEMO_SCHEDULE.map((row) => ({
          车间: row.workshop,
          设备: row.equipment,
          产品名称: row.product,
          批号: row.batch,
          工序: row.process,
          开始时间: row.start,
          结束时间: row.end,
          状态: row.status,
        }))
      ),
      "排产明细"
    );
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(
        DEMO_STANDARD_TIMES.map((row) => ({
          产品名称: row.product,
          车间: row.workshop,
          工序: row.process,
          "标准工时（小时）": row.standardHours,
          "容许偏差（%）": row.tolerancePercent,
          备注: row.note,
        }))
      ),
      "标准工时"
    );
    XLSX.writeFile(workbook, "APS-mini导入模板.xlsx");
  }

  function exportAlerts() {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      XLSX.utils.json_to_sheet(
        filteredAlerts.map((item) => ({
          风险等级: item.level,
          异常类型: item.type,
          产品: item.product,
          批号: item.batch,
          时间或交期: item.time,
          异常说明: item.detail,
          建议处理: item.action,
        }))
      ),
      "异常清单"
    );
    XLSX.writeFile(workbook, "APS-mini当前筛选异常.xlsx");
  }

  function resetDemo() {
    setPlan(DEMO_PLAN);
    setSchedule(DEMO_SCHEDULE);
    setStandardTimes(DEMO_STANDARD_TIMES);
    setRules(DEFAULT_RULES);
    setImportState({ plan: false, schedule: false, standardTime: false });
    setNotice("已恢复示例数据");
  }

  function updateProcessClearanceRule(
    id: string,
    patch: Partial<Pick<ProcessClearanceRule, "process" | "hours">>
  ) {
    setRules({
      ...rules,
      processClearanceRules: rules.processClearanceRules.map((rule) =>
        rule.id === id ? { ...rule, ...patch } : rule
      ),
    });
  }

  function addProcessClearanceRule() {
    setRules({
      ...rules,
      processClearanceRules: [
        ...rules.processClearanceRules,
        { id: uid("clearance"), process: "", hours: rules.defaultClearanceHours },
      ],
    });
  }

  function removeProcessClearanceRule(id: string) {
    setRules({
      ...rules,
      processClearanceRules: rules.processClearanceRules.filter(
        (rule) => rule.id !== id
      ),
    });
  }

  const navCount: Partial<Record<Tab, number>> = {
    alerts: alerts.length,
    schedule: schedule.length,
    plan: plan.length,
    standardTime: standardTimes.length,
  };

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark">A</div>
          <div>
            <strong>APS-mini</strong>
            <span>排产审核工作台</span>
          </div>
        </div>
        <nav>
          {NAV.map((item) => (
            <button
              className={tab === item.id ? "nav-item active" : "nav-item"}
              key={item.id}
              onClick={() => setTab(item.id)}
            >
              <span>
                <b>{item.label}</b>
                <small>{item.hint}</small>
              </span>
              {navCount[item.id] !== undefined && (
                <em>{navCount[item.id]}</em>
              )}
            </button>
          ))}
        </nav>
        <div className="sidebar-note">
          <span className="dot" />
          <p>
            <b>自动保存已开启</b>
            <small>数据仅保存在当前浏览器</small>
          </p>
        </div>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <p className="eyebrow">生产计划 · 排产审核 · 风险闭环</p>
            <h1>{NAV.find((item) => item.id === tab)?.label}</h1>
          </div>
          <div className="top-actions">
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={importWorkbook}
              hidden
            />
            <button className="button secondary" onClick={downloadTemplate}>
              下载导入模板
            </button>
            <button className="button primary" onClick={() => fileRef.current?.click()}>
              上传 Excel 并合并审核
            </button>
          </div>
        </header>

        <div className="notice">
          <span>i</span>
          <p>{notice}</p>
          <button onClick={() => setNotice("")}>×</button>
        </div>

        {tab === "overview" && (
          <div className="page-stack">
            <section className="metric-grid">
              <article className="metric-card metric-dark">
                <span>排产记录</span>
                <strong>{schedule.length}</strong>
                <small>{plan.length} 个计划批次参与匹配</small>
              </article>
              <article className="metric-card">
                <span>高风险</span>
                <strong className="danger-text">{summary.high}</strong>
                <small>建议优先确认资源或交期</small>
              </article>
              <article className="metric-card">
                <span>中风险</span>
                <strong className="warning-text">{summary.medium}</strong>
                <small>需要核实排产依据</small>
              </article>
              <article className="metric-card">
                <span>计划未排产</span>
                <strong>{summary.unscheduled}</strong>
                <small>按要求完成日期优先处理</small>
              </article>
            </section>

            <section className="panel priority-panel">
              <div className="panel-heading">
                <div>
                  <p className="eyebrow">今日审核顺序</p>
                  <h2>优先处理事项</h2>
                </div>
                <button className="text-button" onClick={() => setTab("alerts")}>
                  查看全部异常 →
                </button>
              </div>
              <div className="priority-list">
                {alerts.slice(0, 5).map((item, index) => (
                  <button
                    key={item.id}
                    className="priority-row"
                    onClick={() => setTab("alerts")}
                  >
                    <span className="priority-index">{String(index + 1).padStart(2, "0")}</span>
                    <StatusPill level={item.level} />
                    <span className="priority-main">
                      <b>{item.type}</b>
                      <small>{item.product} · {item.batch}</small>
                    </span>
                    <span className="priority-detail">{item.detail}</span>
                    <span className="chevron">›</span>
                  </button>
                ))}
                {!alerts.length && (
                  <div className="empty-state">
                    <b>当前未发现异常</b>
                    <span>新增或导入数据后会自动重新判断</span>
                  </div>
                )}
              </div>
            </section>

            <section className="two-column">
              <article className="panel rule-card">
                <p className="eyebrow">审核口径</p>
                <h2>当前规则</h2>
                <dl>
                  <div>
                    <dt>交期依据</dt>
                    <dd>许可发货日期</dd>
                  </div>
                  <div>
                    <dt>交期预警</dt>
                    <dd>提前 {rules.warningDays} 天</dd>
                  </div>
                  <div>
                    <dt>检验周期</dt>
                    <dd>{rules.testingDays} 天</dd>
                  </div>
                  <div>
                    <dt>清场间隔</dt>
                    <dd>按工序设置</dd>
                  </div>
                  <div>
                    <dt>标准工时</dt>
                    <dd>{standardTimes.length} 条基准</dd>
                  </div>
                </dl>
              </article>
              <article className="panel guide-card">
                <p className="eyebrow">推荐流程</p>
                <h2>三步完成审核</h2>
                <ol>
                  <li><span>1</span><p><b>更新生产计划</b><small>录入许可发货日期和待生产批次</small></p></li>
                  <li><span>2</span><p><b>维护排产明细</b><small>按车间、设备和开始时间排序</small></p></li>
                  <li><span>3</span><p><b>处理异常清单</b><small>优先确认高风险和未排产批次</small></p></li>
                </ol>
              </article>
            </section>
          </div>
        )}

        {tab === "alerts" && (
          <section className="panel table-panel">
            <div className="panel-heading table-tools">
              <div>
                <p className="eyebrow">自动判断结果</p>
                <h2>{filteredAlerts.length} 条异常待核实</h2>
              </div>
              <div className="filters">
                <input
                  value={query}
                  onChange={(event) => setQuery(event.target.value)}
                  placeholder="搜索产品、批号或异常"
                  aria-label="搜索异常"
                />
                <select
                  className="type-filter"
                  value={alertType}
                  onChange={(event) => setAlertType(event.target.value)}
                  aria-label="按异常类型筛选"
                >
                  <option value="全部类型">全部异常类型（{alerts.length}）</option>
                  {alertTypes.map((item) => (
                    <option value={item.type} key={item.type}>
                      {item.type}（{item.count}）
                    </option>
                  ))}
                </select>
                <select
                  value={level}
                  onChange={(event) =>
                    setLevel(event.target.value as "全部" | AlertLevel)
                  }
                  aria-label="按风险等级筛选"
                >
                  <option value="全部">全部风险</option>
                  <option>高</option>
                  <option>中</option>
                  <option>低</option>
                </select>
                {(query || level !== "全部" || alertType !== "全部类型") && (
                  <button
                    className="button secondary"
                    onClick={() => {
                      setQuery("");
                      setLevel("全部");
                      setAlertType("全部类型");
                    }}
                  >
                    清除筛选
                  </button>
                )}
                <button className="button secondary" onClick={exportAlerts}>
                  导出当前筛选
                </button>
              </div>
            </div>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>风险</th>
                    <th>异常类型</th>
                    <th>产品 / 批号</th>
                    <th>时间或交期</th>
                    <th>异常说明</th>
                    <th>建议处理</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredAlerts.map((item) => (
                    <tr key={item.id}>
                      <td><StatusPill level={item.level} /></td>
                      <td><b>{item.type}</b></td>
                      <td>{item.product}<small className="cell-sub">{item.batch}</small></td>
                      <td>{item.time}</td>
                      <td>{item.detail}</td>
                      <td>{item.action}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!filteredAlerts.length && (
                <div className="empty-state table-empty">
                  <b>没有符合筛选条件的异常</b>
                  <span>可调整筛选条件或上传最新表格</span>
                </div>
              )}
            </div>
          </section>
        )}

        {tab === "schedule" && (
          <section className="panel table-panel">
            <div className="panel-heading">
              <div>
                <p className="eyebrow">唯一排产数据底表</p>
                <h2>排产明细</h2>
              </div>
              <span className="helper">连写批号导入时自动拆分为一批一行</span>
            </div>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>车间</th><th>设备</th><th>产品</th><th>批号</th>
                    <th>工序</th><th>开始时间</th><th>结束时间</th><th>状态</th><th />
                  </tr>
                </thead>
                <tbody>
                  {schedule.map((row) => (
                    <tr key={row.id}>
                      <td>{row.workshop}</td><td>{row.equipment}</td><td>{row.product}</td>
                      <td><b>{row.batch}</b></td><td>{row.process}</td>
                      <td>{formatDate(row.start, true)}</td><td>{formatDate(row.end, true)}</td>
                      <td><span className="tag">{row.status}</span></td>
                      <td><button className="delete-button" onClick={() => setSchedule(schedule.filter((item) => item.id !== row.id))}>删除</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        )}

        {tab === "plan" && (
          <section className="panel table-panel">
            <div className="panel-heading">
              <div>
                <p className="eyebrow">交期审核底表</p>
                <h2>生产计划</h2>
              </div>
              <span className="helper">许可发货日期优先于原预计发货日期</span>
            </div>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>产品</th><th>规格</th><th>批号</th><th>原预计发货</th>
                    <th>许可发货</th><th>要求生产完成</th><th>状态</th><th />
                  </tr>
                </thead>
                <tbody>
                  {plan.map((row) => (
                    <tr key={row.id}>
                      <td>{row.product}</td><td>{row.specification || "—"}</td>
                      <td><b>{row.batch}</b></td><td>{formatDate(row.originalShipDate)}</td>
                      <td><b>{formatDate(row.permittedShipDate)}</b></td>
                      <td>{formatDate(row.requiredFinishDate)}</td>
                      <td><span className="tag">{row.status}</span></td>
                      <td><button className="delete-button" onClick={() => setPlan(plan.filter((item) => item.id !== row.id))}>删除</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        )}

        {tab === "standardTime" && (
          <section className="panel table-panel">
            <div className="panel-heading">
              <div>
                <p className="eyebrow">产品 · 车间 · 工序</p>
                <h2>标准工时基准</h2>
              </div>
              <span className="helper">
                同一批次同一工序的多个时段会先合计，再与标准工时比较
              </span>
            </div>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>产品</th>
                    <th>车间</th>
                    <th>工序</th>
                    <th>标准工时</th>
                    <th>容许偏差</th>
                    <th>备注</th>
                    <th />
                  </tr>
                </thead>
                <tbody>
                  {standardTimes.map((row) => (
                    <tr key={row.id}>
                      <td><b>{row.product}</b></td>
                      <td>{row.workshop}</td>
                      <td>{row.process || "全部工序"}</td>
                      <td>{row.standardHours.toFixed(1)} 小时/批</td>
                      <td>
                        ±
                        {Number.isFinite(row.tolerancePercent)
                          ? row.tolerancePercent
                          : rules.standardTimeTolerancePercent}
                        %
                      </td>
                      <td>{row.note || "—"}</td>
                      <td>
                        <button
                          className="delete-button"
                          onClick={() =>
                            setStandardTimes(
                              standardTimes.filter((item) => item.id !== row.id)
                            )
                          }
                        >
                          删除
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {!standardTimes.length && (
                <div className="empty-state table-empty">
                  <b>尚未导入标准工时</b>
                  <span>
                    下载模板，在“标准工时”sheet中填写后上传即可
                  </span>
                </div>
              )}
            </div>
          </section>
        )}

        {tab === "rules" && (
          <div className="rules-layout">
            <section className="panel rules-panel">
              <div className="panel-heading">
                <div>
                  <p className="eyebrow">自动判断参数</p>
                  <h2>规则参数</h2>
                </div>
              </div>
              <label>
                <span><b>交期预警天数</b><small>进入要求生产完成日期前多少天时提示</small></span>
                <input type="number" min="0" value={rules.warningDays} onChange={(event) => setRules({ ...rules, warningDays: Number(event.target.value) })} />
              </label>
              <label>
                <span><b>其他工序默认清场间隔（小时）</b><small>工序未单独配置时使用此值</small></span>
                <input type="number" min="0" step="0.5" value={rules.defaultClearanceHours} onChange={(event) => setRules({ ...rules, defaultClearanceHours: Number(event.target.value) })} />
              </label>
              <label>
                <span><b>标准工时默认容许偏差（%）</b><small>标准工时表未单独填写偏差时使用此值</small></span>
                <input
                  type="number"
                  min="0"
                  step="1"
                  value={rules.standardTimeTolerancePercent}
                  onChange={(event) =>
                    setRules({
                      ...rules,
                      standardTimeTolerancePercent: Number(event.target.value),
                    })
                  }
                />
              </label>
              <div className="process-rules">
                <div className="process-rules-heading">
                  <span>
                    <b>按工序设置清场间隔</b>
                    <small>优先按排产明细中的“工序”字段匹配</small>
                  </span>
                  <button className="button secondary" onClick={addProcessClearanceRule}>
                    新增工序
                  </button>
                </div>
                <div className="process-rules-list">
                  {rules.processClearanceRules.map((rule) => (
                    <div className="process-rule-row" key={rule.id}>
                      <input
                        className="process-name-input"
                        value={rule.process}
                        placeholder="例如：称量"
                        aria-label="工序名称"
                        onChange={(event) =>
                          updateProcessClearanceRule(rule.id, {
                            process: event.target.value,
                          })
                        }
                      />
                      <span>最小间隔</span>
                      <input
                        className="process-hours-input"
                        type="number"
                        min="0"
                        step="0.5"
                        value={rule.hours}
                        aria-label={`${rule.process || "该工序"}最小清场间隔`}
                        onChange={(event) =>
                          updateProcessClearanceRule(rule.id, {
                            hours: Number(event.target.value),
                          })
                        }
                      />
                      <span>小时</span>
                      <button
                        className="delete-button"
                        onClick={() => removeProcessClearanceRule(rule.id)}
                      >
                        删除
                      </button>
                    </div>
                  ))}
                  {!rules.processClearanceRules.length && (
                    <p className="process-rules-empty">
                      暂无单独配置，所有工序使用默认间隔
                    </p>
                  )}
                </div>
              </div>
              <label>
                <span><b>默认检验周期（天）</b><small>未填写要求生产完成日期时，从许可发货日期倒推</small></span>
                <input type="number" min="0" value={rules.testingDays} onChange={(event) => setRules({ ...rules, testingDays: Number(event.target.value) })} />
              </label>
            </section>
            <aside className="panel rules-explain">
              <p className="eyebrow">固定口径</p>
              <h2>批次与日期规则</h2>
              <ul>
                <li><b>批次匹配：</b>优先按批号前 8 位识别同一生产批。</li>
                <li><b>连写批号：</b>如“32607094-95”，导入后拆成两条记录。</li>
                <li><b>清场间隔：</b>优先按工序匹配专用值；未配置工序使用默认值。</li>
                <li><b>标准工时：</b>按产品、车间和工序匹配；同批同工序的分段排产先合计。</li>
                <li><b>交期依据：</b>优先使用生产检验跟踪表中的许可发货日期。</li>
                <li><b>新增数据：</b>上传或删除记录后，全部异常立即重新计算。</li>
              </ul>
              <button className="button secondary full" onClick={resetDemo}>恢复示例数据</button>
            </aside>
          </div>
        )}

        <footer>
          <span>APS-mini · 浏览器本地版</span>
          <span>最近计算：{alerts.length} 条异常</span>
        </footer>
      </section>
    </main>
  );
}
