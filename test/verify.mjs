// 사용법: node test/verify.mjs 2026-07 파일1.xlsx 파일2.xlsx ...
import { createRequire } from 'module';
import fs from 'fs';
import { parseWorkbook, compute, summarize } from '../src/calc.js';
const require = createRequire(import.meta.url);
const XLSX = require(process.env.XLSX_PATH || 'xlsx');
const [period, ...files] = process.argv.slice(2);
const parsed = files.map(f => parseWorkbook(XLSX, new Uint8Array(fs.readFileSync(f)), f.split('/').pop()));
const { lines, checks } = compute(parsed, period);
const s = summarize(lines);
console.log(JSON.stringify({ period, warnings: checks.warnings, outOfPeriod: checks.outOfPeriod, otherDept: checks.otherDept,
  rev: Object.fromEntries(Object.entries(s.segments).map(([k, v]) => [k, { A: v.A, B: v.B, REV: v.REV, JJ: v.JJ, U_A1: v.U_A1, U_A2: v.U_A2, D_C: v.D_C }])), total: s.total.REV }, null, 1));
