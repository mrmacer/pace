/* ─────────────────────────────────────────────────────────────────────────
   FULL ROSTER PATCH — Select Student shows every eligible PACE student on
   open; the search box filters that complete roster.

   Functional where possible: student-select.js is a pure module (required
   directly), and eligibility is exercised through the REAL roster.js loaded
   under node:vm with a stubbed GRAPH — the same source production runs.
   app.js is never loaded into a DOM in this suite (see
   patch010-specialists-notes.test.js), so its wiring is asserted
   structurally, like the other app.js-level tests here.

   Run with: node tests/patch-full-roster.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");
const SELECT = require("../student-select.js");
const app = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");
const html = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
const sw = fs.readFileSync(path.join(ROOT, "sw.js"), "utf8");
const selectSrc = fs.readFileSync(path.join(ROOT, "student-select.js"), "utf8");

/* ── Fixture: a production-shaped roster through the REAL roster.js ───── */

const SCHEMA = {
  "Student First Name": "field_1", "Student Last Name": "field_2", "Teacher": "field_3",
  "Classroom": "field_4", "Active": "field_5", "PACE Enabled": "field_6"
};

function row(id, first, last, { active = true, pace = true } = {}) {
  const item = { id: String(id), field_1: first, field_2: last, field_3: "Teacher X", field_4: "Room 1" };
  item.field_5 = active;
  item.field_6 = pace;
  return item;
}

// 40 eligible students — deliberately far more than the old cap of 8.
const FIRST = ["Ava", "Ben", "Cara", "Dev", "Eli", "Fay", "Gus", "Hana"];
const LAST = ["Abbott", "Baker", "Cruz", "Dunn", "Evans"];
const eligibleRows = [];
FIRST.forEach((f, i) => LAST.forEach((l, j) => eligibleRows.push(row(100 + i * 10 + j, f, l))));

const excludedRows = [
  row(900, "Ina", "Inactive", { active: false, pace: true }),
  row(901, "Pete", "Notpace", { active: true, pace: false }),
  row(902, "Both", "Off", { active: false, pace: false }),
  (() => { const r = row(903, "Blank", "PaceBlank"); delete r.field_6; return r; })() // PACE Enabled cell absent => defaults to not eligible
];

async function loadEligible(rows) {
  const requested = [];
  const context = vm.createContext({
    console,
    CONFIG: { LISTS: { students: "IEP_Students_2026_27" }, STUDENT_ROSTER_FALLBACK: "FALLBACK_SHOULD_NOT_BE_USED" },
    GRAPH: {
      async getListItems(listName) { requested.push(listName); return rows; },
      async getListSchema() { return SCHEMA; }
    }
  });
  vm.runInContext(fs.readFileSync(path.join(ROOT, "roster.js"), "utf8"), context, { filename: "roster.js" });
  const roster = vm.runInContext("ROSTER", context);
  await roster.refresh();
  assert.equal(roster.error, null);
  return { eligible: roster.getPaceEnabled(), requested };
}

const names = list => list.map(s => s.name);

(async () => {
  const { eligible, requested } = await loadEligible([...excludedRows, ...eligibleRows]);

  /* 3/4 + data source: eligibility unchanged, authoritative list unchanged */
  assert.deepEqual(requested, ["IEP_Students_2026_27"], "roster must still come from IEP_Students_2026_27 only");
  assert.equal(eligible.length, 40, "only Active + PACE Enabled students are eligible");
  for (const banned of ["Ina Inactive", "Pete Notpace", "Both Off", "Blank PaceBlank"]) {
    assert.equal(names(eligible).includes(banned), false, `${banned} must stay excluded`);
  }

  /* 1/2: initial render shows the COMPLETE eligible roster — no cap */
  const initial = SELECT.renderModel(eligible, "");
  assert.equal(initial.students.length, 40, "all 40 eligible students must be rendered on open");
  assert.ok(initial.students.length > 8, "the old 8-student cap must be gone");
  assert.equal(new Set(initial.students.map(s => s.id)).size, 40, "no student dropped or duplicated");
  assert.equal(initial.isFiltering, false);

  /* sorting: last name, then first name, deterministic */
  const order = initial.students.map(s => `${s.lastName}, ${s.firstName}`);
  assert.deepEqual(order, order.slice().sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" })));
  assert.equal(order[0], "Abbott, Ava");
  assert.equal(order[order.length - 1], "Evans, Hana");
  assert.deepEqual(names(SELECT.sort(eligible)), names(SELECT.sort(eligible.slice().reverse())), "order must not depend on input order");

  /* label formatting is presentation only */
  assert.equal(SELECT.displayLabel(eligible.find(s => s.name === "Ava Abbott")), "Abbott, Ava");
  assert.equal(SELECT.displayLabel({ id: "x", name: "Only Name" }), "Only Name", "rows without first/last keep their display name");

  /* 5/6/7/8: search runs against the complete roster; partial first/last; case-insensitive */
  assert.equal(SELECT.renderModel(eligible, "ava").students.length, 5, "partial first name (all five Avas)");
  assert.equal(SELECT.renderModel(eligible, "AVA").students.length, 5, "case-insensitive");
  assert.equal(SELECT.renderModel(eligible, "abbo").students.length, 8, "partial last name (all eight Abbotts)");
  assert.equal(SELECT.renderModel(eligible, "aBbOtT").students.length, 8, "case-insensitive last name");
  assert.deepEqual(names(SELECT.renderModel(eligible, "ava abbott").students), ["Ava Abbott"], "first + last");
  assert.deepEqual(names(SELECT.renderModel(eligible, "abbott ava").students), ["Ava Abbott"], "last + first");
  // A student far beyond the old top-8 is still found — search is not limited to what was displayed.
  assert.deepEqual(names(SELECT.renderModel(eligible, "hana evans").students), ["Hana Evans"]);
  assert.deepEqual(names(SELECT.renderModel(eligible, "inactive").students), [], "excluded students are never searchable");

  /* 9: clearing search restores every eligible student */
  const cleared = SELECT.renderModel(eligible, "");
  assert.deepEqual(names(cleared.students), names(initial.students));
  assert.equal(SELECT.renderModel(eligible, "   ").students.length, 40, "whitespace-only query restores the full roster");

  /* 10/11: counts are computed, not hardcoded */
  assert.equal(initial.countText, "40 students available");
  assert.equal(SELECT.renderModel(eligible.slice(0, 13), "").countText, "13 students available", "count follows the actual roster size");
  assert.equal(SELECT.renderModel(eligible, "ava").countText, "5 of 40 students");
  assert.equal(SELECT.renderModel(eligible, "hana evans").countText, "1 of 40 students");
  assert.equal(SELECT.renderModel(eligible.slice(0, 1), "").countText, "1 student available");
  assert.equal(SELECT.renderModel(eligible.slice(0, 1), "zzz").countText, "0 of 1 student");
  assert.equal(/\b40\b|\b164\b/.test(selectSrc.replace(/\/\*[\s\S]*?\*\//g, "")), false, "no hardcoded roster size in the module");

  /* 12: zero matches */
  const none = SELECT.renderModel(eligible, "zzzz");
  assert.equal(none.students.length, 0);
  assert.equal(none.emptyMessage, "No matching PACE-enabled student found.");
  assert.match(none.emptyHint, /Active and PACE Enabled/);
  assert.equal(none.countText, "0 of 40 students");
  assert.equal(SELECT.renderModel(eligible, "ava").emptyMessage, "", "no message when there are matches");
  assert.equal(SELECT.renderModel([], "").emptyMessage, "No PACE-enabled students are available.");
  assert.equal(SELECT.renderModel([], "").countText, "");

  /* the module never mutates students (name stays the SharePoint value) */
  const before = JSON.stringify(eligible);
  SELECT.renderModel(eligible, "ava");
  SELECT.sort(eligible);
  assert.equal(JSON.stringify(eligible), before, "student objects/names must be untouched");

  /* 1/2 structurally: app.js no longer caps or slices the roster */
  assert.doesNotMatch(app, /STUDENT_SUGGESTION_LIMIT/, "the suggestion cap must be removed");
  {
    const start = app.indexOf("function renderStudentGrid(");
    const end = app.indexOf("\n}\n", start);
    const body = app.slice(start, end);
    assert.doesNotMatch(body, /\.slice\(/, "renderStudentGrid must not slice the roster");
    assert.doesNotMatch(body, /Start typing a student's name/, "the blank-state placeholder-instead-of-roster must be gone");
    assert.match(body, /PACE_STUDENT_SELECT\.renderModel\(cachedStudents, filter\)/, "the grid must render the model built from the complete cachedStudents");
    assert.match(body, /model\.students\.map\(studentCardHtml\)/);
  }
  assert.match(app, /addEventListener\("input", e => renderStudentGrid\(e\.target\.value\)\)/, "search still filters live, no Enter/Search button");
  assert.match(app, /document\.getElementById\("studentRosterCount"\)\.textContent = model\.countText/);

  /* UI markup + delivery */
  assert.match(html, /id="studentRosterCount"/);
  assert.match(html, /id="studentSearch"[^>]*search-input-hero/, "the large search field is preserved");
  assert.match(html, /<script src="student-select\.js"><\/script>\s*<script src="pace-data\.js">/, "module loads before app.js, ahead of pace-data");
  assert.match(sw, /"\.\/student-select\.js"/, "new file must be in the service-worker shell list");
  // v17 introduced this patch's assets; later patches bump it further (e.g. v18 for staff corrections).
assert.ok(Number(sw.match(/CACHE_NAME = "pace-tracker-shell-v(\d+)"/)?.[1]) >= 17, "cache version must be v17 or later so installed browsers pick this up");

  /* 13: selecting a student behaves exactly as before */
  assert.match(app, /btn\.addEventListener\("click", \(\) => selectStudent\(btn\.dataset\.studentId\)\)/);
  assert.match(app, /const student = cachedStudents\.find\(s => s\.id === studentId\);\s*\n\s*if \(!student\) return;/, "selection still resolves the student by id from the full cache");
  assert.match(app, /const existing = openForRoom\.find\(v => v\.Student === student\.name\);/, "duplicate-open check still keyed on the unchanged student.name");
  assert.match(app, /STATE\.student = student;/);
  assert.match(app, /studentName: STATE\.student\?\.name \|\| ""/, "the visit's Student value is still student.name, not the display label");
  assert.match(app, /data-student-id="\$\{escHtml\(s\.id\)\}"/);

  /* 14: no SharePoint write behavior touched — the module is pure */
  assert.doesNotMatch(selectSrc.replace(/\/\*[\s\S]*?\*\//g, ""), /GRAPH|PACE_DATA|fetch\(|localStorage|createMappedListItem|updateMappedListItem|_post|_patch/, "student-select.js code must not read or write anything external");

  /* 15: required-Notes behavior from 2806e32 remains intact */
  assert.match(app, /if \(!STATE\.notes \|\| !STATE\.notes\.trim\(\)\)/);
  assert.match(app, /Add a brief note before submitting this PACE visit\./);
  assert.match(html, /for="noteText">Notes\s*<span class="req">\*<\/span>/);
  assert.match(app, /notes: open \? "" : STATE\.notes/);

  console.log("Full-roster Select Student tests passed.");
})().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
