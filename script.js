/****************************************************
 * School Admin Pro - Pro Pack v2 (FULL)
 * ✅ Pagination (Students + Teacher Summary)
 * ✅ Sorting (click header)
 * ✅ Date range filter (Students by PaymentDate r[8])  <-- FIXED
 * ✅ Role permissions (User = view only)
 * ✅ Export TSV (Excel Khmer OK) + Export Teacher PDF
 * ✅ Print Student Report Detailed
 ****************************************************/

const WEB_APP_URL =
  "https://script.google.com/macros/s/AKfycbxA4S8f76hYxZ8izLtwPyxK20YI9li7Eo-QCDKfqCQ3ZZUEECY_odkybKRlzB4Kb_cI0A/exec";

let allStudents = [];
let studentViewRows = []; // after filter+sort
let teacherRows = [];
let teacherViewRows = []; // after search+sort

let currentUserRole = "User";
let currentUsername = "-";

let isEditMode = false;
let originalName = "";

/* ---------------- IMPORTANT INDEX FIX ----------------
   According to your sheet:
   A=r[0] name, B=r[1] gender, C=r[2] grade, D=r[3] teacher, E=r[4] fee
   F=r[5] teacher80, G=r[6] school20, H=r[7] other, I=r[8] Payment Date ✅, J=r[9] timestamp, K=r[10] days
------------------------------------------------------ */
const PAYMENT_DATE_INDEX = 8; // ✅ Column I

/* ---------------- Pagination + Sort State ---------------- */
let studentPage = 1;
let studentRowsPerPage = 20;
let studentSortKey = "name";
let studentSortDir = "asc";

let teacherPage = 1;
let teacherRowsPerPage = 20;
let teacherSortKey = "teacher";
let teacherSortDir = "asc";

/* ---------------- Helpers ---------------- */
function $(id) {
  return document.getElementById(id);
}

function toNumber(val) {
  const s = String(val ?? "").replace(/[^\d.-]/g, "");
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
}

function formatKHR(n) {
  const x = Math.round(Number(n) || 0);
  return x.toLocaleString("en-US") + " ៛";
}

function nowStamp() {
  const d = new Date();
  return d.toLocaleDateString("km-KH") + " " + d.toLocaleTimeString("km-KH");
}

function setLastSync(which) {
  const t = nowStamp();
  const el = $(which === "dashboard" ? "lastSyncDashboard" : "lastSyncStudents");
  if (el) el.innerText = t;
}

/* Parse date safely (supports YYYY-MM-DD and dd/mm/yyyy, etc.) */
function parseDateAny(x) {
  if (!x) return null;

  // If already Date
  if (x instanceof Date && !isNaN(x.getTime())) return x;

  const s = String(x).trim();

  // ISO yyyy-mm-dd
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const d = new Date(s + "T00:00:00");
    return isNaN(d.getTime()) ? null : d;
  }

  // dd/mm/yyyy or dd-mm-yyyy
  const m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) {
    const dd = Number(m[1]),
      mm = Number(m[2]),
      yy = Number(m[3]);
    const d = new Date(yy, mm - 1, dd);
    return isNaN(d.getTime()) ? null : d;
  }

  // try native
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/* ---------------- TSV Download (UTF-8 BOM) ---------------- */
function downloadTSV(filename, text) {
  const BOM = "\ufeff";
  const content = BOM + text.replace(/\n/g, "\r\n");
  const blob = new Blob([content], {
    type: "text/tab-separated-values;charset=utf-8;",
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* ---------------- API Core ---------------- */
async function callAPI(funcName, ...args) {
  const url = `${WEB_APP_URL}?func=${funcName}&args=${encodeURIComponent(
    JSON.stringify(args)
  )}`;
  try {
    const response = await fetch(url);
    return await response.json();
  } catch (error) {
    console.error("API Error:", error);
    return null;
  }
}

/* ---------------- Auth ---------------- */
async function login() {
  const u = $("username")?.value.trim();
  const p = $("password")?.value.trim();

  if (!u || !p) {
    return Swal.fire(
      "តម្រូវការ",
      "សូមបញ្ចូលឈ្មោះអ្នកប្រើប្រាស់ និងពាក្យសម្ងាត់",
      "warning"
    );
  }

  Swal.fire({
    title: "កំពុងផ្ទៀងផ្ទាត់...",
    didOpen: () => Swal.showLoading(),
    allowOutsideClick: false,
  });

  const res = await callAPI("checkLogin", u, p);

  if (res && res.success) {
    currentUserRole = res.role || "User";
    currentUsername = u;

    const loginSec = $("loginSection");
    loginSec.classList.remove("d-flex");
    loginSec.classList.add("d-none");
    $("mainApp").style.display = "block";

    applyPermissions();
    showSection("dashboard");

    Swal.fire({
      icon: "success",
      title: "ជោគជ័យ!",
      text: "អ្នកបានចូលប្រើប្រាស់ដោយជោគជ័យ!",
      timer: 1200,
      showConfirmButton: false,
    });
  } else {
    Swal.fire(
      "បរាជ័យ",
      "សូមបញ្ចូលឈ្មោះអ្នកប្រើប្រាស់ឬពាក្យសម្ងាត់ម្តងទៀត!",
      "error"
    );
  }
}

function logout() {
  Swal.fire({
    title: "តើអ្នកចង់ចាកចេញមែនទេ?",
    icon: "question",
    showCancelButton: true,
    confirmButtonText: "បាទ ចាកចេញ",
    cancelButtonText: "បោះបង់",
  }).then((result) => {
    if (!result.isConfirmed) return;
    location.reload(); // simple + clean
  });
}

function applyPermissions() {
  const isAdmin = currentUserRole === "Admin";

  const rb = $("roleBadge");
  const ub = $("userBadge");
  if (rb) {
    rb.innerText = isAdmin ? "ADMIN" : "USER";
    rb.classList.toggle("pill-admin", isAdmin);
    rb.classList.toggle("pill-user", !isAdmin);
  }
  if (ub) ub.innerHTML = `<i class="bi bi-person-circle"></i> ${currentUsername}`;

  document.querySelectorAll(".admin-only").forEach((el) => {
    el.style.display = isAdmin ? "" : "none";
  });

  const note = document.querySelector(".user-note");
  if (note) note.classList.toggle("d-none", isAdmin);

  if (!isAdmin) {
    window.openStudentModal = () =>
      Swal.fire("Permission", "User អាចមើលបានតែប៉ុណ្ណោះ។", "info");
    window.editStudent = () =>
      Swal.fire("Permission", "User អាចមើលបានតែប៉ុណ្ណោះ។", "info");
    window.confirmDelete = () =>
      Swal.fire("Permission", "User អាចមើលបានតែប៉ុណ្ណោះ។", "info");
    window.submitStudent = () =>
      Swal.fire("Permission", "User អាចមើលបានតែប៉ុណ្ណោះ។", "info");
  }
}

/* ---------------- Navigation ---------------- */
function showSection(section) {
  $("dashboardSection").style.display = section === "dashboard" ? "block" : "none";
  $("studentSection").style.display = section === "students" ? "block" : "none";

  if (section === "dashboard") loadDashboard();
  if (section === "students") loadStudents();
}

async function refreshAll() {
  await Promise.allSettled([loadDashboard(), loadStudents()]);
}

/* =========================================================
   DASHBOARD (Teacher Summary v2)
========================================================= */
async function loadDashboard() {
  const res = await callAPI("getTeacherData");
  if (!res || !Array.isArray(res.rows)) {
    return Swal.fire("Network", "មិនអាចទាញទិន្នន័យ Dashboard បានទេ។", "warning");
  }

  teacherRows = res.rows;

  teacherPage = 1;
  teacherRowsPerPage = Number($("teacherRowsPerPage")?.value || 20);

  applyTeacherView();
  computeDashboardFromTeachers(teacherRows);

  setLastSync("dashboard");
  bindTeacherSortEvents();
}

function computeDashboardFromTeachers(rows) {
  let totalStudents = 0,
    totalFee = 0,
    teacher80 = 0,
    school20 = 0;

  rows.forEach((r) => {
    totalStudents += toNumber(r[2]);
    totalFee += toNumber(r[3]);
    teacher80 += toNumber(r[4]);
    school20 += toNumber(r[5]);
  });

  if (teacher80 === 0 && school20 === 0 && totalFee > 0) {
    teacher80 = totalFee * 0.8;
    school20 = totalFee * 0.2;
  }

  $("statsRow").innerHTML = `
    <div class="stat-card accent-purple">
      <div class="label">គ្រូសរុប</div>
      <div class="value">${rows.length.toLocaleString("en-US")}</div>
      <div class="sub">ចំនួនគ្រូទាំងអស់</div>
    </div>

    <div class="stat-card accent-green">
      <div class="label">សិស្សសរុប</div>
      <div class="value">${totalStudents.toLocaleString("en-US")}</div>
      <div class="sub">គណនាពី Teacher Summary</div>
    </div>

    <div class="stat-card accent-blue">
      <div class="label">ទឹកប្រាក់សរុប</div>
      <div class="value" style="color:#16a34a">${formatKHR(totalFee)}</div>
      <div class="sub">ចំណូលសរុបទាំងអស់</div>
    </div>

    <div class="stat-card accent-red">
      <div class="label">សាលា (20%) / គ្រូ (80%)</div>
      <div class="value" style="font-size:18px; line-height:1.2">
        <span style="color:#ef4444">${formatKHR(school20)}</span>
        <span style="color:#94a3b8"> • </span>
        <span style="color:#2563eb">${formatKHR(teacher80)}</span>
      </div>
      <div class="sub">បែងចែក 20% និង 80%</div>
    </div>
  `;
}

/* Teacher view apply: search + sort + paginate */
function applyTeacherView() {
  teacherRowsPerPage = Number($("teacherRowsPerPage")?.value || 20);
  const q = ($("searchTeacher")?.value || "").toLowerCase().trim();

  teacherViewRows = teacherRows
    .map((r) => ({
      teacher: String(r[0] ?? ""),
      gender: String(r[1] ?? ""),
      students: toNumber(r[2]),
      totalFee: toNumber(r[3]),
      teacher80: toNumber(r[4]) || toNumber(r[3]) * 0.8,
      school20: toNumber(r[5]) || toNumber(r[3]) * 0.2,
      raw: r,
    }))
    .filter((o) => !q || o.teacher.toLowerCase().includes(q))
    .sort((a, b) => compareByKey(a, b, teacherSortKey, teacherSortDir));

  teacherPage = clampPage(teacherPage, teacherViewRows.length, teacherRowsPerPage);
  renderTeacherPage();
  updateTeacherSortIndicators();
}

function renderTeacherPage() {
  const { pageItems, startIndex, endIndex, totalPages, totalItems } = paginate(
    teacherViewRows,
    teacherPage,
    teacherRowsPerPage
  );

  $("teacherBody").innerHTML = pageItems
    .map(
      (o) => `
    <tr>
      <td>${escapeHtml(o.teacher)}</td>
      <td>${escapeHtml(o.gender)}</td>
      <td>${o.students}</td>
      <td class="fw-bold text-primary">${formatKHR(o.totalFee)}</td>
      <td class="text-success">${formatKHR(o.teacher80)}</td>
      <td class="text-danger">${formatKHR(o.school20)}</td>
    </tr>
  `
    )
    .join("");

  $("teacherPagePill").innerText = `${teacherPage}/${Math.max(1, totalPages)}`;
  $("teacherPageInfo").innerText = `Showing ${startIndex}-${endIndex} of ${totalItems}`;
}

function teacherPrevPage() {
  teacherPage = Math.max(1, teacherPage - 1);
  renderTeacherPage();
}
function teacherNextPage() {
  const totalPages = Math.max(1, Math.ceil(teacherViewRows.length / teacherRowsPerPage));
  teacherPage = Math.min(totalPages, teacherPage + 1);
  renderTeacherPage();
}

function bindTeacherSortEvents() {
  const ths = document.querySelectorAll("#teacherTable thead th.sortable");
  ths.forEach((th) => {
    th.onclick = () => {
      const key = th.getAttribute("data-key");
      if (!key) return;
      if (teacherSortKey === key) teacherSortDir = teacherSortDir === "asc" ? "desc" : "asc";
      else {
        teacherSortKey = key;
        teacherSortDir = "asc";
      }
      teacherPage = 1;
      applyTeacherView();
    };
  });

  $("searchTeacher")?.addEventListener("input", () => {
    teacherPage = 1;
    applyTeacherView();
  });
  $("teacherRowsPerPage")?.addEventListener("change", () => {
    teacherPage = 1;
    applyTeacherView();
  });
}

function updateTeacherSortIndicators() {
  document.querySelectorAll("#teacherTable thead th.sortable").forEach((th) => {
    const key = th.getAttribute("data-key");
    const ind = th.querySelector(".sort-ind");
    if (!ind) return;
    if (key === teacherSortKey) ind.textContent = teacherSortDir === "asc" ? "▲" : "▼";
    else ind.textContent = "";
  });
}

/* =========================================================
   STUDENTS (v2: filters + date range + sort + pagination)
========================================================= */
async function loadStudents() {
  $("studentLoading")?.classList.remove("d-none");
  const res = await callAPI("getStudentData");
  $("studentLoading")?.classList.add("d-none");

  if (!res || !Array.isArray(res.rows)) {
    return Swal.fire("Network", "មិនអាចទាញទិន្នន័យ Students បានទេ។", "warning");
  }

  allStudents = res.rows;

  setupStudentFilterOptions(allStudents);

  studentPage = 1;
  studentRowsPerPage = Number($("studentRowsPerPage")?.value || 20);

  applyStudentFilters();

  setLastSync("students");
  bindStudentSortEvents();
}

function setupStudentFilterOptions(rows) {
  const teachers = new Set();
  const grades = new Set();

  rows.forEach((r) => {
    if (r[3]) teachers.add(String(r[3]).trim());
    if (r[2]) grades.add(String(r[2]).trim());
  });

  const teacherSel = $("filterTeacher");
  const gradeSel = $("filterGrade");

  if (teacherSel) {
    const list = ["ALL", ...Array.from(teachers).sort((a, b) => a.localeCompare(b, "km"))];
    teacherSel.innerHTML = list
      .map((t) => `<option value="${escapeHtml(t)}">${t === "ALL" ? "All Teachers" : t}</option>`)
      .join("");
  }

  if (gradeSel) {
    const list = ["ALL", ...Array.from(grades).sort((a, b) => a.localeCompare(b, "km"))];
    gradeSel.innerHTML = list
      .map((g) => `<option value="${escapeHtml(g)}">${g === "ALL" ? "All Grades" : g}</option>`)
      .join("");
  }
}

function applyStudentFilters() {
  studentRowsPerPage = Number($("studentRowsPerPage")?.value || 20);

  const q = ($("searchStudent")?.value || "").toLowerCase().trim();
  const teacher = $("filterTeacher")?.value || "ALL";
  const grade = $("filterGrade")?.value || "ALL";
  const gender = $("filterGender")?.value || "ALL";

  const from = $("dateFrom")?.value ? new Date($("dateFrom").value + "T00:00:00") : null;
  const to = $("dateTo")?.value ? new Date($("dateTo").value + "T23:59:59") : null;

  const mapped = allStudents.map((r, idx) => ({
    idx,
    name: String(r[0] ?? ""),
    gender: String(r[1] ?? ""),
    grade: String(r[2] ?? ""),
    teacher: String(r[3] ?? ""),
    fee: toNumber(r[4]),
    feeText: String(r[4] ?? ""),

    // ✅ FIX: Payment Date is r[8] not r[7]
    payDateRaw: r[PAYMENT_DATE_INDEX],
    payDate: parseDateAny(r[PAYMENT_DATE_INDEX]),

    raw: r,
  }));

  studentViewRows = mapped
    .filter((o) => {
      const matchQ = !q || o.name.toLowerCase().includes(q) || o.teacher.toLowerCase().includes(q);
      const matchTeacher = teacher === "ALL" || o.teacher === teacher;
      const matchGrade = grade === "ALL" || o.grade === grade;
      const matchGender = gender === "ALL" || o.gender === gender;

      // ✅ Date filter uses PaymentDate (r[8])
      let matchDate = true;
      if (from || to) {
        if (!o.payDate) matchDate = false;
        else {
          if (from && o.payDate < from) matchDate = false;
          if (to && o.payDate > to) matchDate = false;
        }
      }

      return matchQ && matchTeacher && matchGrade && matchGender && matchDate;
    })
    .sort((a, b) => compareByKey(a, b, studentSortKey, studentSortDir));

  studentPage = clampPage(studentPage, studentViewRows.length, studentRowsPerPage);

  renderStudentPage();
  renderStudentQuickStats(studentViewRows);
  updateStudentSortIndicators();
}

function clearStudentFilters() {
  $("searchStudent").value = "";
  $("filterTeacher").value = "ALL";
  $("filterGrade").value = "ALL";
  $("filterGender").value = "ALL";
  $("dateFrom").value = "";
  $("dateTo").value = "";

  studentPage = 1;
  applyStudentFilters();
}

function renderStudentQuickStats(rows) {
  const count = rows.length;
  let totalFee = 0;
  rows.forEach((o) => (totalFee += o.fee));

  const teacher80 = totalFee * 0.8;
  const school20 = totalFee * 0.2;

  $("studentStatsRow").innerHTML = `
    <div class="stat-card accent-green">
      <div class="label">សិស្ស (Filtered)</div>
      <div class="value">${count.toLocaleString("en-US")}</div>
      <div class="sub">តាម Filter/ Search/ Date</div>
    </div>

    <div class="stat-card accent-blue">
      <div class="label">ទឹកប្រាក់ (Filtered)</div>
      <div class="value" style="color:#16a34a">${formatKHR(totalFee)}</div>
      <div class="sub">សរុបតាម Filter</div>
    </div>

    <div class="stat-card accent-purple">
      <div class="label">គ្រូ 80% (Filtered)</div>
      <div class="value" style="color:#2563eb">${formatKHR(teacher80)}</div>
      <div class="sub">គណនា 80%</div>
    </div>

    <div class="stat-card accent-red">
      <div class="label">សាលា 20% (Filtered)</div>
      <div class="value" style="color:#ef4444">${formatKHR(school20)}</div>
      <div class="sub">គណនា 20%</div>
    </div>
  `;
}

function renderStudentPage() {
  const { pageItems, startIndex, endIndex, totalPages, totalItems } = paginate(
    studentViewRows,
    studentPage,
    studentRowsPerPage
  );

  const isAdmin = currentUserRole === "Admin";

  $("studentBody").innerHTML = pageItems
    .map(
      (o) => `
    <tr>
      <td class="fw-bold text-primary">${escapeHtml(o.name)}</td>
      <td class="d-none d-md-table-cell">${escapeHtml(o.gender)}</td>
      <td class="d-none d-md-table-cell">${escapeHtml(o.grade)}</td>
      <td>${escapeHtml(o.teacher)}</td>
      <td class="text-success small fw-bold">${escapeHtml(o.feeText || formatKHR(o.fee))}</td>
      <td>${escapeHtml(o.payDateRaw || "")}</td>
      <td>
        <div class="btn-group">
          <button class="btn btn-sm btn-outline-info" title="វិក្កយបត្រ" onclick="printReceipt(${o.idx})">
            <i class="bi bi-printer"></i>
          </button>
          ${
            isAdmin
              ? `
            <button class="btn btn-sm btn-outline-warning" title="កែប្រែ" onclick="editStudent(${o.idx})">
              <i class="bi bi-pencil"></i>
            </button>
            <button class="btn btn-sm btn-outline-danger" title="លុប" onclick="confirmDelete(${o.idx})">
              <i class="bi bi-trash"></i>
            </button>
          `
              : ""
          }
        </div>
      </td>
    </tr>
  `
    )
    .join("");

  $("studentPagePill").innerText = `${studentPage}/${Math.max(1, totalPages)}`;
  $("studentPageInfo").innerText = `Showing ${startIndex}-${endIndex} of ${totalItems}`;
}

function studentPrevPage() {
  studentPage = Math.max(1, studentPage - 1);
  renderStudentPage();
}
function studentNextPage() {
  const totalPages = Math.max(1, Math.ceil(studentViewRows.length / studentRowsPerPage));
  studentPage = Math.min(totalPages, studentPage + 1);
  renderStudentPage();
}

function bindStudentSortEvents() {
  const ths = document.querySelectorAll("#studentTable thead th.sortable");
  ths.forEach((th) => {
    th.onclick = () => {
      const key = th.getAttribute("data-key");
      if (!key) return;
      if (studentSortKey === key) studentSortDir = studentSortDir === "asc" ? "desc" : "asc";
      else {
        studentSortKey = key;
        studentSortDir = "asc";
      }
      studentPage = 1;
      applyStudentFilters();
    };
  });

  $("studentRowsPerPage")?.addEventListener("change", () => {
    studentPage = 1;
    applyStudentFilters();
  });

  $("searchStudent")?.addEventListener("input", () => {
    studentPage = 1;
    applyStudentFilters();
  });
  $("filterTeacher")?.addEventListener("change", () => {
    studentPage = 1;
    applyStudentFilters();
  });
  $("filterGrade")?.addEventListener("change", () => {
    studentPage = 1;
    applyStudentFilters();
  });
  $("filterGender")?.addEventListener("change", () => {
    studentPage = 1;
    applyStudentFilters();
  });
  $("dateFrom")?.addEventListener("change", () => {
    studentPage = 1;
    applyStudentFilters();
  });
  $("dateTo")?.addEventListener("change", () => {
    studentPage = 1;
    applyStudentFilters();
  });

  $("searchStudent")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      studentPage = 1;
      applyStudentFilters();
    }
  });
}

function updateStudentSortIndicators() {
  document.querySelectorAll("#studentTable thead th.sortable").forEach((th) => {
    const key = th.getAttribute("data-key");
    const ind = th.querySelector(".sort-ind");
    if (!ind) return;
    if (key === studentSortKey) ind.textContent = studentSortDir === "asc" ? "▲" : "▼";
    else ind.textContent = "";
  });
}

/* =========================================================
   Export TSV (Khmer OK in Excel)
========================================================= */
function exportTeacherTSV() {
  if (!teacherRows.length) return Swal.fire("Info", "មិនមានទិន្នន័យគ្រូសម្រាប់ Export។", "info");

  const header = ["Teacher", "Gender", "Students", "TotalFee", "Teacher80", "School20"];
  const lines = [header.join("\t")];

  teacherRows.forEach((r) => {
    const row = [
      String(r[0] ?? ""),
      String(r[1] ?? ""),
      String(r[2] ?? ""),
      String(r[3] ?? ""),
      String(r[4] ?? ""),
      String(r[5] ?? ""),
    ];
    lines.push(row.join("\t"));
  });

  downloadTSV(`Teacher_Summary_${new Date().toISOString().slice(0, 10)}.tsv`, lines.join("\n"));
}

function exportStudentTSV() {
  const rows = studentViewRows.length ? studentViewRows.map((o) => o.raw) : allStudents;
  if (!rows.length) return Swal.fire("Info", "មិនមានទិន្នន័យសិស្សសម្រាប់ Export។", "info");

  const header = ["StudentName", "Gender", "Grade", "Teacher", "Fee", "PaymentDate"];
  const lines = [header.join("\t")];

  rows.forEach((r) => {
    const row = [
      String(r[0] ?? ""),
      String(r[1] ?? ""),
      String(r[2] ?? ""),
      String(r[3] ?? ""),
      String(r[4] ?? ""),
      // ✅ FIX: PaymentDate is r[8]
      String(r[PAYMENT_DATE_INDEX] ?? ""),
    ];
    lines.push(row.join("\t"));
  });

  downloadTSV(`Students_${new Date().toISOString().slice(0, 10)}.tsv`, lines.join("\n"));
}

/* =========================================================
   Export Teacher PDF (Khmer OK)
========================================================= */
function exportTeacherPDF() {
  if (!teacherRows.length) {
    return Swal.fire("Info", "មិនមានទិន្នន័យ Teacher Summary សម្រាប់ Export PDF។", "info");
  }

  let totalTeachers = teacherRows.length;
  let totalStudents = 0;
  let totalFee = 0;

  teacherRows.forEach((r) => {
    totalStudents += toNumber(r[2]);
    totalFee += toNumber(r[3]);
  });

  const total80 = totalFee * 0.8;
  const total20 = totalFee * 0.2;

  const trs = teacherRows
    .map((r) => {
      const fee = toNumber(r[3]);
      const t80 = toNumber(r[4]) || fee * 0.8;
      const s20 = toNumber(r[5]) || fee * 0.2;
      return `
      <tr>
        <td class="left">${escapeHtml(r[0] ?? "")}</td>
        <td class="center">${escapeHtml(r[1] ?? "")}</td>
        <td class="center">${toNumber(r[2])}</td>
        <td class="right">${formatKHR(fee)}</td>
        <td class="right blue">${formatKHR(t80)}</td>
        <td class="right red">${formatKHR(s20)}</td>
      </tr>
    `;
    })
    .join("");

  const printWindow = window.open("", "", "height=900,width=1100");
  const html = `
  <html><head><title>Teacher Summary PDF</title>
  <style>
    @page{size:A4 portrait;margin:12mm;}
    body{font-family:'Noto Serif Khmer','Khmer OS Siemreap',sans-serif;color:#000;}
    .header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px;}
    .logoBox{width:70px;text-align:center;}
    .logoBox img{width:70px;}
    .rightHeader{text-align:center;font-family:'Khmer OS Muol Light','Noto Serif Khmer',serif;font-size:14px;line-height:1.7;}
    .docTitle{text-align:center;margin:10px 0 14px;font-family:'Khmer OS Muol Light','Noto Serif Khmer',serif;font-size:18px;text-decoration:underline;}
    .stats{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:12px;}
    .stat{border:1px solid #000;border-radius:6px;padding:8px;text-align:center;}
    .stat .label{font-size:11px;font-weight:700;}
    .stat .value{font-size:13px;font-weight:800;margin-top:2px;}
    table{width:100%;border-collapse:collapse;font-size:12px;}
    th,td{border:1px solid #000;padding:8px;}
    th{background:#f2f2f2;font-weight:800;text-align:center;}
    td.left{text-align:left;}
    td.center{text-align:center;}
    td.right{text-align:right;font-weight:700;}
    .blue{color:#0d6efd;}
    .red{color:#dc3545;}
    .footer{margin-top:14px;display:flex;justify-content:space-between;padding:0 50px;}
    .sig{width:240px;text-align:center;}
    .sig .role{font-family:'Khmer OS Muol Light','Noto Serif Khmer',serif;font-size:13px;margin-bottom:70px;}
    .sig .line{border-bottom:1px dotted #000;}
    .sig .name{margin-top:12px;font-weight:800;}
    *{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;}
  </style>
  </head><body>
    <div class="header">
      <div class="logoBox">
        <img src="https://blogger.googleusercontent.com/img/a/AVvXsEi33gP-LjadWAMAbW6z8mKj7NUYkZeslEJ4sVFw7WK3o9fQ-JTQFMWEe06xxew4lj7WKpfuk8fadTm5kXo3GSW9jNaQHE8SrCs8_bUFDV8y4TOJ1Zhbu0YKVnWIgL7sTPuEPMrmrtuNqwDPWKHOvy6PStAaSrCz-GpLfsQNyq-BAElq9EI3etjnYsft0Pvo" />
        <div style="font-size:11px;margin-top:4px;">សាលាបឋមសិក្សាសម្តេចព្រះរាជអគ្គមហេសី</div>
      </div>
      <div class="rightHeader">ព្រះរាជាណាចក្រកម្ពុជា<br/>ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
    </div>

    <div class="docTitle">របាយការណ៍សង្ខេបគ្រូបង្រៀន (Teacher Summary)</div>

    <div class="stats">
      <div class="stat"><div class="label">គ្រូសរុប</div><div class="value">${totalTeachers}</div></div>
      <div class="stat"><div class="label">សិស្សសរុប</div><div class="value">${totalStudents}</div></div>
      <div class="stat"><div class="label">ទឹកប្រាក់សរុប</div><div class="value">${formatKHR(totalFee)}</div></div>
      <div class="stat"><div class="label">គ្រូ 80% / សាលា 20%</div><div class="value"><span class="blue">${formatKHR(total80)}</span> / <span class="red">${formatKHR(total20)}</span></div></div>
    </div>

    <table>
      <thead>
        <tr>
          <th style="width:34%;">គ្រូ</th>
          <th style="width:12%;">ភេទ</th>
          <th style="width:10%;">សិស្ស</th>
          <th style="width:14%;">តម្លៃសរុប</th>
          <th style="width:15%;">៨០%</th>
          <th style="width:15%;">២០%</th>
        </tr>
      </thead>
      <tbody>${trs}</tbody>
    </table>

    <div style="text-align:right;margin-top:10px;">ថ្ងៃទី........ខែ........ឆ្នាំ២០២៦</div>

    <div class="footer">
      <div class="sig"><div class="role">បានពិនិត្យ និងឯកភាព<br/>នាយកសាលា</div><div class="line"></div></div>
      <div class="sig"><div class="role">អ្នកចេញវិក្កយបត្រ</div><div class="name">ហម ម៉ាលីនដា</div></div>
    </div>

    <script>
      window.onload=function(){ window.print(); setTimeout(function(){window.close()},600); }
    </script>
  </body></html>`;

  printWindow.document.write(html);
  printWindow.document.close();
}

/* =========================================================
   Print Student Report Detailed (uses current filtered view)
========================================================= */
function printStudentReportDetailed() {
  const rows = studentViewRows.length ? studentViewRows.map((o) => o.raw) : allStudents;

  const printWindow = window.open("", "", "height=900,width=1100");
  const totalStudents = rows.length;
  const totalFemale = rows.filter((s) => s[1] === "Female" || s[1] === "ស្រី").length;

  let totalFee = 0;

  const tableRows = rows
    .map((r) => {
      const feeNum = toNumber(r[4]);
      totalFee += feeNum;

      const teacherPart = feeNum * 0.8;
      const schoolPart = feeNum * 0.2;

      // ✅ FIX: PaymentDate is r[8]
      let payDate = r[PAYMENT_DATE_INDEX];
      if (!payDate || String(payDate).includes("KHR")) payDate = new Date().toLocaleDateString("km-KH");

      return `
      <tr>
        <td style="border:1px solid #000;padding:6px;text-align:left;">${escapeHtml(r[0] ?? "")}</td>
        <td style="border:1px solid #000;padding:6px;text-align:center;">${escapeHtml(r[1] ?? "")}</td>
        <td style="border:1px solid #000;padding:6px;text-align:center;">${escapeHtml(r[2] ?? "")}</td>
        <td style="border:1px solid #000;padding:6px;text-align:left;">${escapeHtml(r[3] ?? "")}</td>
        <td style="border:1px solid #000;padding:6px;text-align:right;font-weight:bold;">${feeNum.toLocaleString()} ៛</td>
        <td style="border:1px solid #000;padding:6px;text-align:right;color:#0d6efd;">${teacherPart.toLocaleString()} ៛</td>
        <td style="border:1px solid #000;padding:6px;text-align:right;color:#dc3545;">${schoolPart.toLocaleString()} ៛</td>
        <td style="border:1px solid #000;padding:6px;text-align:center;">${escapeHtml(payDate)}</td>
      </tr>
    `;
    })
    .join("");

  const fee80 = totalFee * 0.8;
  const fee20 = totalFee * 0.2;

  const reportHTML = `
  <html><head><title>Student Report Detailed</title>
  <style>
    body{font-family:'Khmer OS Siemreap','Noto Serif Khmer',sans-serif;padding:20px;color:#000;background:#fff;}
    .header-wrapper{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:22px;}
    .left-header{text-align:center;}
    .right-header{text-align:center;font-family:'Khmer OS Muol Light','Noto Serif Khmer',serif;font-size:14px;line-height:1.7;}
    .logo-box{width:70px;margin:0 auto 5px;}
    .logo-box img{width:100%;display:block;}
    .school-kh{font-family:'Khmer OS Muol Light','Noto Serif Khmer',serif;font-size:14px;line-height:1.8;}
    .report-title{text-align:center;font-family:'Khmer OS Muol Light','Noto Serif Khmer',serif;font-size:18px;text-decoration:underline;margin:0 0 14px;}
    .stats{display:grid;grid-template-columns:repeat(5,1fr);gap:8px;margin-bottom:16px;}
    .stat{border:1px solid #000;padding:6px;text-align:center;border-radius:4px;}
    .stat .label{font-size:10px;font-weight:700;}
    .stat .value{font-size:12px;font-weight:800;margin-top:2px;}
    table{width:100%;border-collapse:collapse;font-size:12px;}
    th{border:1px solid #000;padding:8px;background:#f2f2f2;}
    td{border:1px solid #000;padding:6px;}
    .date-section{text-align:right;font-size:13px;margin-top:14px;padding-right:60px;}
    .signature-wrapper{display:flex;justify-content:space-between;padding:0 80px;margin-top:18px;}
    .sig-box{text-align:center;width:220px;}
    .sig-role{font-family:'Khmer OS Muol Light','Noto Serif Khmer',serif;font-size:13px;margin-bottom:60px;}
    .sig-line{border-bottom:1px dotted #000;width:100%;margin-top:30px;}
    .sig-name{font-weight:800;font-size:13px;margin-top:10px;}
    @media print{@page{size:A4 landscape;margin:1cm;}}
  </style></head><body>
    <div class="header-wrapper">
      <div class="left-header">
        <div class="logo-box">
          <img src="https://blogger.googleusercontent.com/img/a/AVvXsEi33gP-LjadWAMAbW6z8mKj7NUYkZeslEJ4sVFw7WK3o9fQ-JTQFMWEe06xxew4lj7WKpfuk8fadTm5kXo3GSW9jNaQHE8SrCs8_bUFDV8y4TOJ1Zhbu0YKVnWIgL7sTPuEPMrmrtuNqwDPWKHOvy6PStAaSrCz-GpLfsQNyq-BAElq9EI3etjnYsft0Pvo" alt="Logo"/>
        </div>
        <div class="school-kh">សាលាបឋមសិក្សាសម្តេចព្រះរាជអគ្គមហេសី<br/>នរោត្តមមុនីនាថសីហនុ</div>
      </div>
      <div class="right-header">ព្រះរាជាណាចក្រកម្ពុជា<br/>ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
    </div>

    <div class="report-title">របាយការណ៍លម្អិតសិស្សរៀនបំប៉នបន្ថែម</div>

    <div class="stats">
      <div class="stat"><div class="label">សិស្សសរុប</div><div class="value">${totalStudents} នាក់</div></div>
      <div class="stat"><div class="label">សរុបស្រី</div><div class="value">${totalFemale} នាក់</div></div>
      <div class="stat"><div class="label">ទឹកប្រាក់សរុប</div><div class="value">${totalFee.toLocaleString()} ៛</div></div>
      <div class="stat"><div class="label">គ្រូ (80%)</div><div class="value">${fee80.toLocaleString()} ៛</div></div>
      <div class="stat"><div class="label">សាលា (20%)</div><div class="value">${fee20.toLocaleString()} ៛</div></div>
    </div>

    <table>
      <thead>
        <tr>
          <th style="width:18%;">ឈ្មោះសិស្ស</th>
          <th style="width:7%;">ភេទ</th>
          <th style="width:8%;">ថ្នាក់</th>
          <th style="width:15%;">គ្រូបង្រៀន</th>
          <th style="width:13%;">តម្លៃសិក្សា</th>
          <th style="width:13%;">គ្រូ (80%)</th>
          <th style="width:13%;">សាលា (20%)</th>
          <th style="width:13%;">ថ្ងៃបង់ប្រាក់</th>
        </tr>
      </thead>
      <tbody>${tableRows}</tbody>
    </table>

    <div class="date-section">ថ្ងៃទី........ខែ........ឆ្នាំ២០២៦</div>

    <div class="signature-wrapper">
      <div class="sig-box">
        <div class="sig-role">បានពិនិត្យ និងឯកភាព<br/>នាយកសាលា</div>
        <div class="sig-line"></div>
      </div>
      <div class="sig-box">
        <div class="sig-role">អ្នកចេញវិក្កយបត្រ</div>
        <div class="sig-name">ហម ម៉ាលីនដា</div>
      </div>
    </div>

    <script>
      window.onload=function(){ window.print(); setTimeout(function(){window.close()},600); }
    </script>
  </body></html>`;

  printWindow.document.write(reportHTML);
  printWindow.document.close();
}

/* =========================================================
   Receipt
========================================================= */
function printReceipt(index) {
  const s = allStudents[index];
  if (!s) return;

  const printWindow = window.open("", "", "height=600,width=800");
  const receiptHTML = `
    <html><head>
      <title>Receipt - ${escapeHtml(s[0] ?? "")}</title>
      <link href="https://fonts.googleapis.com/css2?family=Noto+Serif+Khmer:wght@400;700&display=swap" rel="stylesheet">
      <style>
        body{font-family:'Noto Serif Khmer',serif;padding:40px;text-align:center;}
        .receipt-box{border:2px solid #333;padding:30px;width:420px;margin:auto;border-radius:12px;}
        .header{font-weight:800;font-size:20px;margin-bottom:5px;color:#4361ee;}
        .line{border-bottom:2px dashed #ccc;margin:15px 0;}
        .details{text-align:left;font-size:15px;line-height:1.85;}
        .footer{margin-top:22px;font-size:12px;font-style:italic;color:#666;}
        .price{font-size:18px;color:#10b981;font-weight:800;}
      </style>
    </head><body>
      <div class="receipt-box">
        <div class="header">វិក្កយបត្របង់ប្រាក់</div>
        <div style="font-size:14px;">សាលារៀន ព្រះរាជអគ្គមហេសី</div>
        <div class="line"></div>
        <div class="details">
          <div>ឈ្មោះសិស្ស: <b>${escapeHtml(s[0] ?? "")}</b></div>
          <div>ភេទ: <b>${escapeHtml(s[1] ?? "")}</b></div>
          <div>ថ្នាក់សិក្សា: <b>${escapeHtml(s[2] ?? "")}</b></div>
          <div>គ្រូបង្រៀន: <b>${escapeHtml(s[3] ?? "")}</b></div>
          <div>តម្លៃសិក្សា: <span class="price">${escapeHtml(s[4] ?? "")}</span></div>
          <div>កាលបរិច្ឆេទ: <b>${new Date().toLocaleDateString("km-KH")}</b></div>
        </div>
        <div class="line"></div>
        <div class="footer">សូមអរគុណ! ការអប់រំគឺជាទ្រព្យសម្បត្តិដែលមិនអាចកាត់ថ្លៃបាន។</div>
      </div>
      <script>window.onload=function(){window.print();window.close();}</script>
    </body></html>
  `;
  printWindow.document.write(receiptHTML);
  printWindow.document.close();
}

/* =========================================================
   CRUD (Admin) + Fee split preview
========================================================= */
function updateFeeSplitPreview() {
  const fee = toNumber($("addFee")?.value);
  $("disp80").innerText = formatKHR(fee * 0.8);
  $("disp20").innerText = formatKHR(fee * 0.2);
}

function openStudentModal() {
  isEditMode = false;
  originalName = "";

  $("modalTitle").innerText = "បញ្ចូលសិស្សថ្មី";
  $("addStudentName").value = "";
  $("addGender").value = "Male";
  $("addGrade").value = "";
  $("addFee").value = "";
  updateFeeSplitPreview();

  bootstrap.Modal.getOrCreateInstance($("studentModal")).show();
}

function editStudent(index) {
  isEditMode = true;
  const r = allStudents[index];
  originalName = r?.[0] ?? "";

  $("modalTitle").innerText = "កែប្រែព័ត៌មាន";
  $("addStudentName").value = r?.[0] ?? "";
  $("addGender").value = r?.[1] ?? "Male";
  $("addGrade").value = r?.[2] ?? "";
  $("addTeacherSelect").value = r?.[3] ?? "";

  const feeValue = String(r?.[4] ?? "").replace(/[^0-9]/g, "");
  $("addFee").value = feeValue;
  updateFeeSplitPreview();

  bootstrap.Modal.getOrCreateInstance($("studentModal")).show();
}

async function submitStudent() {
  if (currentUserRole !== "Admin") return;

  const name = $("addStudentName").value.trim();
  const teacher = $("addTeacherSelect").value;
  const feeNum = toNumber($("addFee").value);

  if (!name || !teacher) {
    return Swal.fire("Error", "សូមបំពេញឈ្មោះសិស្ស និងជ្រើសរើសគ្រូ", "error");
  }

  const form = {
    studentName: name,
    gender: $("addGender").value,
    grade: $("addGrade").value,
    teacherName: teacher,
    schoolFee: formatKHR(feeNum),
    teacherFeeVal: formatKHR(feeNum * 0.8),
    schoolFeeVal: formatKHR(feeNum * 0.2),
    paymentDate: new Date().toISOString().split("T")[0],
    startDate: new Date().toISOString().split("T")[0],
  };

  Swal.fire({ title: "កំពុងរក្សាទុក...", didOpen: () => Swal.showLoading(), allowOutsideClick: false });

  const res = isEditMode
    ? await callAPI("updateStudentData", originalName, form)
    : await callAPI("saveStudentToTeacherSheet", form);

  if (res && res.success) {
    Swal.fire("ជោគជ័យ", res.message || "រក្សាទុកបានសម្រេច", "success");
    bootstrap.Modal.getOrCreateInstance($("studentModal")).hide();
    await refreshAll();
  } else {
    Swal.fire("Error", res?.message || "រក្សាទុកមិនបានសម្រេច", "error");
  }
}

async function confirmDelete(index) {
  if (currentUserRole !== "Admin") return;

  const name = allStudents[index]?.[0] || "";
  const teacher = allStudents[index]?.[3] || "";

  Swal.fire({
    title: "លុបទិន្នន័យ?",
    text: `តើអ្នកចង់លុបសិស្ស ${name}?`,
    icon: "warning",
    showCancelButton: true,
    confirmButtonColor: "#ef4444",
    confirmButtonText: "បាទ លុបវា!",
    cancelButtonText: "បោះបង់",
  }).then(async (result) => {
    if (!result.isConfirmed) return;

    Swal.fire({ title: "កំពុងលុប...", didOpen: () => Swal.showLoading() });
    const res = await callAPI("deleteStudentData", name, teacher);

    if (res && res.success) {
      Swal.fire("Deleted!", res.message || "លុបបានសម្រេច", "success");
      await refreshAll();
    } else {
      Swal.fire("Error", res?.message || "លុបមិនបានសម្រេច", "error");
    }
  });
}

/* =========================================================
   Core utilities (paginate + sort compare + escape)
========================================================= */
function paginate(items, page, perPage) {
  const totalItems = items.length;
  const totalPages = Math.max(1, Math.ceil(totalItems / perPage));
  const p = Math.min(Math.max(1, page), totalPages);

  const start = (p - 1) * perPage;
  const end = Math.min(start + perPage, totalItems);
  const pageItems = items.slice(start, end);

  return {
    pageItems,
    startIndex: totalItems ? start + 1 : 0,
    endIndex: end,
    totalPages,
    totalItems,
  };
}

function clampPage(page, totalItems, perPage) {
  const totalPages = Math.max(1, Math.ceil(totalItems / perPage));
  return Math.min(Math.max(1, page), totalPages);
}

function compareByKey(a, b, key, dir) {
  const mul = dir === "asc" ? 1 : -1;

  const va = a[key];
  const vb = b[key];

  // date
  if (va instanceof Date || vb instanceof Date) {
    const ta = va instanceof Date ? va.getTime() : -Infinity;
    const tb = vb instanceof Date ? vb.getTime() : -Infinity;
    return (ta - tb) * mul;
  }

  // number
  if (typeof va === "number" || typeof vb === "number") {
    const na = Number(va) || 0;
    const nb = Number(vb) || 0;
    return (na - nb) * mul;
  }

  // string
  return (
    String(va ?? "").localeCompare(String(vb ?? ""), "km", { sensitivity: "base" }) * mul
  );
}

function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

/* ---------------- Init ---------------- */
document.addEventListener("DOMContentLoaded", () => {
  $("addFee")?.addEventListener("input", updateFeeSplitPreview);
});



