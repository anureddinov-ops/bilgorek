(function () {
"use strict";
/* =====================================================================
   BilGörək — Admin Panel məntiqi (admin.js)
   db.js (BilGorekDB) və lib/xlsx.full.min.js (SheetJS, lokal) istifadə edir.
   ===================================================================== */
"use strict";

const PAGE_SIZE = 15;
const DIFFICULTIES = BilGorekDB.DIFFICULTIES; // ["Asan","Orta","Çətin"]

/* ======================= ÜMUMİ KÖMƏKÇİLƏR ======================= */
function escapeHtml(str) {
  const d = document.createElement("div");
  d.textContent = str === undefined || str === null ? "" : String(str);
  return d.innerHTML;
}

let toastTimer = null;
function showToast(msg, isError) {
  const t = document.getElementById("admin-toast") || (() => {
    const el = document.createElement("div");
    el.id = "admin-toast"; el.className = "toast";
    const host = document.getElementById("view-admin") || document.body;
    host.appendChild(el);
    return el;
  })();
  t.textContent = msg;
  t.classList.toggle("error", !!isError);
  t.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove("show"), 3200);
}

function downloadBlob(filename, content, mime) {
  const blob = (content instanceof Blob) ? content : new Blob([content], { type: mime || "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

function formatDate(iso) {
  try {
    const d = new Date(iso);
    return d.toLocaleDateString("az-AZ") + " " + d.toLocaleTimeString("az-AZ", { hour: "2-digit", minute: "2-digit" });
  } catch (e) { return iso || ""; }
}

/* Generic confirm modal (Promise<boolean>) */
function confirmDialog(title, message, okText) {
  const overlay = document.getElementById("confirm-modal-overlay");
  document.getElementById("confirm-title").textContent = title;
  document.getElementById("confirm-message").textContent = message;
  const okBtn = document.getElementById("confirm-ok");
  okBtn.textContent = okText || "Təsdiqlə";
  overlay.style.display = "flex";
  return new Promise((resolve) => {
    function cleanup(result) {
      overlay.style.display = "none";
      okBtn.removeEventListener("click", onOk);
      document.getElementById("confirm-cancel").removeEventListener("click", onCancel);
      resolve(result);
    }
    function onOk() { cleanup(true); }
    function onCancel() { cleanup(false); }
    okBtn.addEventListener("click", onOk);
    document.getElementById("confirm-cancel").addEventListener("click", onCancel);
  });
}

/* ======================= CSV (öz həyata keçirməmiz — SheetJS-dən asılı deyil) ======================= */
function csvParse(text) {
  const rows = [];
  let row = [], field = "", inQuotes = false;
  // BOM təmizlə
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else { inQuotes = false; }
      } else field += c;
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ',') { row.push(field); field = ""; }
      else if (c === '\r') { /* skip */ }
      else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ""; }
      else field += c;
    }
  }
  if (field.length > 0 || row.length > 0) { row.push(field); rows.push(row); }
  return rows.filter(r => !(r.length === 1 && r[0] === ""));
}

function csvField(val) {
  const s = (val === undefined || val === null) ? "" : String(val);
  if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function csvStringify(headers, rows) {
  const lines = [headers.map(csvField).join(",")];
  rows.forEach(r => lines.push(headers.map(h => csvField(r[h])).join(",")));
  return "﻿" + lines.join("\r\n");
}

/* ======================= SUAL SIRA <-> DB FORMATI ÇEVİRMƏSİ =======================
   İdxal/İxrac zamanı "kateqoriya" sütunu İNSANA OXUNAQLI AD (məs. "Azərbaycan") olur,
   DB-də isə daxili AÇAR (məs. "azerbaycan") saxlanılır. */
const ROW_COLUMNS = ["id", "sual", "cavabA", "cavabB", "cavabC", "cavabD", "duzgunCavab", "kateqoriya", "cetinlik", "izah", "ipucu", "aktiv"];

function dbRowToExportRow(dbRow, categoriesByKey) {
  const cat = categoriesByKey[dbRow.kateqoriya];
  return {
    id: dbRow.id,
    sual: dbRow.sual,
    cavabA: dbRow.cavabA, cavabB: dbRow.cavabB, cavabC: dbRow.cavabC, cavabD: dbRow.cavabD,
    duzgunCavab: dbRow.duzgunCavab,
    kateqoriya: cat ? cat.name : dbRow.kateqoriya,
    cetinlik: dbRow.cetinlik,
    izah: dbRow.izah || "",
    ipucu: dbRow.ipucu || "",
    aktiv: dbRow.aktiv ? "true" : "false"
  };
}

function resolveCategoryInput(value, categoriesCache) {
  if (!value) return null;
  const v = String(value).trim().toLowerCase();
  const byKey = categoriesCache.find(c => c.key.toLowerCase() === v);
  if (byKey) return byKey.key;
  const byName = categoriesCache.find(c => c.name.toLowerCase() === v);
  if (byName) return byName.key;
  return null;
}

function parseBoolLike(value, defaultVal) {
  if (value === undefined || value === null || value === "") return defaultVal;
  const v = String(value).trim().toLowerCase();
  if (["true", "1", "bəli", "beli", "aktiv", "yes"].includes(v)) return true;
  if (["false", "0", "xeyr", "passiv", "no"].includes(v)) return false;
  return null; // yanlış qiymət
}

/* ======================= TAB NAVİQASİYASI ======================= */
document.querySelectorAll(".tab-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
  });
});

/* ======================= KATEQORİYALAR ======================= */
let categoriesCache = [];
let categoriesByKeyCache = {};

async function loadCategoriesCache() {
  categoriesCache = await BilGorekDB.getCategories();
  categoriesByKeyCache = {};
  categoriesCache.forEach(c => { categoriesByKeyCache[c.key] = c; });
  return categoriesCache;
}

function fillCategorySelect(selectEl, includeEmpty, emptyLabel) {
  selectEl.innerHTML = "";
  if (includeEmpty) {
    const opt = document.createElement("option");
    opt.value = ""; opt.textContent = emptyLabel || "Bütün kateqoriyalar";
    selectEl.appendChild(opt);
  }
  categoriesCache.forEach(c => {
    const opt = document.createElement("option");
    opt.value = c.key; opt.textContent = c.icon + " " + c.name;
    selectEl.appendChild(opt);
  });
}

function fillDifficultySelect(selectEl, includeEmpty) {
  selectEl.innerHTML = "";
  if (includeEmpty) {
    const opt = document.createElement("option");
    opt.value = ""; opt.textContent = "Bütün çətinliklər";
    selectEl.appendChild(opt);
  }
  DIFFICULTIES.forEach(d => {
    const opt = document.createElement("option");
    opt.value = d; opt.textContent = d;
    selectEl.appendChild(opt);
  });
}

async function renderCategoriesTab() {
  const list = document.getElementById("category-list");
  list.innerHTML = `<div class="loading-msg">Yüklənir...</div>`;
  await loadCategoriesCache();
  const allQuestions = await BilGorekDB.getAllQuestions();
  list.innerHTML = "";
  if (categoriesCache.length === 0) {
    list.innerHTML = `<div class="empty-msg">Hələ kateqoriya yoxdur.</div>`;
    return;
  }
  categoriesCache.forEach(cat => {
    const count = allQuestions.filter(q => q.kateqoriya === cat.key).length;
    const item = document.createElement("div");
    item.className = "category-item";
    item.innerHTML = `
      <div class="ci-top">
        <span class="ci-icon">${escapeHtml(cat.icon)}</span>
        <span class="ci-name">${escapeHtml(cat.name)}</span>
      </div>
      <span class="ci-count">${count} sual</span>
      <div class="ci-actions">
        <button class="btn btn-secondary btn-sm" data-action="edit">✏️ Redaktə</button>
        <button class="btn btn-danger btn-sm" data-action="delete">🗑️ Sil</button>
      </div>
    `;
    item.querySelector('[data-action="edit"]').addEventListener("click", () => openCategoryModal(cat));
    item.querySelector('[data-action="delete"]').addEventListener("click", () => deleteCategoryFlow(cat, count));
    list.appendChild(item);
  });
}

let editingCategoryKey = null;
function openCategoryModal(existing) {
  editingCategoryKey = existing ? existing.key : null;
  document.getElementById("category-modal-title").textContent = existing ? "Kateqoriyanı redaktə et" : "Yeni kateqoriya";
  document.getElementById("cf-name").value = existing ? existing.name : "";
  document.getElementById("cf-icon").value = existing ? existing.icon : "📚";
  document.getElementById("cf-name").closest(".field").classList.remove("has-error");
  document.getElementById("category-modal-overlay").style.display = "flex";
}
function closeCategoryModal() { document.getElementById("category-modal-overlay").style.display = "none"; }

document.getElementById("btn-add-category").addEventListener("click", () => openCategoryModal(null));
document.getElementById("cf-cancel").addEventListener("click", closeCategoryModal);
document.getElementById("cf-save").addEventListener("click", async () => {
  const name = document.getElementById("cf-name").value.trim();
  const icon = document.getElementById("cf-icon").value.trim() || "📚";
  const nameField = document.getElementById("cf-name").closest(".field");
  if (!name) { nameField.classList.add("has-error"); return; }
  nameField.classList.remove("has-error");
  try {
    if (editingCategoryKey) {
      await BilGorekDB.updateCategory(editingCategoryKey, { name, icon });
      showToast("Kateqoriya yeniləndi ✅");
    } else {
      await BilGorekDB.addCategory({ name, icon });
      showToast("Kateqoriya əlavə edildi ✅");
    }
    closeCategoryModal();
    await renderCategoriesTab();
    await refreshFilterDropdowns();
  } catch (e) {
    showToast("Xəta: " + e.message, true);
  }
});

async function deleteCategoryFlow(cat, questionCount) {
  const overlay = document.getElementById("category-delete-modal-overlay");
  const msgEl = document.getElementById("cd-message");
  const modeField = document.getElementById("cd-mode-field");
  const modeReassign = document.getElementById("cd-mode-reassign");
  const modeCascade = document.getElementById("cd-mode-cascade");
  const reassignField = document.getElementById("cd-reassign-field");
  const reassignSelect = document.getElementById("cd-reassign-select");
  const hasOtherCategory = categoriesCache.filter(c => c.key !== cat.key).length > 0;

  function syncReassignVisibility() {
    reassignField.style.display = (questionCount > 0 && modeReassign.checked) ? "flex" : "none";
  }

  if (questionCount > 0) {
    msgEl.textContent = `"${cat.name}" kateqoriyasında ${questionCount} sual var. Necə silmək istəyirsən?`;
    modeField.style.display = "block";
    reassignSelect.innerHTML = "";
    categoriesCache.filter(c => c.key !== cat.key).forEach(c => {
      const opt = document.createElement("option");
      opt.value = c.key; opt.textContent = c.icon + " " + c.name;
      reassignSelect.appendChild(opt);
    });
    if (hasOtherCategory) {
      modeReassign.checked = true;
      modeReassign.disabled = false;
    } else {
      modeCascade.checked = true;
      modeReassign.disabled = true;
    }
    syncReassignVisibility();
  } else {
    msgEl.textContent = `"${cat.name}" kateqoriyasını silmək istədiyinizə əminsiniz?`;
    modeField.style.display = "none";
    reassignField.style.display = "none";
  }

  overlay.style.display = "flex";

  return new Promise((resolve) => {
    function cleanup() {
      overlay.style.display = "none";
      confirmBtn.removeEventListener("click", onConfirm);
      cancelBtn.removeEventListener("click", onCancel);
      modeReassign.removeEventListener("change", syncReassignVisibility);
      modeCascade.removeEventListener("change", syncReassignVisibility);
    }
    const confirmBtn = document.getElementById("cd-confirm");
    const cancelBtn = document.getElementById("cd-cancel");
    modeReassign.addEventListener("change", syncReassignVisibility);
    modeCascade.addEventListener("change", syncReassignVisibility);
    async function onConfirm() {
      try {
        if (questionCount > 0 && modeCascade.checked) {
          const allQ = await BilGorekDB.getAllQuestions();
          const ids = allQ.filter(q => q.kateqoriya === cat.key).map(q => q.id);
          if (ids.length) await BilGorekDB.deleteQuestions(ids);
          await BilGorekDB.deleteCategory(cat.key);
        } else {
          const reassignTo = questionCount > 0 ? reassignSelect.value : undefined;
          await BilGorekDB.deleteCategory(cat.key, reassignTo);
        }
        showToast("Kateqoriya silindi 🗑️");
        cleanup();
        await renderCategoriesTab();
        await refreshFilterDropdowns();
        await loadQuestions();
        resolve(true);
      } catch (e) {
        showToast("Xəta: " + e.message, true);
        cleanup();
        resolve(false);
      }
    }
    function onCancel() { cleanup(); resolve(false); }
    confirmBtn.addEventListener("click", onConfirm);
    cancelBtn.addEventListener("click", onCancel);
  });
}

/* ======================= SUALLAR CƏDVƏLİ ======================= */
let allQuestionsCache = [];
let filteredQuestions = [];
let selectedIds = new Set();
let currentPage = 1;

async function loadQuestions() {
  document.getElementById("questions-tbody").innerHTML = `<tr><td colspan="8" class="loading-msg">Yüklənir...</td></tr>`;
  allQuestionsCache = await BilGorekDB.getAllQuestions();
  applyFiltersAndRender();
}

function applyFiltersAndRender() {
  const search = document.getElementById("search-input").value.trim().toLowerCase();
  const catFilter = document.getElementById("filter-category").value;
  const diffFilter = document.getElementById("filter-difficulty").value;
  const statusFilter = document.getElementById("filter-status").value;

  filteredQuestions = allQuestionsCache.filter(q => {
    if (search && !q.sual.toLowerCase().includes(search)) return false;
    if (catFilter && q.kateqoriya !== catFilter) return false;
    if (diffFilter && q.cetinlik !== diffFilter) return false;
    if (statusFilter === "active" && !q.aktiv) return false;
    if (statusFilter === "inactive" && q.aktiv) return false;
    return true;
  });
  filteredQuestions.sort((a, b) => b.id - a.id);
  currentPage = 1;
  renderQuestionsTable();
}

function renderQuestionsTable() {
  const tbody = document.getElementById("questions-tbody");
  tbody.innerHTML = "";

  if (filteredQuestions.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" class="empty-msg">Heç bir sual tapılmadı.</td></tr>`;
    renderPagination();
    updateBulkBar();
    return;
  }

  const totalPages = Math.max(1, Math.ceil(filteredQuestions.length / PAGE_SIZE));
  if (currentPage > totalPages) currentPage = totalPages;
  const start = (currentPage - 1) * PAGE_SIZE;
  const pageItems = filteredQuestions.slice(start, start + PAGE_SIZE);

  pageItems.forEach(q => {
    const cat = categoriesByKeyCache[q.kateqoriya];
    const tr = document.createElement("tr");
    const shortText = q.sual.length > 90 ? q.sual.slice(0, 90) + "…" : q.sual;
    tr.innerHTML = `
      <td><input type="checkbox" class="row-check" data-id="${q.id}" ${selectedIds.has(q.id) ? "checked" : ""}></td>
      <td class="nowrap">#${q.id}</td>
      <td class="q-text">${escapeHtml(shortText)}</td>
      <td class="nowrap">${cat ? escapeHtml(cat.icon + " " + cat.name) : escapeHtml(q.kateqoriya)}</td>
      <td class="nowrap"><span class="diff-chip">${escapeHtml(q.cetinlik)}</span></td>
      <td class="nowrap">${escapeHtml(q.duzgunCavab)}</td>
      <td class="nowrap"><span class="status-chip ${q.aktiv ? "active" : "inactive"}">${q.aktiv ? "Aktiv" : "Passiv"}</span></td>
      <td class="nowrap">
        <div class="row-actions">
          <button class="icon-action" data-action="edit" title="Redaktə et">✏️</button>
          <button class="icon-action danger" data-action="delete" title="Sil">🗑️</button>
        </div>
      </td>
    `;
    tr.querySelector('[data-action="edit"]').addEventListener("click", () => openQuestionModal(q));
    tr.querySelector('[data-action="delete"]').addEventListener("click", () => deleteQuestionFlow(q));
    tr.querySelector(".row-check").addEventListener("change", (e) => {
      if (e.target.checked) selectedIds.add(q.id); else selectedIds.delete(q.id);
      updateBulkBar();
    });
    tbody.appendChild(tr);
  });

  renderPagination(totalPages);
  updateBulkBar();
  updateSelectAllState(pageItems);
}

function updateSelectAllState(pageItems) {
  const selectAll = document.getElementById("select-all");
  if (pageItems.length === 0) { selectAll.checked = false; return; }
  selectAll.checked = pageItems.every(q => selectedIds.has(q.id));
}

document.getElementById("select-all").addEventListener("change", (e) => {
  const start = (currentPage - 1) * PAGE_SIZE;
  const pageItems = filteredQuestions.slice(start, start + PAGE_SIZE);
  pageItems.forEach(q => { if (e.target.checked) selectedIds.add(q.id); else selectedIds.delete(q.id); });
  renderQuestionsTable();
});

function renderPagination(totalPages) {
  const el = document.getElementById("pagination");
  totalPages = totalPages || Math.max(1, Math.ceil(filteredQuestions.length / PAGE_SIZE));
  if (totalPages <= 1) { el.innerHTML = ""; return; }
  el.innerHTML = `
    <button class="btn btn-secondary btn-sm" id="pg-prev" ${currentPage <= 1 ? "disabled" : ""}>‹ Əvvəlki</button>
    <span>Səhifə ${currentPage} / ${totalPages}</span>
    <button class="btn btn-secondary btn-sm" id="pg-next" ${currentPage >= totalPages ? "disabled" : ""}>Növbəti ›</button>
  `;
  document.getElementById("pg-prev").addEventListener("click", () => { currentPage--; renderQuestionsTable(); });
  document.getElementById("pg-next").addEventListener("click", () => { currentPage++; renderQuestionsTable(); });
}

function updateBulkBar() {
  const bar = document.getElementById("bulk-bar");
  const count = selectedIds.size;
  if (count === 0) { bar.style.display = "none"; return; }
  bar.style.display = "flex";
  document.getElementById("bulk-count").textContent = count + " seçildi";
}

document.getElementById("bulk-activate").addEventListener("click", () => bulkSetActive(true));
document.getElementById("bulk-deactivate").addEventListener("click", () => bulkSetActive(false));
async function bulkSetActive(active) {
  await BilGorekDB.bulkSetActive(Array.from(selectedIds), active);
  showToast((active ? "Aktivləşdirildi" : "Passivləşdirildi") + " ✅ (" + selectedIds.size + " sual)");
  selectedIds.clear();
  await loadQuestions();
}
document.getElementById("bulk-delete").addEventListener("click", async () => {
  const count = selectedIds.size;
  const ok = await confirmDialog("Sualları sil", count + " sualı silmək istədiyinizə əminsiniz? Bu əməliyyat geri qaytarıla bilməz.", "Sil");
  if (!ok) return;
  await BilGorekDB.deleteQuestions(Array.from(selectedIds));
  showToast(count + " sual silindi 🗑️");
  selectedIds.clear();
  await loadQuestions();
});

document.getElementById("search-input").addEventListener("input", applyFiltersAndRender);
document.getElementById("filter-category").addEventListener("change", applyFiltersAndRender);
document.getElementById("filter-difficulty").addEventListener("change", applyFiltersAndRender);
document.getElementById("filter-status").addEventListener("change", applyFiltersAndRender);

async function refreshFilterDropdowns() {
  await loadCategoriesCache();
  fillCategorySelect(document.getElementById("filter-category"), true, "Bütün kateqoriyalar");
  fillDifficultySelect(document.getElementById("filter-difficulty"), true);
  fillCategorySelect(document.getElementById("qf-kateqoriya"), false);
}

/* ======================= SUAL FORMASI (Əlavə / Redaktə) ======================= */
let editingQuestionId = null;

function openQuestionModal(existing) {
  editingQuestionId = existing ? existing.id : null;
  document.getElementById("question-modal-title").textContent = existing ? "Sualı redaktə et" : "Yeni sual";
  document.getElementById("duplicate-warn-box").classList.remove("show");

  document.getElementById("qf-sual").value = existing ? existing.sual : "";
  document.getElementById("qf-cavabA").value = existing ? existing.cavabA : "";
  document.getElementById("qf-cavabB").value = existing ? existing.cavabB : "";
  document.getElementById("qf-cavabC").value = existing ? existing.cavabC : "";
  document.getElementById("qf-cavabD").value = existing ? existing.cavabD : "";
  document.getElementById("qf-duzgunCavab").value = existing ? existing.duzgunCavab : "";
  document.getElementById("qf-kateqoriya").value = existing ? existing.kateqoriya : (categoriesCache[0] ? categoriesCache[0].key : "");
  document.getElementById("qf-cetinlik").value = existing ? existing.cetinlik : "Orta";
  document.getElementById("qf-izah").value = existing ? existing.izah : "";
  document.getElementById("qf-ipucu").value = existing ? existing.ipucu : "";
  document.getElementById("qf-aktiv").checked = existing ? !!existing.aktiv : true;

  document.querySelectorAll("#question-form .field").forEach(f => f.classList.remove("has-error"));
  document.getElementById("question-modal-overlay").style.display = "flex";
  document.getElementById("qf-save").dataset.forceSave = "0";
}
function closeQuestionModal() { document.getElementById("question-modal-overlay").style.display = "none"; }

document.getElementById("btn-add-question").addEventListener("click", () => openQuestionModal(null));
document.getElementById("qf-cancel").addEventListener("click", closeQuestionModal);

function getQuestionFormData() {
  return {
    sual: document.getElementById("qf-sual").value.trim(),
    cavabA: document.getElementById("qf-cavabA").value.trim(),
    cavabB: document.getElementById("qf-cavabB").value.trim(),
    cavabC: document.getElementById("qf-cavabC").value.trim(),
    cavabD: document.getElementById("qf-cavabD").value.trim(),
    duzgunCavab: document.getElementById("qf-duzgunCavab").value,
    kateqoriya: document.getElementById("qf-kateqoriya").value,
    cetinlik: document.getElementById("qf-cetinlik").value,
    izah: document.getElementById("qf-izah").value.trim(),
    ipucu: document.getElementById("qf-ipucu").value.trim(),
    aktiv: document.getElementById("qf-aktiv").checked
  };
}

function setFieldError(id, hasError) {
  const field = document.getElementById(id).closest(".field");
  field.classList.toggle("has-error", hasError);
}

function validateQuestionFormFields(data) {
  let valid = true;
  if (!data.sual) { setFieldError("qf-sual", true); valid = false; } else setFieldError("qf-sual", false);
  ["cavabA", "cavabB", "cavabC", "cavabD"].forEach(k => {
    const bad = !data[k];
    setFieldError("qf-" + k, bad);
    if (bad) valid = false;
  });
  if (!data.duzgunCavab) { setFieldError("qf-duzgunCavab", true); valid = false; } else setFieldError("qf-duzgunCavab", false);
  if (!data.kateqoriya) { setFieldError("qf-kateqoriya", true); valid = false; } else setFieldError("qf-kateqoriya", false);

  if (valid) {
    const answers = [data.cavabA, data.cavabB, data.cavabC, data.cavabD].map(a => a.trim().toLowerCase());
    const uniq = new Set(answers);
    if (uniq.size < 4) {
      showToast("Cavab variantları eyni ola bilməz.", true);
      valid = false;
    }
  }
  return valid;
}

document.getElementById("qf-save").addEventListener("click", async () => {
  const data = getQuestionFormData();
  if (!validateQuestionFormFields(data)) return;

  const forceSave = document.getElementById("qf-save").dataset.forceSave === "1";
  if (!forceSave) {
    const dup = await BilGorekDB.findDuplicateQuestion(data.sual, editingQuestionId);
    if (dup) {
      document.getElementById("duplicate-warn-box").classList.add("show");
      document.getElementById("qf-save").textContent = "⚠️ Yenə də saxla";
      document.getElementById("qf-save").dataset.forceSave = "1";
      return;
    }
  }

  try {
    if (editingQuestionId) {
      await BilGorekDB.updateQuestion(editingQuestionId, data);
      showToast("Sual yeniləndi ✅");
    } else {
      await BilGorekDB.addQuestion(data);
      showToast("Sual əlavə edildi ✅");
    }
    document.getElementById("qf-save").textContent = "💾 Yadda saxla";
    closeQuestionModal();
    await loadQuestions();
  } catch (e) {
    showToast("Xəta: " + e.message, true);
  }
});

async function deleteQuestionFlow(q) {
  const ok = await confirmDialog("Sualı sil", `"${q.sual.slice(0, 60)}${q.sual.length > 60 ? "…" : ""}" sualını silmək istədiyinizə əminsiniz?`, "Sil");
  if (!ok) return;
  await BilGorekDB.deleteQuestion(q.id);
  showToast("Sual silindi 🗑️");
  await loadQuestions();
}

/* ======================= İDXAL ======================= */
let importParsedRows = []; // { rowNumber, raw, status, reasons, normalized, isDuplicate, dupMatchId }

document.getElementById("dropzone").addEventListener("click", () => document.getElementById("import-file-input").click());
document.getElementById("import-file-input").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  try {
    const rawRows = await parseImportFile(file);
    await validateAndPreviewImport(rawRows);
  } catch (err) {
    showToast("Fayl oxunmadı: " + err.message, true);
  }
  e.target.value = "";
});

async function parseImportFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".json")) {
    const text = await file.text();
    let data;
    try { data = JSON.parse(text); } catch (e) { throw new Error("JSON formatı yanlışdır"); }
    if (data && !Array.isArray(data) && Array.isArray(data.questions)) {
      throw new Error("Bu fayl tam ehtiyat nüsxəsi kimi görünür. Zəhmət olmasa 'İxrac / Ehtiyat nüsxəsi' bölməsindəki 'Bərpa et' funksiyasından istifadə edin.");
    }
    if (!Array.isArray(data)) throw new Error("JSON faylı sual siyahısı (massiv) olmalıdır");
    return data;
  }
  if (name.endsWith(".csv")) {
    const text = await file.text();
    const rows = csvParse(text);
    if (rows.length === 0) return [];
    const headers = rows[0].map(h => h.trim());
    return rows.slice(1).map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i] !== undefined ? r[i] : ""; });
      return obj;
    });
  }
  if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { defval: "" });
  }
  throw new Error("Dəstəklənməyən fayl formatı. .xlsx, .csv və ya .json seçin.");
}

async function validateAndPreviewImport(rawRows) {
  await loadCategoriesCache();
  const existingQuestions = await BilGorekDB.getAllQuestions();
  const existingTextsLower = new Set(existingQuestions.map(q => q.sual.trim().toLowerCase()));
  const existingById = {};
  existingQuestions.forEach(q => { existingById[q.id] = q; });

  const batchTextsSeen = new Map(); // lowertext -> first rowNumber

  importParsedRows = rawRows.map((raw, idx) => {
    const rowNumber = idx + 2; // 1-ci sətir başlıqdır (insan gözü ilə say)
    const reasons = [];
    let statusIsError = false;

    const sual = (raw.sual || "").toString().trim();
    if (!sual) { reasons.push("Sual mətni boşdur"); statusIsError = true; }

    const cavabA = (raw.cavabA || "").toString().trim();
    const cavabB = (raw.cavabB || "").toString().trim();
    const cavabC = (raw.cavabC || "").toString().trim();
    const cavabD = (raw.cavabD || "").toString().trim();
    if (!cavabA || !cavabB || !cavabC || !cavabD) { reasons.push("Cavab variantlarından biri (A/B/C/D) boşdur"); statusIsError = true; }

    const answersLower = [cavabA, cavabB, cavabC, cavabD].map(a => a.toLowerCase());
    if (new Set(answersLower).size < 4 && cavabA && cavabB && cavabC && cavabD) {
      reasons.push("Cavab variantları eyni ola bilməz");
      statusIsError = true;
    }

    let duzgunCavab = (raw.duzgunCavab || "").toString().trim().toUpperCase();
    if (!["A", "B", "C", "D"].includes(duzgunCavab)) {
      reasons.push("Düzgün cavab A/B/C/D olmalıdır (verilib: \"" + (raw.duzgunCavab || "") + "\")");
      statusIsError = true;
    }

    let kateqoriyaKey = resolveCategoryInput(raw.kateqoriya, categoriesCache);
    if (!kateqoriyaKey) {
      reasons.push("Kateqoriya tapılmadı: \"" + (raw.kateqoriya || "") + "\"");
      statusIsError = true;
    }

    let cetinlik = (raw.cetinlik || "").toString().trim();
    if (!cetinlik) cetinlik = "Orta";
    else {
      const match = DIFFICULTIES.find(d => d.toLowerCase() === cetinlik.toLowerCase());
      if (!match) { reasons.push("Çətinlik səviyyəsi yanlışdır (Asan/Orta/Çətin olmalıdır)"); statusIsError = true; }
      else cetinlik = match;
    }

    const aktivParsed = parseBoolLike(raw.aktiv, true);
    if (aktivParsed === null) { reasons.push("aktiv sahəsi true/false olmalıdır"); statusIsError = true; }

    let idVal;
    if (raw.id !== undefined && raw.id !== null && String(raw.id).trim() !== "") {
      const n = Number(raw.id);
      if (isNaN(n)) { reasons.push("id ədəd olmalıdır"); statusIsError = true; }
      else idVal = n;
    }

    // Təkrar yoxlanışı (yalnız sual mətni boş deyilsə mənalıdır)
    let isDuplicate = false;
    let dupReason = "";
    if (sual) {
      const lower = sual.toLowerCase();
      if (existingTextsLower.has(lower)) { isDuplicate = true; dupReason = "Bazada eyni mətnli sual artıq var"; }
      else if (idVal !== undefined && existingById[idVal]) { isDuplicate = true; dupReason = "Bu ID (" + idVal + ") ilə sual artıq var"; }
      else if (batchTextsSeen.has(lower)) { isDuplicate = true; dupReason = "Fayl daxilində təkrarlanır (sətir " + batchTextsSeen.get(lower) + ")"; }
      else { batchTextsSeen.set(lower, rowNumber); }
    }

    return {
      rowNumber,
      raw,
      status: statusIsError ? "error" : "ok",
      reasons,
      isDuplicate,
      dupReason,
      normalized: statusIsError ? null : {
        id: idVal,
        sual, cavabA, cavabB, cavabC, cavabD, duzgunCavab,
        kateqoriya: kateqoriyaKey,
        cetinlik,
        izah: (raw.izah || "").toString().trim(),
        ipucu: (raw.ipucu || "").toString().trim(),
        aktiv: aktivParsed === null ? true : aktivParsed
      }
    };
  });

  renderImportPreview();
}

function renderImportPreview() {
  const area = document.getElementById("import-preview-area");
  area.style.display = "block";

  const total = importParsedRows.length;
  const errorCount = importParsedRows.filter(r => r.status === "error").length;
  const dupCount = importParsedRows.filter(r => r.status === "ok" && r.isDuplicate).length;
  const okCount = total - errorCount;

  document.getElementById("import-summary-grid").innerHTML = `
    <div class="import-stat"><div class="n">${total}</div><div class="l">Cəmi sətir</div></div>
    <div class="import-stat"><div class="n">${okCount}</div><div class="l">Düzgün</div></div>
    <div class="import-stat dup"><div class="n">${dupCount}</div><div class="l">Təkrar</div></div>
    <div class="import-stat err"><div class="n">${errorCount}</div><div class="l">Səhv</div></div>
  `;

  const tbody = document.getElementById("import-preview-tbody");
  tbody.innerHTML = "";
  importParsedRows.forEach((r, i) => {
    const tr = document.createElement("tr");
    if (r.status === "error") tr.classList.add("preview-row-error");
    else if (r.isDuplicate) tr.classList.add("preview-row-dup");

    const checked = r.status === "ok" ? "checked" : "";
    const disabled = r.status === "error" ? "disabled" : "";
    const statusLabel = r.status === "error" ? "❌ Səhv" : (r.isDuplicate ? "🔁 Təkrar" : "✅ Düzgün");
    const reasonHtml = r.status === "error"
      ? `<div class="preview-reason">${escapeHtml(r.reasons.join("; "))}</div>`
      : (r.isDuplicate ? `<div class="preview-reason" style="color:var(--gold-2)">${escapeHtml(r.dupReason)}</div>` : "");

    const questionText = (r.raw.sual || "").toString();
    const shortQ = questionText.length > 70 ? questionText.slice(0, 70) + "…" : questionText;

    tr.innerHTML = `
      <td><input type="checkbox" class="import-row-check" data-idx="${i}" ${checked} ${disabled}></td>
      <td class="nowrap">${r.rowNumber}</td>
      <td class="q-text">${escapeHtml(shortQ)}${reasonHtml}</td>
      <td class="nowrap">${escapeHtml((r.raw.kateqoriya || "").toString())}</td>
      <td class="nowrap">${statusLabel}</td>
    `;
    tbody.appendChild(tr);
  });
}

document.getElementById("btn-confirm-import").addEventListener("click", async () => {
  const skipDup = document.getElementById("opt-skip-duplicates").checked;
  const updateExisting = document.getElementById("opt-update-existing").checked;

  const checks = document.querySelectorAll(".import-row-check");
  const includedIdx = new Set();
  checks.forEach(c => { if (c.checked) includedIdx.add(Number(c.dataset.idx)); });

  let added = 0, updated = 0, skipped = 0, failed = 0;
  const existingQuestions = await BilGorekDB.getAllQuestions();
  const existingTextMap = new Map(existingQuestions.map(q => [q.sual.trim().toLowerCase(), q]));

  for (let i = 0; i < importParsedRows.length; i++) {
    const r = importParsedRows[i];
    if (r.status === "error") { failed++; continue; }
    if (!includedIdx.has(i)) { skipped++; continue; }

    const norm = r.normalized;
    try {
      if (r.isDuplicate) {
        if (updateExisting) {
          const match = (norm.id !== undefined && existingQuestions.find(q => q.id === norm.id)) || existingTextMap.get(norm.sual.toLowerCase());
          if (match) {
            await BilGorekDB.updateQuestion(match.id, norm);
            updated++;
          } else {
            await BilGorekDB.addQuestion(norm);
            added++;
          }
        } else if (skipDup) {
          skipped++;
        } else {
          await BilGorekDB.addQuestion(norm);
          added++;
        }
      } else {
        if (norm.id !== undefined) {
          await BilGorekDB.putQuestionRaw(norm);
        } else {
          await BilGorekDB.addQuestion(norm);
        }
        added++;
      }
    } catch (e) {
      failed++;
    }
  }

  showToast(`İdxal tamamlandı: ${added} əlavə edildi, ${updated} yeniləndi, ${skipped} ötürüldü, ${failed} səhv ✅`);
  document.getElementById("import-preview-area").style.display = "none";
  importParsedRows = [];
  await loadQuestions();
  await renderCategoriesTab();
});

/* ======================= ŞABLONLAR ======================= */
async function buildExampleRows() {
  await loadCategoriesCache();
  const allQuestions = await BilGorekDB.getAllQuestions();
  const rows = [];
  for (const cat of categoriesCache) {
    const existing = allQuestions.find(q => q.kateqoriya === cat.key && q.aktiv);
    if (existing) {
      rows.push(dbRowToExportRow(existing, categoriesByKeyCache));
    } else {
      rows.push({
        id: "", sual: `${cat.name} kateqoriyasından nümunə sual?`,
        cavabA: "Variant A", cavabB: "Variant B", cavabC: "Variant C", cavabD: "Variant D",
        duzgunCavab: "A", kateqoriya: cat.name, cetinlik: "Orta", izah: "", ipucu: "", aktiv: "true"
      });
    }
  }
  return rows;
}

document.getElementById("tpl-blank-xlsx").addEventListener("click", () => {
  const ws = XLSX.utils.aoa_to_sheet([ROW_COLUMNS]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Suallar");
  XLSX.writeFile(wb, "bilgorek-bos-sablon.xlsx");
});

document.getElementById("tpl-example-xlsx").addEventListener("click", async () => {
  const rows = await buildExampleRows();
  const ws = XLSX.utils.json_to_sheet(rows, { header: ROW_COLUMNS });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Suallar");
  XLSX.writeFile(wb, "bilgorek-numuneli-sablon.xlsx");
});

document.getElementById("tpl-csv").addEventListener("click", async () => {
  const rows = await buildExampleRows();
  const csv = csvStringify(ROW_COLUMNS, rows);
  downloadBlob("bilgorek-sablon.csv", csv, "text/csv;charset=utf-8");
});

document.getElementById("tpl-json").addEventListener("click", async () => {
  const rows = await buildExampleRows();
  downloadBlob("bilgorek-sablon.json", JSON.stringify(rows, null, 2), "application/json");
});

/* ======================= İXRAC ======================= */
async function getExportRows() {
  await loadCategoriesCache();
  const all = await BilGorekDB.getAllQuestions();
  return all.sort((a, b) => a.id - b.id).map(q => dbRowToExportRow(q, categoriesByKeyCache));
}

document.getElementById("export-xlsx").addEventListener("click", async () => {
  const rows = await getExportRows();
  const ws = XLSX.utils.json_to_sheet(rows, { header: ROW_COLUMNS });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Suallar");
  XLSX.writeFile(wb, "bilgorek-suallar.xlsx");
  showToast("Excel faylı endirildi ✅");
});
document.getElementById("export-csv").addEventListener("click", async () => {
  const rows = await getExportRows();
  downloadBlob("bilgorek-suallar.csv", csvStringify(ROW_COLUMNS, rows), "text/csv;charset=utf-8");
  showToast("CSV faylı endirildi ✅");
});
document.getElementById("export-json").addEventListener("click", async () => {
  const rows = await getExportRows();
  downloadBlob("bilgorek-suallar.json", JSON.stringify(rows, null, 2), "application/json");
  showToast("JSON faylı endirildi ✅");
});

/* ======================= EHTİYAT NÜSXƏSİ ======================= */
document.getElementById("btn-create-backup").addEventListener("click", async () => {
  const data = await BilGorekDB.exportAll();
  const stamp = new Date().toISOString().slice(0, 10);
  downloadBlob(`bilgorek-ehtiyat-${stamp}.json`, JSON.stringify(data, null, 2), "application/json");
  showToast("Ehtiyat nüsxəsi yaradıldı 💾");
});

let restoreFile = null;
document.getElementById("restore-dropzone").addEventListener("click", () => document.getElementById("restore-file-input").click());
document.getElementById("restore-file-input").addEventListener("change", (e) => {
  restoreFile = e.target.files[0] || null;
  document.getElementById("restore-file-name").textContent = restoreFile ? "Seçilmiş fayl: " + restoreFile.name : "";
  document.getElementById("btn-do-restore").disabled = !restoreFile;
});

document.getElementById("btn-do-restore").addEventListener("click", async () => {
  if (!restoreFile) return;
  const mode = document.querySelector('input[name="restore-mode"]:checked').value;
  const modeLabel = mode === "overwrite" ? "MÖVCUD MƏLUMATLARIN ÜZƏRİNƏ YAZILACAQ" : "mövcud məlumatlarla birləşdiriləcək";
  const ok = await confirmDialog(
    "Ehtiyat nüsxəsini bərpa et",
    mode === "overwrite"
      ? "Diqqət: bütün mövcud suallar, kateqoriyalar və lider lövhəsi SİLİNƏCƏK və faylın içindəkilərlə əvəz olunacaq. Davam edilsin?"
      : "Fayldakı suallar və kateqoriyalar mövcud məlumatlarla birləşdiriləcək (id/açar üst-üstə düşərsə fayldakı məlumat üstünlük təşkil edəcək). Davam edilsin?",
    "Bərpa et"
  );
  if (!ok) return;

  try {
    const text = await restoreFile.text();
    const data = JSON.parse(text);
    const summary = await BilGorekDB.restoreAll(data, mode);
    showToast(`Bərpa tamamlandı: ${summary.categoriesAdded + summary.categoriesUpdated} kateqoriya, ${summary.questionsAdded + summary.questionsUpdated} sual, ${summary.leaderboardAdded} reytinq qeydi ✅`);
    restoreFile = null;
    document.getElementById("restore-file-name").textContent = "";
    document.getElementById("btn-do-restore").disabled = true;
    await loadCategoriesCache();
    await refreshFilterDropdowns();
    await renderCategoriesTab();
    await loadQuestions();
    await loadGameParamsForm();
  } catch (e) {
    showToast("Bərpa alınmadı: " + e.message, true);
  }
});

/* ======================= SERVER (FIREBASE) ======================= */
function renderServerStatus() {
  const statusBox = document.getElementById("server-status-box");
  const connectBlock = document.getElementById("server-connect-block");
  const disconnectBlock = document.getElementById("server-disconnect-block");
  if (!statusBox) return;
  const backend = BilGorekDB.getBackendName ? BilGorekDB.getBackendName() : null;
  if (backend === "firestore") {
    statusBox.textContent = "🟢 Server ilə qoşulu — suallar bütün cihazlarla sinxrondur.";
    connectBlock.style.display = "none";
    disconnectBlock.style.display = "block";
  } else {
    statusBox.textContent = "⚪ Yalnız bu cihazda — server qoşulmayıb, suallar yalnız bu telefonda saxlanılır.";
    connectBlock.style.display = "block";
    disconnectBlock.style.display = "none";
  }
}

document.getElementById("btn-connect-firebase").addEventListener("click", async () => {
  const input = document.getElementById("firebase-config-input");
  const field = input.closest(".field");
  let cfg;
  try {
    cfg = JSON.parse(input.value);
    if (!cfg || typeof cfg !== "object" || !cfg.apiKey || !cfg.projectId) throw new Error("missing fields");
  } catch (e) {
    field.classList.add("has-error");
    return;
  }
  field.classList.remove("has-error");
  const ok = await confirmDialog(
    "Serverə qoşul",
    "Bu cihaz Firebase serverinə qoşulacaq. Server boşdursa, bu cihazdakı suallar ora köçürüləcək; server artıq doldurulubsa, oradakı suallar bu cihazda göstəriləcək. Davam edilsin?",
    "Qoşul"
  );
  if (!ok) return;
  try {
    BilGorekDB.saveFirebaseConfig(cfg);
    showToast("Serverə qoşulur, səhifə yenilənir…");
    setTimeout(() => location.reload(), 700);
  } catch (e) {
    showToast("Qoşulma alınmadı: " + e.message, true);
  }
});

document.getElementById("btn-disconnect-firebase").addEventListener("click", async () => {
  const ok = await confirmDialog(
    "Yalnız bu cihazda işlə",
    "Bu cihaz serverdən ayrılacaq və yalnız bu telefonda saxlanan məlumatlarla işləyəcək. Digər cihazlar serverə qoşulu qalmağa davam edəcək. Davam edilsin?",
    "Ayrıl"
  );
  if (!ok) return;
  BilGorekDB.removeFirebaseConfig();
  showToast("Serverdən ayrıldı, səhifə yenilənir…");
  setTimeout(() => location.reload(), 700);
});

/* ======================= ADMİN PAROLU ======================= */
document.getElementById("btn-change-pin").addEventListener("click", async () => {
  const errBox = document.getElementById("pin-error");
  const currentInput = document.getElementById("pin-current");
  const newInput = document.getElementById("pin-new");
  const confirmInput = document.getElementById("pin-new-confirm");
  errBox.style.color = "var(--danger)";
  errBox.textContent = "";

  const current = currentInput.value;
  const next = newInput.value.trim();
  const confirm = confirmInput.value.trim();

  const ok = await BilGorekDB.checkAdminPin(current);
  if (!ok) { errBox.textContent = "Hazırkı parol yanlışdır."; return; }
  if (next.length < 4) { errBox.textContent = "Yeni parol ən azı 4 simvol olmalıdır."; return; }
  if (next !== confirm) { errBox.textContent = "Yeni parollar üst-üstə düşmür."; return; }

  await BilGorekDB.setAdminPin(next);
  errBox.style.color = "var(--success)";
  errBox.textContent = "Parol dəyişdirildi ✅";
  currentInput.value = ""; newInput.value = ""; confirmInput.value = "";
  showToast("Admin parolu dəyişdirildi ✅");
});

/* ======================= OYUN PARAMETRLƏRİ ======================= */
async function loadGameParamsForm() {
  const gp = await BilGorekDB.getGameParams();
  document.getElementById("gp-questionsPerRound").value = gp.questionsPerRound;
  document.getElementById("gp-timePerQuestion").value = gp.timePerQuestion;
  document.getElementById("gp-basePoints").value = gp.basePoints;
  document.getElementById("gp-timeBonusPerSec").value = gp.timeBonusPerSec;
}
document.getElementById("btn-save-gameparams").addEventListener("click", async () => {
  const patch = {
    questionsPerRound: Math.max(1, Number(document.getElementById("gp-questionsPerRound").value) || 10),
    timePerQuestion: Math.max(5, Number(document.getElementById("gp-timePerQuestion").value) || 20),
    basePoints: Math.max(0, Number(document.getElementById("gp-basePoints").value) || 0),
    timeBonusPerSec: Math.max(0, Number(document.getElementById("gp-timeBonusPerSec").value) || 0)
  };
  await BilGorekDB.setGameParams(patch);
  showToast("Oyun parametrləri yadda saxlanıldı ✅");
});

/* ======================= BAŞLANĞIC YÜKLƏMƏ ======================= */
async function refreshAdminData() {
  try {
    await BilGorekDB.ready();
    await refreshFilterDropdowns();
    await loadQuestions();
    await renderCategoriesTab();
    await loadGameParamsForm();
    renderServerStatus();
  } catch (e) {
    console.error(e);
    showToast("Məlumat bazası açılmadı: " + e.message, true);
  }
}
refreshAdminData();
// Oyun ekranındakı "⚙️ Admin panel" düyməsi admin görünüşünü göstərəndə bu
// funksiyanı çağırır ki, ekran hər açılanda ən son məlumatlar görünsün.
window.__bilgorekAdminRefresh = refreshAdminData;

})();
