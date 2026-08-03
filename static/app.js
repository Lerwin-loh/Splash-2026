const fileInput = document.getElementById("fileInput");
const previousFileInput = document.getElementById("previousFileInput");
const dropZone = document.getElementById("dropZone");
const previousDropZone = document.getElementById("previousDropZone");
const previousSectionToggle = document.getElementById("previousSectionToggle");
const previousSectionBody = document.getElementById("previousSectionBody");
const fileName = document.getElementById("fileName");
const fileSize = document.getElementById("fileSize");
const previousFileName = document.getElementById("previousFileName");
const previousWorksheetSelect = document.getElementById("previousWorksheetSelect");
const previousApplicantCount = document.getElementById("previousApplicantCount");
const worksheetSelect = document.getElementById("worksheetSelect");
const applicantCountInput = document.getElementById("applicantCount");
const columnsBox = document.getElementById("columnsBox");
const previewTable = document.getElementById("previewTable");
const firstColumn = document.getElementById("firstColumn");
const secondColumn = document.getElementById("secondColumn");
const thirdColumn = document.getElementById("thirdColumn");
const maxApplicants = document.getElementById("maxApplicants");
const totalMentors = document.getElementById("totalMentors");
const mentorSum = document.getElementById("mentorSum");
const validationBox = document.getElementById("validationBox");
const generateButton = document.getElementById("generateButton");
const loading = document.getElementById("loading");
const errorBox = document.getElementById("errorBox");
const resultsSection = document.getElementById("resultsSection");
const summaryCards = document.getElementById("summaryCards");
const downloadLink = document.getElementById("downloadLink");
const domains = JSON.parse(document.getElementById("domainData").textContent);

let selectedFile = null;
let previousFile = null;
let inspectData = null;
let downloadUrl = null;

function clearDownloadState() {
  if (downloadUrl) URL.revokeObjectURL(downloadUrl);
  downloadUrl = null;
  downloadLink.removeAttribute("href");
  resultsSection.classList.add("hidden");
  summaryCards.innerHTML = "";
}

function clearCurrentInspection() {
  inspectData = null;
  worksheetSelect.innerHTML = "";
  applicantCountInput.value = "-";
  columnsBox.textContent = "";
  previewTable.innerHTML = "";
  [firstColumn, secondColumn, thirdColumn].forEach((select) => {
    select.innerHTML = '<option value="">Choose column</option>';
  });
  clearDownloadState();
}

function clearPreviousInspection() {
  previousApplicantCount.value = "-";
  previousWorksheetSelect.innerHTML = "";
  clearDownloadState();
}

function formatBytes(bytes) {
  if (!bytes) return "";
  const units = ["B", "KB", "MB"];
  let value = bytes;
  let unit = 0;
  while (value > 1024 && unit < units.length - 1) {
    value /= 1024;
    unit += 1;
  }
  return `${value.toFixed(unit === 0 ? 0 : 1)} ${units[unit]}`;
}

function safeNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function mentorCounts() {
  const counts = {};
  document.querySelectorAll(".mentor-input").forEach((input) => {
    counts[input.dataset.domain] = Number(input.value || 0);
  });
  return counts;
}

function mappingValues() {
  return {
    first: firstColumn.value,
    second: secondColumn.value,
    third: thirdColumn.value,
  };
}

function setError(message) {
  errorBox.textContent = message || "";
  errorBox.classList.toggle("hidden", !message);
}

function addOptions(select, columns, selected) {
  select.innerHTML = '<option value="">Choose column</option>';
  columns.forEach((column) => {
    const option = document.createElement("option");
    option.value = column;
    option.textContent = column;
    option.selected = column === selected;
    select.appendChild(option);
  });
}

function renderPreview(rows, columns) {
  previewTable.innerHTML = "";
  if (!rows.length) return;
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  columns.forEach((column) => {
    const th = document.createElement("th");
    th.textContent = column;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const tbody = document.createElement("tbody");
  rows.forEach((row) => {
    const bodyRow = document.createElement("tr");
    columns.forEach((column) => {
      const td = document.createElement("td");
      td.textContent = row[column] || "";
      bodyRow.appendChild(td);
    });
    tbody.appendChild(bodyRow);
  });
  previewTable.append(thead, tbody);
}

function renderValidation() {
  const counts = mentorCounts();
  const total = Number(totalMentors.value || 0);
  const max = Number(maxApplicants.value || 0);
  const sum = Object.values(counts).reduce((acc, value) => acc + value, 0);
  const activeDomains = Object.values(counts).filter((value) => value > 0).length;
  const applicants = inspectData ? inspectData.applicantCount : 0;
  const previousRows = safeNumber(previousApplicantCount.value);
  const group1 = Math.ceil(applicants / 2);
  const group2 = Math.floor(applicants / 2);
  const initialTotalCapacity = sum * 5;
  const likelyExpansion = Math.max(group1, group2) * 2 > initialTotalCapacity;
  const mappings = mappingValues();
  const mapped = mappings.first && mappings.second && mappings.third && new Set(Object.values(mappings)).size === 3;
  const validConfig = selectedFile && mapped && max > 0 && applicants > 0 && applicants <= max && total > 0 && sum === total && activeDomains >= 2 && Object.values(counts).every((value) => value >= 0);

  mentorSum.value = `${sum} / ${total}`;
  validationBox.innerHTML = `
    <div class="validation-row full"><span>Total configured mentors</span><strong>${total}</strong></div>
    <div class="validation-row full"><span>Sum across subdomains</span><strong class="${sum === total ? "ok" : "bad"}">${sum}</strong></div>
    <div class="validation-row full"><span>Mentor totals match</span><strong class="${sum === total ? "ok" : "bad"}">${sum === total ? "Yes" : "No"}</strong></div>
    <div class="validation-row compact"><span>Current registration rows</span><strong>${applicants}</strong></div>
    <div class="validation-row compact"><span>Previous allocation rows</span><strong>${previousRows}</strong></div>
    <div class="validation-row compact"><span>Estimated Group 1 size</span><strong>${group1}</strong></div>
    <div class="validation-row compact"><span>Estimated Group 2 size</span><strong>${group2}</strong></div>
    ${domains.map((domain) => `<div class="validation-row compact"><span>${domain} initial capacity</span><strong>${counts[domain] * 5}</strong></div>`).join("")}
    <div class="validation-row compact"><span>Initial total capacity per group</span><strong>${initialTotalCapacity}</strong></div>
    <div class="validation-row full">
      <span class="label-with-info">
        Capacity expansion likely
        <span class="info-icon" tabindex="0" aria-label="Capacity expansion information" data-tooltip="The app starts with 5 participants per mentor. If that is not enough to allocate everyone, it automatically tries 6, then 7, and so on until allocation is possible.">i</span>
      </span>
      <strong class="${likelyExpansion ? "warn" : "ok"}">${likelyExpansion ? "Yes" : "No"}</strong>
    </div>
    <div class="validation-row full"><span>Ready to generate</span><strong class="${validConfig ? "ok" : "bad"}">${validConfig ? "Yes" : "No"}</strong></div>
  `;
  generateButton.disabled = !validConfig;
}

async function inspectWorkbook() {
  if (!selectedFile) return;
  setError("");
  const formData = new FormData();
  formData.append("file", selectedFile);
  if (worksheetSelect.value) formData.append("worksheet", worksheetSelect.value);
  const response = await fetch("/api/inspect", { method: "POST", body: formData });
  const data = await response.json();
  if (!response.ok) {
    setError(data.error || "Unable to inspect workbook.");
    return;
  }
  inspectData = data;
  worksheetSelect.innerHTML = "";
  data.worksheets.forEach((sheet) => {
    const option = document.createElement("option");
    option.value = sheet;
    option.textContent = sheet;
    option.selected = sheet === data.selectedWorksheet;
    worksheetSelect.appendChild(option);
  });
  applicantCountInput.value = data.applicantCount;
  columnsBox.textContent = `Detected columns: ${data.columns.join(", ")}`;
  addOptions(firstColumn, data.columns, data.suggestedMappings.first);
  addOptions(secondColumn, data.columns, data.suggestedMappings.second);
  addOptions(thirdColumn, data.columns, data.suggestedMappings.third);
  renderPreview(data.previewRows, data.columns);
  renderValidation();
}

function setFile(file) {
  selectedFile = file;
  clearCurrentInspection();
  fileName.textContent = file ? file.name : "No file selected";
  fileSize.textContent = file ? formatBytes(file.size) : "";
  if (file) inspectWorkbook();
  renderValidation();
}

function setPreviousFile(file) {
  previousFile = file || null;
  clearPreviousInspection();
  previousFileName.textContent = file
    ? `${file.name} (${formatBytes(file.size)})`
    : "No previous allocation selected";
  if (file) setPreviousSectionOpen(true);
  if (file) inspectPreviousWorkbook();
  renderValidation();
}

function setPreviousSectionOpen(open) {
  previousSectionBody.classList.toggle("hidden", !open);
  previousSectionToggle.setAttribute("aria-expanded", open ? "true" : "false");
  previousSectionToggle.querySelector(".toggle-icon").textContent = open ? "^" : "v";
}

fileInput.addEventListener("change", () => setFile(fileInput.files[0]));
previousFileInput.addEventListener("change", () => setPreviousFile(previousFileInput.files[0]));
worksheetSelect.addEventListener("change", inspectWorkbook);
previousWorksheetSelect.addEventListener("change", inspectPreviousWorkbook);
previousSectionToggle.addEventListener("click", () => {
  setPreviousSectionOpen(previousSectionBody.classList.contains("hidden"));
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragging");
});

dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragging"));
dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragging");
  if (event.dataTransfer.files.length) {
    fileInput.files = event.dataTransfer.files;
    setFile(event.dataTransfer.files[0]);
  }
});

previousDropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  previousDropZone.classList.add("dragging");
});

previousDropZone.addEventListener("dragleave", () => previousDropZone.classList.remove("dragging"));
previousDropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  previousDropZone.classList.remove("dragging");
  if (event.dataTransfer.files.length) {
    previousFileInput.files = event.dataTransfer.files;
    setPreviousFile(event.dataTransfer.files[0]);
  }
});

async function inspectPreviousWorkbook() {
  if (!previousFile) return;
  setError("");
  const formData = new FormData();
  formData.append("previousFile", previousFile);
  if (previousWorksheetSelect.value) formData.append("previousWorksheet", previousWorksheetSelect.value);
  const response = await fetch("/api/inspect-previous", { method: "POST", body: formData });
  const data = await response.json();
  if (!response.ok) {
    setError(data.error || "Unable to inspect previous allocation workbook.");
    return;
  }
  previousWorksheetSelect.innerHTML = "";
  data.worksheets.forEach((sheet) => {
    const option = document.createElement("option");
    option.value = sheet;
    option.textContent = sheet;
    option.selected = sheet === data.selectedWorksheet;
    previousWorksheetSelect.appendChild(option);
  });
  previousApplicantCount.value = data.applicantCount;
  if (!data.hasGeneratedColumns) {
    setError(`Previous worksheet is missing generated columns: ${data.missingColumns.join(", ")}`);
  }
  renderValidation();
}

document.querySelectorAll("input, select").forEach((element) => {
  element.addEventListener("input", renderValidation);
  element.addEventListener("change", renderValidation);
});

function filenameFromHeader(header) {
  const match = /filename="?([^"]+)"?/i.exec(header || "");
  return match ? match[1] : "mentoring_allocation.xlsx";
}

function renderResults(analytics) {
  const cards = [
    ["Total applicants", analytics.total_applicants],
    ["Group 1 size", analytics.groups["Group 1"].count],
    ["Group 2 size", analytics.groups["Group 2"].count],
    ["Total mentors", analytics.total_mentors],
    ["Group 1 max participants per mentor", analytics.groups["Group 1"].final_capacity_per_mentor],
    ["Group 2 max participants per mentor", analytics.groups["Group 2"].final_capacity_per_mentor],
    ["Previously allocated rows", analytics.previous_allocation_rows],
    ["Existing applicants matched", analytics.matched_existing_applicants],
    ["New applicants allocated", analytics.new_applicants_allocated],
    ["Unallocated applicants", analytics.total_unsuccessful_applicants],
  ];
  summaryCards.innerHTML = cards.map(([label, value]) => `<div class="card"><span>${label}</span><strong>${value}</strong></div>`).join("");
  if (analytics.capacity_warnings && analytics.capacity_warnings.length) {
    summaryCards.innerHTML += analytics.capacity_warnings.map((item) => `
      <div class="card warning-card">
        <span>${item.group} ${item.domain} over expected capacity</span>
        <strong>+${item.over_by}</strong>
      </div>
    `).join("");
  }
  resultsSection.classList.remove("hidden");
}

generateButton.addEventListener("click", async () => {
  if (!selectedFile) return;
  setError("");
  loading.classList.remove("hidden");
  generateButton.disabled = true;
  const formData = new FormData();
  formData.append("file", selectedFile);
  formData.append("worksheet", worksheetSelect.value);
  formData.append("mappings", JSON.stringify(mappingValues()));
  formData.append("maxApplicants", maxApplicants.value);
  formData.append("totalMentors", totalMentors.value);
  formData.append("mentorCounts", JSON.stringify(mentorCounts()));
  if (previousFile) formData.append("previousFile", previousFile);
  if (previousFile && previousWorksheetSelect.value) formData.append("previousWorksheet", previousWorksheetSelect.value);

  try {
    const response = await fetch("/api/process", { method: "POST", body: formData });
    if (!response.ok) {
      const data = await response.json();
      throw new Error(data.error || "Processing failed.");
    }
    const blob = await response.blob();
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    downloadUrl = URL.createObjectURL(blob);
    downloadLink.href = downloadUrl;
    downloadLink.download = filenameFromHeader(response.headers.get("Content-Disposition"));
    const analytics = JSON.parse(atob(response.headers.get("X-Allocation-Analytics")));
    renderResults(analytics);
  } catch (error) {
    setError(error.message);
  } finally {
    loading.classList.add("hidden");
    renderValidation();
  }
});

renderValidation();
