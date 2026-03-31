const DAY_MS = 24 * 60 * 60 * 1000;

const initialTasks = [
  {
    id: crypto.randomUUID(),
    name: "Discovery",
    start: "2026-03-24",
    end: "2026-03-28",
    color: "#ff7a59",
    owner: "RB",
    complete: false,
    milestone: false,
  },
  {
    id: crypto.randomUUID(),
    name: "Design Sprint",
    start: "2026-03-29",
    end: "2026-04-04",
    color: "#0f9d8a",
    owner: "AK",
    complete: false,
    milestone: false,
  },
  {
    id: crypto.randomUUID(),
    name: "Build Phase",
    start: "2026-04-02",
    end: "2026-04-10",
    color: "#3366ff",
    owner: "JT",
    complete: false,
    milestone: false,
  },
];

const state = {
  tasks: initialTasks,
  planStart: "2026-03-22",
  planEnd: "2026-04-14",
  planTitle: "",
  zoom: 100,
  drag: null,
  labelColumnWidth: 320,
  compactStartColumnWidth: 74,
  compactEndColumnWidth: 74,
  detailLevel: "compressed",
  showCompactDates: true,
  editingTaskId: null,
  editingTaskDateField: null,
  editingTaskNameId: null,
  editingTaskOwnerId: null,
  rowDrag: null,
};

const board = document.getElementById("gantt-board");
const emptyState = document.getElementById("empty-state");
const form = document.getElementById("task-form");
const appShell = document.querySelector(".app-shell");
const panelToggle = document.getElementById("panel-toggle");
const nameInput = document.getElementById("task-name");
const startInput = document.getElementById("task-start");
const endInput = document.getElementById("task-end");
const ownerInput = document.getElementById("task-owner");
const colorInput = document.getElementById("task-color");
const colorValue = document.getElementById("task-color-value");
const planForm = document.getElementById("plan-form");
const planTitleInput = document.getElementById("plan-title");
const planStartInput = document.getElementById("plan-start");
const planEndInput = document.getElementById("plan-end");
const detailLevelInput = document.getElementById("detail-level");
const showCompactDatesInput = document.getElementById("show-compact-dates");
const zoomOutButton = document.getElementById("zoom-out");
const zoomInButton = document.getElementById("zoom-in");
const zoomLabel = document.getElementById("zoom-label");
const planHeading = document.getElementById("plan-heading");
const planRangeLabel = document.getElementById("plan-range-label");
const exportCsvButton = document.getElementById("export-csv");
const exportExcelButton = document.getElementById("export-excel");
const importClipboardButton = document.getElementById("import-clipboard");
const importCsvInput = document.getElementById("import-csv");

initialiseFormDefaults();
initialisePlanDefaults();
initialiseZoom();
syncLayoutVars();
render();

panelToggle.addEventListener("click", () => {
  const collapsed = appShell.classList.toggle("panel-collapsed");
  panelToggle.textContent = collapsed ? "Show panel" : "Hide panel";
  panelToggle.setAttribute("aria-expanded", String(!collapsed));
  const minZoom = getMinimumZoom();
  if (state.zoom < minZoom) {
    state.zoom = minZoom;
  }
  syncLayoutVars();
  syncZoomUi();
  render();
});

window.addEventListener("resize", () => {
  const minZoom = getMinimumZoom();
  if (state.zoom < minZoom) {
    state.zoom = minZoom;
  }
  syncLayoutVars();
  syncZoomUi();
  render();
});

form.addEventListener("submit", (event) => {
  event.preventDefault();

  const name = nameInput.value.trim();
  const start = startInput.value;
  const end = endInput.value;
  const owner = normaliseOwner(ownerInput.value);
  const color = colorInput.value;

  if (!name || !start || !end) {
    return;
  }

  if (toDate(start) > toDate(end)) {
    endInput.setCustomValidity("End date must be after the start date.");
    endInput.reportValidity();
    return;
  }

  endInput.setCustomValidity("");

  state.tasks.push({
    id: crypto.randomUUID(),
    name,
    start,
    end,
    color,
    owner,
    complete: false,
    milestone: false,
  });
  expandPlanToIncludeRange(toDate(start), toDate(end));
  initialisePlanDefaults();

  form.reset();
  colorInput.value = color;
  colorValue.textContent = color.toLowerCase();
  initialiseFormDefaults(start);
  render();
});

colorInput.addEventListener("input", () => {
  colorValue.textContent = colorInput.value.toLowerCase();
});

planForm.addEventListener("submit", (event) => {
  event.preventDefault();

  const nextPlanTitle = planTitleInput.value.trim();
  const nextPlanStart = planStartInput.value;
  const nextPlanEnd = planEndInput.value;

  if (toDate(nextPlanStart) > toDate(nextPlanEnd)) {
    planEndInput.setCustomValidity("Plan end must be after the plan start.");
    planEndInput.reportValidity();
    return;
  }

  const hasOutOfRangeTask = state.tasks.some((task) => {
    return !isWithinRange(toDate(task.start), toDate(task.end), toDate(nextPlanStart), toDate(nextPlanEnd));
  });

  if (hasOutOfRangeTask) {
    planEndInput.setCustomValidity("Expand the plan range so it still includes every task.");
    planEndInput.reportValidity();
    return;
  }

  planEndInput.setCustomValidity("");
  state.planTitle = nextPlanTitle;
  state.planStart = nextPlanStart;
  state.planEnd = nextPlanEnd;
  state.detailLevel = detailLevelInput.value;
  state.showCompactDates = showCompactDatesInput.checked;
  syncCompactDateControl();
  render();
});

detailLevelInput.addEventListener("change", () => {
  state.detailLevel = detailLevelInput.value;
  if (state.detailLevel === "summary" || state.detailLevel === "compressed") {
    state.showCompactDates = true;
  }
  syncCompactDateControl();
  render();
});

showCompactDatesInput.addEventListener("change", () => {
  state.showCompactDates = showCompactDatesInput.checked;
  render();
});

zoomOutButton.addEventListener("click", () => {
  updateZoom(-10);
});

zoomInButton.addEventListener("click", () => {
  updateZoom(10);
});

exportCsvButton.addEventListener("click", () => {
  exportTasksToCsv();
});

exportExcelButton.addEventListener("click", () => {
  exportTasksToExcel();
});

importClipboardButton.addEventListener("click", async () => {
  try {
    const text = await navigator.clipboard.readText();
    if (!text.trim()) {
      window.alert("Clipboard is empty.");
      return;
    }
    importTasksFromCsv(text);
  } catch (error) {
    window.alert("Clipboard import is unavailable. Your browser may require clipboard permission.");
  }
});

importCsvInput.addEventListener("change", async (event) => {
  const [file] = event.target.files || [];
  if (!file) {
    return;
  }

  try {
    const text = await file.text();
    importTasksFromCsv(text);
  } catch (error) {
    window.alert("Could not import that CSV file.");
  } finally {
    importCsvInput.value = "";
  }
});

function initialiseFormDefaults(baseDate = formatInputDate(new Date())) {
  const startDate = toDate(baseDate);
  const endDate = addDays(startDate, 4);

  startInput.value = formatInputDate(startDate);
  endInput.value = formatInputDate(endDate);
  colorValue.textContent = colorInput.value.toLowerCase();
}

function initialisePlanDefaults() {
  planTitleInput.value = state.planTitle;
  planStartInput.value = state.planStart;
  planEndInput.value = state.planEnd;
  detailLevelInput.value = state.detailLevel;
  showCompactDatesInput.checked = state.showCompactDates;
}

function initialiseZoom() {
  syncCompactDateControl();
  syncZoomUi();
}

function render() {
  if (!state.tasks.length) {
    board.innerHTML = "";
    board.appendChild(emptyState);
    syncPlanSummary();
    return;
  }

  const range = getTimelineRange();
  const timelineMode = getTimelineMode(range);
  const scale = buildTimelineScale(range.start, range.end, timelineMode);
  const grid = document.createElement("div");
  grid.className = "gantt-grid";
  if (shouldShowCompactDates()) {
    grid.classList.add("compact-dates");
  }
  if (state.detailLevel === "compressed") {
    grid.classList.add("compressed-view");
  }
  if (state.detailLevel === "summary") {
    grid.classList.add("summary-view");
  }
  grid.style.setProperty("--day-count", scale.length);

  grid.appendChild(renderHeader(scale, timelineMode));
  grid.appendChild(renderLabelResizer());

  state.tasks.forEach((task, index) => {
    grid.appendChild(renderTaskRow(task, scale, range.start, timelineMode, index));
  });

  board.innerHTML = "";
  board.appendChild(grid);
  syncPlanHeading();
  syncPlanSummary();
}

function renderHeader(scale, timelineMode) {
  const header = document.createElement("div");
  header.className = "timeline-header";
  if (timelineMode === "month") {
    header.classList.add("month-mode");
  }

  const cornerCell = document.createElement("div");
  cornerCell.className = `corner-cell${timelineMode === "month" ? " month-corner" : ""}`;
  cornerCell.innerHTML = `
    <div class="eyebrow">Tasks</div>
    <h2>Drag to edit dates</h2>
    <div class="timeline-subtitle">${scale.length} ${timelineMode} slots</div>
  `;
  header.appendChild(cornerCell);

  if (shouldShowCompactDates()) {
    header.appendChild(renderCompactDateHeaderCell("Start", "start"));
    header.appendChild(renderCompactDateHeaderCell("End", "end"));
  }

  if (timelineMode === "month") {
    renderMonthBands(header, scale);
    return header;
  }

  scale.forEach((slot) => {
    const cell = document.createElement("div");
    cell.className = "timeline-day";

    if (timelineMode === "day" && isWeekend(slot.start)) {
      cell.classList.add("weekend");
    }

    cell.innerHTML = getTimelineCellMarkup(slot, timelineMode);
    header.appendChild(cell);
  });

  return header;
}

function renderMonthBands(header, scale) {
  const yearGroups = groupMonthsByYear(scale);
  const timelineStartColumn = getTimelineGridStartColumn();

  yearGroups.forEach((group, index) => {
    const yearCell = document.createElement("div");
    yearCell.className = `timeline-year-band ${getYearShadeClass(index)}`;
    yearCell.style.gridColumn = `${timelineStartColumn + group.startIndex} / span ${group.months.length}`;
    yearCell.style.gridRow = "1";
    yearCell.textContent = String(group.year);
    header.appendChild(yearCell);
  });

  scale.forEach((slot, index) => {
    const monthCell = document.createElement("div");
    monthCell.className = `timeline-month-band ${getYearShadeClass(getYearGroupIndex(yearGroups, slot.start.getFullYear()))}`;
    if (shouldRotateMonthLabel()) {
      monthCell.classList.add("rotated");
    }
    monthCell.style.gridColumn = `${timelineStartColumn + index}`;
    monthCell.style.gridRow = "2";
    monthCell.innerHTML = `
      <div class="timeline-month">${slot.start.toLocaleString("en-GB", { month: "short" })}</div>
    `;
    header.appendChild(monthCell);
  });
}

function renderCompactDateHeaderCell(text, field) {
  const cell = document.createElement("div");
  cell.className = "compact-date-header";
  cell.textContent = text;
  const resizer = document.createElement("div");
  resizer.className = "compact-date-resizer";
  resizer.setAttribute("aria-label", `Resize ${field} date column`);
  resizer.addEventListener("pointerdown", (event) => handleCompactDateResizeStart(event, field));
  cell.appendChild(resizer);
  return cell;
}

function renderCompactDateCell(task, field) {
  const cell = document.createElement("button");
  cell.type = "button";
  cell.className = "compact-date-cell";
  cell.textContent = formatCompactDate(task[field], field);
  cell.title = `Edit ${field} date`;
  cell.addEventListener("click", () => {
    state.editingTaskId = task.id;
    state.editingTaskDateField = field;
    render();
  });
  return cell;
}

function renderTaskRow(task, scale, timelineStart, timelineMode, index) {
  const row = document.createElement("div");
  row.className = "task-row";
  row.dataset.taskId = task.id;
  row.dataset.taskIndex = String(index);

  if (state.rowDrag?.taskId === task.id) {
    row.classList.add("is-row-dragging");
  }

  const label = document.createElement("div");
  label.className = "task-label";

  const taskInfo = document.createElement("div");
  taskInfo.className = "task-info";
  taskInfo.appendChild(renderTaskTitleRow(task));

  if (state.editingTaskId === task.id) {
    taskInfo.appendChild(renderTaskDateEditor(task));
  } else if (state.detailLevel === "detailed") {
    const dateButton = document.createElement("button");
    dateButton.type = "button";
    dateButton.className = "task-date-button";
    dateButton.textContent = `${formatRange(task.start, task.end)} • ${getDuration(task.start, task.end)} days`;
    dateButton.addEventListener("click", () => {
      state.editingTaskId = task.id;
      render();
    });
    taskInfo.appendChild(dateButton);
  }

  const colorPickerWrap = document.createElement("label");
  colorPickerWrap.className = "task-color-picker";
  colorPickerWrap.title = "Change task color";

  const taskColorInput = document.createElement("input");
  taskColorInput.type = "color";
  taskColorInput.value = task.color;
  taskColorInput.setAttribute("aria-label", `Change color for ${task.name}`);
  taskColorInput.addEventListener("input", (event) => {
    task.color = event.target.value;
    render();
  });

  const ownerBadge = renderTaskOwnerControl(task);
  const reorderHandle = document.createElement("button");
  reorderHandle.type = "button";
  reorderHandle.className = "task-reorder-handle";
  reorderHandle.setAttribute("aria-label", `Move ${task.name} up or down`);
  reorderHandle.title = "Drag to reorder task";
  reorderHandle.textContent = "⋮⋮";
  reorderHandle.addEventListener("pointerdown", (event) => handleTaskReorderStart(event, task.id));

  colorPickerWrap.appendChild(taskColorInput);
  label.append(reorderHandle, taskInfo, ownerBadge, colorPickerWrap);
  row.appendChild(label);

  if (shouldShowCompactDates()) {
    row.appendChild(renderCompactDateCell(task, "start"));
    row.appendChild(renderCompactDateCell(task, "end"));
  }

  scale.forEach((slot) => {
    const cell = document.createElement("div");
    cell.className = "day-cell";
    if (timelineMode === "day" && isWeekend(slot.start)) {
      cell.classList.add("weekend");
    }
    if (timelineMode === "month" || timelineMode === "year") {
      cell.classList.add(getYearShadeClass(getYearBandIndex(scale, slot.start.getFullYear(), timelineMode)));
    }
    row.appendChild(cell);
  });

  const layer = document.createElement("div");
  layer.className = "task-bar-layer";

  const bar = document.createElement("div");
  bar.className = "task-bar";
  if (task.milestone) {
    bar.classList.add("milestone");
  }
  bar.dataset.taskId = task.id;
  bar.style.background = getTaskDisplayColor(task);
  const taskPosition = getTaskPosition(task, scale);
  bar.style.left = task.milestone
    ? `${(taskPosition.centerUnits * getDayWidth()) - 12}px`
    : `${taskPosition.startUnits * getDayWidth()}px`;
  bar.style.width = task.milestone
    ? "24px"
    : `${Math.max(taskPosition.widthUnits * getDayWidth(), 6)}px`;

  const startHandle = document.createElement("div");
  startHandle.className = "resize-handle start";
  startHandle.dataset.dragType = "resize-start";

  const endHandle = document.createElement("div");
  endHandle.className = "resize-handle end";
  endHandle.dataset.dragType = "resize-end";

  const labelText = document.createElement("div");
  labelText.className = "task-bar-label";
  if (task.complete) {
    labelText.classList.add("is-complete");
  }
  labelText.textContent = task.name;
  labelText.hidden = shouldHideTaskLabel(task, scale, timelineMode);

  if (!task.milestone) {
    bar.append(startHandle, labelText, endHandle);
  } else {
    bar.append(labelText);
  }
  bar.addEventListener("pointerdown", (event) => handlePointerDown(event, task.id));
  layer.appendChild(bar);
  row.appendChild(layer);

  return row;
}

function handlePointerDown(event, taskId) {
  const target = event.target;
  const dayWidth = getDayWidth();

  const dragType = target.dataset.dragType || "move";
  const task = state.tasks.find((entry) => entry.id === taskId);

  if (!task) {
    return;
  }

  state.drag = {
    taskId,
    dragType,
    originX: event.clientX,
    originalStart: task.start,
    originalEnd: task.end,
    dayWidth,
    timelineMode: getTimelineMode(getTimelineRange()),
  };

  window.addEventListener("pointermove", handlePointerMove);
  window.addEventListener("pointerup", handlePointerUp);
  window.addEventListener("pointercancel", handlePointerUp);
}

function handlePointerMove(event) {
  if (!state.drag) {
    return;
  }

  const stepDelta = Math.round((event.clientX - state.drag.originX) / state.drag.dayWidth);
  const task = state.tasks.find((entry) => entry.id === state.drag.taskId);

  if (!task) {
    return;
  }

  const originalStart = toDate(state.drag.originalStart);
  const originalEnd = toDate(state.drag.originalEnd);

  if (state.drag.dragType === "move") {
    const nextStart = addTimelineUnits(originalStart, stepDelta, state.drag.timelineMode);
    const nextEnd = task.milestone ? nextStart : addTimelineUnits(originalEnd, stepDelta, state.drag.timelineMode);
    if (isWithinPlan(nextStart, nextEnd)) {
      task.start = formatInputDate(nextStart);
      task.end = formatInputDate(nextEnd);
    }
  }

  if (state.drag.dragType === "resize-start") {
    const nextStart = addTimelineUnits(originalStart, stepDelta, state.drag.timelineMode);
    if (nextStart <= toDate(task.end) && isWithinPlan(nextStart, toDate(task.end))) {
      task.start = formatInputDate(nextStart);
    }
  }

  if (state.drag.dragType === "resize-end") {
    const nextEnd = addTimelineUnits(originalEnd, stepDelta, state.drag.timelineMode);
    if (nextEnd >= toDate(task.start) && isWithinPlan(toDate(task.start), nextEnd)) {
      task.end = formatInputDate(nextEnd);
    }
  }

  render();
}

function handlePointerUp(event) {
  state.drag = null;
  window.removeEventListener("pointermove", handlePointerMove);
  window.removeEventListener("pointerup", handlePointerUp);
  window.removeEventListener("pointercancel", handlePointerUp);
}

function handleTaskReorderStart(event, taskId) {
  event.preventDefault();
  event.stopPropagation();

  const taskIndex = state.tasks.findIndex((task) => task.id === taskId);
  if (taskIndex === -1) {
    return;
  }

  state.rowDrag = {
    taskId,
    currentIndex: taskIndex,
  };

  window.addEventListener("pointermove", handleTaskReorderMove);
  window.addEventListener("pointerup", handleTaskReorderEnd);
  window.addEventListener("pointercancel", handleTaskReorderEnd);
  render();
}

function handleTaskReorderMove(event) {
  if (!state.rowDrag) {
    return;
  }

  const nextIndex = getTaskIndexFromPointer(event.clientY);
  if (nextIndex === -1 || nextIndex === state.rowDrag.currentIndex) {
    return;
  }

  moveTask(state.rowDrag.taskId, nextIndex);
  state.rowDrag.currentIndex = nextIndex;
  render();
}

function handleTaskReorderEnd() {
  state.rowDrag = null;
  window.removeEventListener("pointermove", handleTaskReorderMove);
  window.removeEventListener("pointerup", handleTaskReorderEnd);
  window.removeEventListener("pointercancel", handleTaskReorderEnd);
  render();
}

function getTaskIndexFromPointer(clientY) {
  const rows = [...board.querySelectorAll(".task-row")];
  if (!rows.length) {
    return -1;
  }

  for (let index = 0; index < rows.length; index += 1) {
    const rect = rows[index].getBoundingClientRect();
    const midpoint = rect.top + (rect.height / 2);
    if (clientY < midpoint) {
      return index;
    }
  }

  return rows.length - 1;
}

function moveTask(taskId, nextIndex) {
  const currentIndex = state.tasks.findIndex((task) => task.id === taskId);
  if (currentIndex === -1 || currentIndex === nextIndex) {
    return;
  }

  const tasks = [...state.tasks];
  const [movedTask] = tasks.splice(currentIndex, 1);
  tasks.splice(nextIndex, 0, movedTask);
  state.tasks = tasks;
}

function getTimelineRange() {
  return {
    start: toDate(state.planStart),
    end: toDate(state.planEnd),
  };
}

function getDuration(start, end) {
  return diffDays(toDate(start), toDate(end)) + 1;
}

function formatRange(start, end) {
  const formatOptions = { day: "numeric", month: "short", year: "numeric" };
  return `${toDate(start).toLocaleDateString("en-GB", formatOptions)} - ${toDate(end).toLocaleDateString("en-GB", formatOptions)}`;
}

function formatShortRange(start, end) {
  const options = { day: "2-digit", month: "2-digit", year: "2-digit" };
  return `${toDate(start).toLocaleDateString("en-GB", options)}-${toDate(end).toLocaleDateString("en-GB", options)}`;
}

function formatCompactDate(value, field) {
  const width = field === "start" ? state.compactStartColumnWidth : state.compactEndColumnWidth;
  const options = width < 62
    ? { day: "2-digit", month: "2-digit" }
    : { day: "2-digit", month: "2-digit", year: "2-digit" };
  return toDate(value).toLocaleDateString("en-GB", options);
}

function diffDays(start, end) {
  return Math.round((stripTime(end) - stripTime(start)) / DAY_MS);
}

function addDays(date, amount) {
  const result = new Date(date);
  result.setDate(result.getDate() + amount);
  return result;
}

function getStartOfWeek(date) {
  const result = new Date(date);
  const day = result.getDay();
  const diffToMonday = day === 0 ? -6 : 1 - day;
  result.setDate(result.getDate() + diffToMonday);
  return result;
}

function toDate(value) {
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function formatInputDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6;
}

function stripTime(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
}

function getDayWidth() {
  return Number.parseFloat(
    getComputedStyle(document.documentElement).getPropertyValue("--day-width")
  );
}

function updateZoom(delta) {
  state.zoom = clamp(state.zoom + delta, getMinimumZoom(), 360);
  syncZoomUi();
  render();
}

function syncLayoutVars() {
  document.documentElement.style.setProperty("--label-column", `${state.labelColumnWidth}px`);
  document.documentElement.style.setProperty("--compact-start-column", `${state.compactStartColumnWidth}px`);
  document.documentElement.style.setProperty("--compact-end-column", `${state.compactEndColumnWidth}px`);
}

function syncCompactDateControl() {
  showCompactDatesInput.checked = state.showCompactDates;
}

function shouldShowCompactDates() {
  return state.showCompactDates && (state.detailLevel === "summary" || state.detailLevel === "compressed");
}

function syncZoomUi() {
  const timelineMode = getTimelineMode(getTimelineRange());
  let width = 42 * (state.zoom / 100);
  if (timelineMode === "week") {
    width *= 2;
  }
  if (timelineMode === "month") {
    width *= 2;
  }
  document.documentElement.style.setProperty("--day-width", `${width}px`);
  zoomLabel.textContent = `${state.zoom}%`;
  const minZoom = getMinimumZoom();
  zoomOutButton.disabled = state.zoom <= minZoom;
  zoomInButton.disabled = state.zoom >= 360;
}

function syncPlanSummary() {
  planRangeLabel.textContent = `Plan: ${formatRange(state.planStart, state.planEnd)}`;
}

function syncPlanHeading() {
  planHeading.textContent = state.planTitle || "Schedule Overview";
}

function isWithinPlan(start, end) {
  const planStart = toDate(state.planStart);
  const planEnd = toDate(state.planEnd);
  return isWithinRange(start, end, planStart, planEnd);
}

function isWithinRange(start, end, rangeStart, rangeEnd) {
  return start >= rangeStart && end <= rangeEnd;
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function getTimelineMode(range = getTimelineRange()) {
  const availableWidth = Math.max(window.innerWidth - getPanelAllowance(), 240);
  const totalDays = Math.max(getDuration(state.planStart, state.planEnd), 1);
  const totalWeeks = Math.max(Math.ceil(totalDays / 7), 1);
  const totalMonths = Math.max(diffMonths(range.start, range.end) + 1, 1);
  const totalYears = Math.max(range.end.getFullYear() - range.start.getFullYear() + 1, 1);
  const daySlotWidth = availableWidth / totalDays;
  const weekSlotWidth = availableWidth / totalWeeks;
  const monthSlotWidth = availableWidth / totalMonths;
  const yearSlotWidth = availableWidth / totalYears;

  if (state.zoom >= 150 || daySlotWidth >= 26) {
    return "day";
  }
  if (state.zoom >= 80 || weekSlotWidth >= 54) {
    return "week";
  }
  if (state.zoom >= 30 || monthSlotWidth >= 44) {
    return "month";
  }
  if (yearSlotWidth >= 36) {
    return "year";
  }
  return "year";
}

function getMinimumZoom() {
  const planDays = getDuration(state.planStart, state.planEnd);
  const availableWidth = Math.max(window.innerWidth - getPanelAllowance(), 240);
  const targetDayWidth = availableWidth / Math.max(planDays, 1);
  const zoom = Math.floor((targetDayWidth / 42) * 100);
  return clamp(zoom, 8, 360);
}

function exportTasksToCsv() {
  const header = ["name", "start", "end", "owner", "color", "complete", "milestone"];
  const metadataRow = ["#plan_title", escapeCsvValue(state.planTitle || "")];
  const rows = state.tasks.map((task) => [
    escapeCsvValue(task.name),
    task.start,
    task.end,
    task.owner || "",
    task.color,
    task.complete ? "true" : "false",
    task.milestone ? "true" : "false",
  ]);
  const csv = [metadataRow.join(","), header.join(","), ...rows.map((row) => row.join(","))].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "gantt-tasks.csv";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function exportTasksToExcel() {
  const range = getTimelineRange();
  const timelineMode = getTimelineMode(range);
  const scale = buildTimelineScale(range.start, range.end, timelineMode);
  const timelineHeaders = scale.map((slot) => getExcelTimelineLabel(slot, timelineMode));
  const html = buildExcelHtml(scale, timelineHeaders, timelineMode);
  const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "gantt-plan.xls";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function buildExcelHtml(scale, timelineHeaders, timelineMode) {
  const timelineHeaderCells = timelineHeaders
    .map((label) => `<th class="timeline-col">${escapeHtml(label)}</th>`)
    .join("");

  const taskRows = state.tasks.map((task) => {
    const color = getTaskDisplayColor(task);
    const timelineCells = scale.map((slot) => {
      const active = rangesOverlap(toDate(task.start), toDate(task.end), slot.start, slot.end);
      const style = active
        ? `background:${color};color:${getContrastTextColor(color)};`
        : "background:#fbf7f0;color:#74665b;";
      const content = active ? "■" : "";
      return `<td class="timeline-cell" style="${style}">${content}</td>`;
    }).join("");

    return `
      <tr>
        <td class="text-cell">${escapeHtml(task.name)}</td>
        <td class="date-cell">${escapeHtml(formatExcelDate(task.start))}</td>
        <td class="date-cell">${escapeHtml(formatExcelDate(task.end))}</td>
        <td class="complete-cell">${task.complete ? "☑" : "☐"}</td>
        <td class="owner-cell">${escapeHtml(task.owner || "")}</td>
        ${timelineCells}
      </tr>
    `;
  }).join("");

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: Arial, sans-serif; }
    table { border-collapse: collapse; }
    th, td { border: 1px solid #d8d1c7; padding: 4px 6px; }
    th { background: #efe6d9; font-weight: 700; }
    .text-cell { min-width: 220px; }
    .date-cell { min-width: 86px; }
    .complete-cell { min-width: 54px; text-align: center; font-size: 14px; }
    .owner-cell { min-width: 56px; text-align: center; text-transform: uppercase; }
    .timeline-col { min-width: 28px; text-align: center; font-size: 11px; }
    .timeline-cell { min-width: 28px; height: 24px; text-align: center; font-size: 11px; }
    .sheet-title { background: #e7dccd; text-align: left; }
  </style>
</head>
<body>
  <table>
    <tr>
      <th class="sheet-title" colspan="${5 + scale.length}">
        ${escapeHtml(state.planTitle || "Schedule Overview")} - ${escapeHtml(timelineMode.toUpperCase())} timeline - ${escapeHtml(formatRange(state.planStart, state.planEnd))}
      </th>
    </tr>
    <tr>
      <th>Task</th>
      <th>Start Date</th>
      <th>End Date</th>
      <th>Complete</th>
      <th>Owner</th>
      ${timelineHeaderCells}
    </tr>
    ${taskRows}
  </table>
</body>
</html>`;
}

function importTasksFromCsv(text) {
  const rows = parseCsv(text);
  if (rows.length < 2) {
    window.alert("The CSV file is empty.");
    return;
  }

  let rowOffset = 0;
  let importedPlanTitle = "";

  if ((rows[0][0] || "").trim().toLowerCase() === "#plan_title") {
    importedPlanTitle = (rows[0][1] || "").trim();
    rowOffset = 1;
  }

  if (rows.length < rowOffset + 2) {
    window.alert("The CSV file is empty.");
    return;
  }

  const headers = rows[rowOffset].map((header) => header.trim().toLowerCase());
  const nameIndex = headers.indexOf("name");
  const startIndex = headers.indexOf("start");
  const endIndex = headers.indexOf("end");
  const ownerIndex = headers.indexOf("owner");
  const colorIndex = headers.indexOf("color");
  const completeIndex = headers.indexOf("complete");
  const milestoneIndex = headers.indexOf("milestone");

  if (nameIndex === -1 || startIndex === -1 || endIndex === -1 || colorIndex === -1) {
    window.alert("CSV must contain name, start, end, and color columns.");
    return;
  }

  const tasks = rows.slice(rowOffset + 1)
    .filter((row) => row.some((value) => value.trim() !== ""))
    .map((row) => ({
      id: crypto.randomUUID(),
      name: (row[nameIndex] || "").trim(),
      start: (row[startIndex] || "").trim(),
      end: (row[endIndex] || "").trim(),
      owner: normaliseOwner(row[ownerIndex] || ""),
      color: normaliseColor((row[colorIndex] || "").trim()),
      complete: parseCompleteValue(row[completeIndex]),
      milestone: parseCompleteValue(row[milestoneIndex]),
    }));

  if (!tasks.length) {
    window.alert("No valid task rows were found in the CSV.");
    return;
  }

  const hasInvalidTask = tasks.some((task) => {
    return !task.name
      || !isIsoDate(task.start)
      || !isIsoDate(task.end)
      || toDate(task.start) > toDate(task.end)
      || (task.milestone && task.start !== task.end);
  });

  if (hasInvalidTask) {
    window.alert("Each task must have a name, valid start/end dates, and milestones must use the same start and end date.");
    return;
  }

  state.tasks = tasks;
  state.planTitle = importedPlanTitle;
  state.editingTaskId = null;
  syncPlanToTasks();
  initialisePlanDefaults();
  const minZoom = getMinimumZoom();
  if (state.zoom < minZoom) {
    state.zoom = minZoom;
  }
  syncZoomUi();
  render();
}

function syncPlanToTasks() {
  const starts = state.tasks.map((task) => toDate(task.start).getTime());
  const ends = state.tasks.map((task) => toDate(task.end).getTime());
  state.planStart = formatInputDate(new Date(Math.min(...starts)));
  state.planEnd = formatInputDate(new Date(Math.max(...ends)));
}

function expandPlanToIncludeRange(start, end) {
  const currentPlanStart = toDate(state.planStart);
  const currentPlanEnd = toDate(state.planEnd);

  if (start < currentPlanStart) {
    state.planStart = formatInputDate(start);
  }

  if (end > currentPlanEnd) {
    state.planEnd = formatInputDate(end);
  }
}

function getExcelTimelineLabel(slot, timelineMode) {
  if (timelineMode === "year") {
    return String(slot.start.getFullYear());
  }
  if (timelineMode === "month") {
    return slot.start.toLocaleDateString("en-GB", { month: "short", year: "2-digit" });
  }
  if (timelineMode === "week") {
    return slot.start.toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit" });
  }
  return slot.start.toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit" });
}

function formatExcelDate(value) {
  return toDate(value).toLocaleDateString("en-GB");
}

function renderTaskDateEditor(task) {
  const wrapper = document.createElement("div");
  wrapper.className = "task-date-editor";

  const startField = document.createElement("input");
  startField.type = "date";
  startField.value = task.start;

  const endField = document.createElement("input");
  endField.type = "date";
  endField.value = task.end;

  const milestoneLabel = document.createElement("label");
  milestoneLabel.className = "task-milestone-toggle";

  const milestoneField = document.createElement("input");
  milestoneField.type = "checkbox";
  milestoneField.checked = Boolean(task.milestone);

  const milestoneText = document.createElement("span");
  milestoneText.textContent = "Milestone";

  milestoneField.addEventListener("change", () => {
    if (milestoneField.checked) {
      endField.value = startField.value;
      endField.disabled = true;
    } else {
      endField.disabled = false;
    }
  });

  milestoneLabel.append(milestoneField, milestoneText);

  if (milestoneField.checked) {
    endField.disabled = true;
  }

  const actions = document.createElement("div");
  actions.className = "task-date-actions";

  const saveButton = document.createElement("button");
  saveButton.type = "button";
  saveButton.className = "mini-button";
  saveButton.textContent = "Save";
  saveButton.addEventListener("click", () => {
    const nextStart = toDate(startField.value);
    const nextEnd = toDate(milestoneField.checked ? startField.value : endField.value);

    if (nextStart > nextEnd) {
      endField.setCustomValidity("End date must be after the start date.");
      endField.reportValidity();
      return;
    }

    endField.setCustomValidity("");
    task.start = startField.value;
    task.end = milestoneField.checked ? startField.value : endField.value;
    task.milestone = milestoneField.checked;
    expandPlanToIncludeRange(nextStart, nextEnd);
    initialisePlanDefaults();
    state.editingTaskId = null;
    state.editingTaskDateField = null;
    render();
  });

  const cancelButton = document.createElement("button");
  cancelButton.type = "button";
  cancelButton.className = "mini-button";
  cancelButton.textContent = "Cancel";
  cancelButton.addEventListener("click", () => {
    state.editingTaskId = null;
    state.editingTaskDateField = null;
    render();
  });

  actions.append(saveButton, cancelButton);
  wrapper.append(startField, endField, milestoneLabel, actions);
  queueMicrotask(() => {
    if (state.editingTaskDateField === "end") {
      endField.focus();
      return;
    }
    startField.focus();
  });
  return wrapper;
}

function renderTaskTitleRow(task) {
  const titleRow = document.createElement("div");
  titleRow.className = "task-title-row";

  const completeButton = document.createElement("button");
  completeButton.type = "button";
  completeButton.className = `task-complete-button${task.complete ? " is-complete" : ""}`;
  completeButton.textContent = task.complete ? "✓" : "";
  completeButton.setAttribute("aria-label", task.complete ? `Mark ${task.name} incomplete` : `Mark ${task.name} complete`);
  completeButton.addEventListener("click", () => {
    task.complete = !task.complete;
    render();
  });
  titleRow.appendChild(completeButton);

  if (state.editingTaskNameId === task.id) {
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.className = "task-name-editor";
    nameInput.value = task.name;
    nameInput.maxLength = 50;
    nameInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        saveTaskName(task, nameInput.value);
      }
      if (event.key === "Escape") {
        state.editingTaskNameId = null;
        render();
      }
    });
    nameInput.addEventListener("blur", () => {
      saveTaskName(task, nameInput.value);
    });
    titleRow.appendChild(nameInput);
    queueMicrotask(() => nameInput.focus());
  } else {
    const titleButton = document.createElement("button");
    titleButton.type = "button";
    titleButton.className = "task-name-button";
    if (task.complete) {
      titleButton.classList.add("is-complete");
    }
    titleButton.textContent = task.name;
    titleButton.title = state.detailLevel === "summary" ? formatShortRange(task.start, task.end) : task.name;
    titleButton.addEventListener("click", () => {
      state.editingTaskNameId = task.id;
      render();
    });
    titleRow.appendChild(titleButton);
  }

  const removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.className = "task-remove-button";
  removeButton.textContent = "x";
  removeButton.setAttribute("aria-label", `Remove ${task.name}`);
  removeButton.addEventListener("click", () => {
    removeTask(task.id);
  });
  titleRow.appendChild(removeButton);

  return titleRow;
}

function renderTaskOwnerControl(task) {
  if (state.editingTaskOwnerId === task.id) {
    const ownerEditor = document.createElement("input");
    ownerEditor.type = "text";
    ownerEditor.className = "task-owner-editor";
    ownerEditor.value = task.owner || "";
    ownerEditor.maxLength = 4;
    ownerEditor.setAttribute("aria-label", `Edit owner for ${task.name}`);
    ownerEditor.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        saveTaskOwner(task, ownerEditor.value);
      }
      if (event.key === "Escape") {
        state.editingTaskOwnerId = null;
        render();
      }
    });
    ownerEditor.addEventListener("blur", () => {
      saveTaskOwner(task, ownerEditor.value);
    });
    queueMicrotask(() => ownerEditor.focus());
    return ownerEditor;
  }

  const ownerBadge = document.createElement("button");
  ownerBadge.type = "button";
  ownerBadge.className = "task-owner-badge";
  ownerBadge.textContent = task.owner || "--";
  ownerBadge.title = task.owner ? `Owner ${task.owner}` : "Set owner";
  ownerBadge.addEventListener("click", () => {
    state.editingTaskOwnerId = task.id;
    render();
  });
  return ownerBadge;
}

function saveTaskName(task, nextName) {
  const trimmed = nextName.trim();
  if (trimmed) {
    task.name = trimmed;
  }
  state.editingTaskNameId = null;
  render();
}

function saveTaskOwner(task, nextOwner) {
  task.owner = normaliseOwner(nextOwner);
  state.editingTaskOwnerId = null;
  render();
}

function removeTask(taskId) {
  state.tasks = state.tasks.filter((task) => task.id !== taskId);
  if (state.editingTaskId === taskId) {
    state.editingTaskId = null;
    state.editingTaskDateField = null;
  }
  if (state.editingTaskNameId === taskId) {
    state.editingTaskNameId = null;
  }
  if (state.editingTaskOwnerId === taskId) {
    state.editingTaskOwnerId = null;
  }

  if (state.tasks.length) {
    syncPlanToTasks();
    initialisePlanDefaults();
  }

  render();
}

function renderLabelResizer() {
  const resizer = document.createElement("div");
  resizer.className = "label-resizer";
  resizer.setAttribute("aria-label", "Resize task title column");
  resizer.addEventListener("pointerdown", handleLabelResizeStart);
  return resizer;
}

function handleLabelResizeStart(event) {
  event.preventDefault();
  const resizer = event.currentTarget;
  const originX = event.clientX;
  const originWidth = state.labelColumnWidth;

  resizer.classList.add("is-dragging");

  function handleMove(moveEvent) {
    state.labelColumnWidth = clamp(originWidth + (moveEvent.clientX - originX), 220, 640);
    syncLayoutVars();
    render();
  }

  function handleUp() {
    resizer.classList.remove("is-dragging");
    window.removeEventListener("pointermove", handleMove);
    window.removeEventListener("pointerup", handleUp);
    window.removeEventListener("pointercancel", handleUp);
  }

  window.addEventListener("pointermove", handleMove);
  window.addEventListener("pointerup", handleUp);
  window.addEventListener("pointercancel", handleUp);
}

function handleCompactDateResizeStart(event, field) {
  event.preventDefault();
  event.stopPropagation();
  const resizer = event.currentTarget;
  const originX = event.clientX;
  const originWidth = field === "start" ? state.compactStartColumnWidth : state.compactEndColumnWidth;

  resizer.classList.add("is-dragging");

  function handleMove(moveEvent) {
    const nextWidth = clamp(originWidth + (moveEvent.clientX - originX), 44, 130);
    if (field === "start") {
      state.compactStartColumnWidth = nextWidth;
    } else {
      state.compactEndColumnWidth = nextWidth;
    }
    syncLayoutVars();
    render();
  }

  function handleUp() {
    resizer.classList.remove("is-dragging");
    window.removeEventListener("pointermove", handleMove);
    window.removeEventListener("pointerup", handleUp);
    window.removeEventListener("pointercancel", handleUp);
  }

  window.addEventListener("pointermove", handleMove);
  window.addEventListener("pointerup", handleUp);
  window.addEventListener("pointercancel", handleUp);
}

function shouldHideTaskLabel(task, scale, timelineMode) {
  const barWidth = task.milestone
    ? 24
    : getTaskPosition(task, scale).widthUnits * getDayWidth();
  const estimatedTextWidth = task.name.length * 7.2;
  return barWidth < estimatedTextWidth + 56;
}

function getTaskPosition(task, scale) {
  const startUnits = getDateUnitsWithinScale(toDate(task.start), scale);
  const endUnits = getDateUnitsWithinScale(addDays(toDate(task.end), 1), scale, true);
  return {
    startUnits,
    widthUnits: Math.max(endUnits - startUnits, 0),
    centerUnits: (startUnits + endUnits) / 2,
  };
}

function getDateUnitsWithinScale(date, scale, endExclusive = false) {
  if (!scale.length) {
    return 0;
  }

  const firstSlotStart = scale[0].start;
  const lastSlotEndExclusive = addDays(scale[scale.length - 1].end, 1);

  if (date <= firstSlotStart) {
    return 0;
  }

  if (date >= lastSlotEndExclusive) {
    return scale.length;
  }

  for (let index = 0; index < scale.length; index += 1) {
    const slot = scale[index];
    const slotStart = slot.start;
    const slotEndExclusive = addDays(slot.end, 1);

    if (date >= slotStart && date < slotEndExclusive) {
      const slotDays = diffDays(slotStart, slotEndExclusive);
      const dayOffset = diffDays(slotStart, date);
      return index + (dayOffset / Math.max(slotDays, 1));
    }

    if (endExclusive && date.getTime() === slotEndExclusive.getTime()) {
      return index + 1;
    }
  }

  return scale.length;
}

function buildTimelineScale(start, end, timelineMode) {
  if (timelineMode === "year") {
    return buildYearArray(start, end);
  }
  if (timelineMode === "month") {
    return buildMonthArray(start, end);
  }
  if (timelineMode === "week") {
    return buildWeekArray(start, end);
  }
  return buildDayArray(start, end);
}

function buildDayArray(start, end) {
  const dates = [];
  let cursor = new Date(start);
  while (cursor <= end) {
    dates.push({ start: new Date(cursor), end: new Date(cursor) });
    cursor = addDays(cursor, 1);
  }
  return dates;
}

function buildWeekArray(start, end) {
  const weeks = [];
  let cursor = getStartOfWeek(start);
  while (cursor <= end) {
    const weekEnd = addDays(cursor, 6);
    weeks.push({ start: new Date(cursor), end: weekEnd > end ? new Date(end) : weekEnd });
    cursor = addDays(cursor, 7);
  }
  return weeks;
}

function buildMonthArray(start, end) {
  const months = [];
  let cursor = new Date(start.getFullYear(), start.getMonth(), 1);
  while (cursor <= end) {
    const monthStart = cursor < start ? new Date(start) : new Date(cursor);
    const monthEnd = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0);
    months.push({ start: monthStart, end: monthEnd > end ? new Date(end) : monthEnd });
    cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
  }
  return months;
}

function buildYearArray(start, end) {
  const years = [];
  let year = start.getFullYear();

  while (year <= end.getFullYear()) {
    const yearStart = new Date(year, 0, 1);
    const yearEnd = new Date(year, 11, 31);
    years.push({
      start: year === start.getFullYear() ? new Date(start) : yearStart,
      end: year === end.getFullYear() ? new Date(end) : yearEnd,
    });
    year += 1;
  }

  return years;
}

function getTimelineCellMarkup(slot, timelineMode) {
  if (timelineMode === "year") {
    return `
      <div class="timeline-month">Year</div>
      <div class="timeline-date">${slot.start.getFullYear()}</div>
    `;
  }

  if (timelineMode === "month") {
    return `
      <div class="timeline-month">${slot.start.toLocaleString("en-GB", { month: "short" })}</div>
    `;
  }

  if (timelineMode === "week") {
    return `
      <div class="timeline-month">Week</div>
      <div class="timeline-date">${slot.start.getDate()} ${slot.start.toLocaleString("en-GB", { month: "short" })}</div>
    `;
  }

  return `
    <div class="timeline-month">${slot.start.toLocaleString("en-GB", { month: "short" })}</div>
    <div class="timeline-date">${slot.start.getDate()}</div>
  `;
}

function getOffsetUnits(timelineStart, taskStart, timelineMode) {
  if (timelineMode === "year") {
    return taskStart.getFullYear() - timelineStart.getFullYear();
  }
  if (timelineMode === "month") {
    return diffMonths(timelineStart, taskStart);
  }
  if (timelineMode === "week") {
    return Math.floor(diffDays(timelineStart, taskStart) / 7);
  }
  return diffDays(timelineStart, taskStart);
}

function getSpanUnits(taskStart, taskEnd, timelineMode) {
  if (timelineMode === "year") {
    return toDate(taskEnd).getFullYear() - toDate(taskStart).getFullYear() + 1;
  }
  if (timelineMode === "month") {
    return diffMonths(toDate(taskStart), toDate(taskEnd)) + 1;
  }
  if (timelineMode === "week") {
    return Math.floor(diffDays(toDate(taskStart), toDate(taskEnd)) / 7) + 1;
  }
  return getDuration(taskStart, taskEnd);
}

function diffMonths(start, end) {
  return (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth());
}

function getPanelAllowance() {
  return appShell.classList.contains("panel-collapsed") ? 120 : 420;
}

function groupMonthsByYear(scale) {
  const groups = [];

  scale.forEach((slot, index) => {
    const year = slot.start.getFullYear();
    const lastGroup = groups[groups.length - 1];

    if (!lastGroup || lastGroup.year !== year) {
      groups.push({
        year,
        startIndex: index,
        months: [slot],
      });
      return;
    }

    lastGroup.months.push(slot);
  });

  return groups;
}

function groupYears(scale) {
  return scale.map((slot, index) => ({
    year: slot.start.getFullYear(),
    startColumn: index + 2,
  }));
}

function getYearGroupIndex(groups, year) {
  return groups.findIndex((group) => group.year === year);
}

function getYearBandIndex(scale, year, timelineMode) {
  if (timelineMode === "year") {
    return getYearGroupIndex(groupYears(scale), year);
  }
  return getYearGroupIndex(groupMonthsByYear(scale), year);
}

function getTimelineGridStartColumn() {
  return shouldShowCompactDates() ? 4 : 2;
}

function getYearShadeClass(index) {
  return index % 2 === 0 ? "year-shade-even" : "year-shade-odd";
}

function shouldRotateMonthLabel() {
  return getDayWidth() < 34;
}

function addTimelineUnits(date, amount, timelineMode) {
  if (timelineMode === "year") {
    return addYears(date, amount);
  }
  if (timelineMode === "month") {
    return addMonths(date, amount);
  }
  if (timelineMode === "week") {
    return addDays(date, amount * 7);
  }
  return addDays(date, amount);
}

function addMonths(date, amount) {
  const result = new Date(date);
  const originalDay = result.getDate();
  result.setDate(1);
  result.setMonth(result.getMonth() + amount);
  const lastDayOfMonth = new Date(result.getFullYear(), result.getMonth() + 1, 0).getDate();
  result.setDate(Math.min(originalDay, lastDayOfMonth));
  return result;
}

function addYears(date, amount) {
  const result = new Date(date);
  const originalMonth = result.getMonth();
  const originalDay = result.getDate();
  result.setDate(1);
  result.setFullYear(result.getFullYear() + amount);
  const lastDayOfMonth = new Date(result.getFullYear(), originalMonth + 1, 0).getDate();
  result.setMonth(originalMonth);
  result.setDate(Math.min(originalDay, lastDayOfMonth));
  return result;
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const nextChar = text[index + 1];

    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        value += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(value);
      value = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && nextChar === "\n") {
        index += 1;
      }
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
      continue;
    }

    value += char;
  }

  if (value.length || row.length) {
    row.push(value);
    rows.push(row);
  }

  return rows;
}

function escapeCsvValue(value) {
  if (/[",\n]/.test(value)) {
    return `"${value.replaceAll('"', '""')}"`;
  }
  return value;
}

function isIsoDate(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(value);
}

function normaliseColor(value) {
  return /^#[0-9a-f]{6}$/i.test(value) ? value : "#ff7a59";
}

function normaliseOwner(value) {
  return String(value || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "").slice(0, 4);
}

function parseCompleteValue(value) {
  return String(value || "").trim().toLowerCase() === "true";
}

function getTaskDisplayColor(task) {
  if (!task.complete) {
    return task.color;
  }

  const { r, g, b } = hexToRgb(task.color);
  const luminance = Math.round(r * 0.299 + g * 0.587 + b * 0.114);
  const greyBase = clamp(Math.round(luminance * 0.92), 96, 188);
  return rgbToHex(greyBase, greyBase, greyBase);
}

function hexToRgb(hex) {
  const clean = hex.replace("#", "");
  return {
    r: Number.parseInt(clean.slice(0, 2), 16),
    g: Number.parseInt(clean.slice(2, 4), 16),
    b: Number.parseInt(clean.slice(4, 6), 16),
  };
}

function rgbToHex(r, g, b) {
  return `#${[r, g, b].map((value) => value.toString(16).padStart(2, "0")).join("")}`;
}

function rangesOverlap(startA, endA, startB, endB) {
  return startA <= endB && endA >= startB;
}

function getContrastTextColor(hex) {
  const { r, g, b } = hexToRgb(hex);
  const luminance = (r * 0.299) + (g * 0.587) + (b * 0.114);
  return luminance > 160 ? "#231b16" : "#ffffff";
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
