const DAY_MS = 24 * 60 * 60 * 1000;

const initialTasks = [
  {
    id: crypto.randomUUID(),
    name: "Kickoff",
    start: "2026-03-24",
    end: "2026-05-23",
    color: "#ff7a59",
    owner: "RB",
    complete: false,
    milestone: false,
  },
  {
    id: crypto.randomUUID(),
    name: "Build",
    start: "2026-05-24",
    end: "2026-09-23",
    color: "#0f9d8a",
    owner: "AK",
    complete: false,
    milestone: false,
  },
  {
    id: crypto.randomUUID(),
    name: "Test",
    start: "2026-09-24",
    end: "2026-10-23",
    color: "#3366ff",
    owner: "JT",
    complete: false,
    milestone: false,
  },
  {
    id: crypto.randomUUID(),
    name: "Train",
    start: "2026-10-24",
    end: "2026-11-23",
    color: "#7a56d8",
    owner: "LM",
    complete: false,
    milestone: false,
  },
  {
    id: crypto.randomUUID(),
    name: "Go-live",
    start: "2026-11-24",
    end: "2027-01-23",
    color: "#d9822b",
    owner: "NS",
    complete: false,
    milestone: false,
  },
];

const initialPlanStartDate = new Date(2026, 2, 22);
const initialPlanEndDate = addDays(addMonths(initialPlanStartDate, 12), -1);

const state = {
  tasks: initialTasks,
  planStart: formatInputDate(initialPlanStartDate),
  planEnd: formatInputDate(initialPlanEndDate),
  planTitle: "",
  zoom: 100,
  drag: null,
  labelColumnWidth: 320,
  compactStartColumnWidth: 74,
  compactEndColumnWidth: 74,
  detailLevel: "compressed",
  showCompactDates: true,
  showTaskDates: true,
  showTaskMeta: true,
  showRelativeTimeline: false,
  chartColorScheme: "warm",
  timelineModeOverride: null,
  editingPlanTitle: false,
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
const timelineLabelModeInput = document.getElementById("timeline-label-mode");
const showCompactDatesInput = document.getElementById("show-compact-dates");
const showTaskDatesInput = document.getElementById("show-task-dates");
const showTaskMetaInput = document.getElementById("show-task-meta");
const showRelativeTimelineInput = document.getElementById("show-relative-timeline");
const chartColorSchemeInput = document.getElementById("chart-colour-scheme");
const zoomOutButton = document.getElementById("zoom-out");
const zoomInButton = document.getElementById("zoom-in");
const zoomLabel = document.getElementById("zoom-label");
const timelineModeButtons = [...document.querySelectorAll(".timeline-mode-button")];
const planHeading = document.getElementById("plan-heading");
const planRangeLabel = document.getElementById("plan-range-label");
const exportCsvButton = document.getElementById("export-csv");
const exportExcelButton = document.getElementById("export-excel");
const exportPptButton = document.getElementById("export-ppt");
const copyChartImageButton = document.getElementById("copy-chart-image");
const importClipboardButton = document.getElementById("import-clipboard");
const importCsvInput = document.getElementById("import-csv");

initialiseFormDefaults();
initialisePlanDefaults();
initialiseZoom();
syncLayoutVars();
render();

planHeading.addEventListener("click", () => {
  state.editingPlanTitle = true;
  syncPlanHeading();
});

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
  state.timelineModeOverride = timelineLabelModeInput.value === "auto" ? null : timelineLabelModeInput.value;
  state.showCompactDates = showCompactDatesInput.checked;
  state.showTaskDates = showTaskDatesInput.checked;
  state.showTaskMeta = showTaskMetaInput.checked;
  state.showRelativeTimeline = showRelativeTimelineInput.checked;
  state.chartColorScheme = chartColorSchemeInput.value;
  syncCompactDateControl();
  syncLayoutVars();
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

showTaskDatesInput.addEventListener("change", () => {
  state.showTaskDates = showTaskDatesInput.checked;
  render();
});

showTaskMetaInput.addEventListener("change", () => {
  state.showTaskMeta = showTaskMetaInput.checked;
  syncLayoutVars();
  render();
});

showRelativeTimelineInput.addEventListener("change", () => {
  state.showRelativeTimeline = showRelativeTimelineInput.checked;
  render();
});

chartColorSchemeInput.addEventListener("change", () => {
  state.chartColorScheme = chartColorSchemeInput.value;
  render();
});

zoomOutButton.addEventListener("click", () => {
  updateZoom(-10);
});

zoomInButton.addEventListener("click", () => {
  updateZoom(10);
});

timelineModeButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setTimelineModeOverride(button.dataset.timelineMode);
  });
});

zoomLabel.addEventListener("change", () => {
  applyZoomInput();
});

zoomLabel.addEventListener("blur", () => {
  applyZoomInput();
});

zoomLabel.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    applyZoomInput();
    zoomLabel.blur();
  }

  if (event.key === "Escape") {
    syncZoomUi();
    zoomLabel.blur();
  }
});

exportCsvButton.addEventListener("click", () => {
  exportTasksToCsv();
});

exportExcelButton.addEventListener("click", () => {
  exportTasksToExcel();
});

exportPptButton.addEventListener("click", async () => {
  await exportChartToPpt();
});

copyChartImageButton.addEventListener("click", async () => {
  await copyChartImageToClipboard();
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
  timelineLabelModeInput.value = state.timelineModeOverride || "auto";
  showCompactDatesInput.checked = state.showCompactDates;
  showTaskDatesInput.checked = state.showTaskDates;
  showTaskMetaInput.checked = state.showTaskMeta;
  showRelativeTimelineInput.checked = state.showRelativeTimeline;
  chartColorSchemeInput.value = state.chartColorScheme;
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
  grid.dataset.chartColorScheme = state.chartColorScheme;
  if (!state.showTaskMeta) {
    grid.classList.add("chart-only-view");
  }
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

  const taskRows = buildTaskRows();

  taskRows.forEach((taskRow, index) => {
    grid.appendChild(renderTaskRow(taskRow, scale, range.start, timelineMode, index));
  });

  board.innerHTML = "";
  board.appendChild(grid);
  syncPlanHeading();
  syncPlanSummary();
}

function renderHeader(scale, timelineMode) {
  const header = document.createElement("div");
  const relativeTimelineStart = getTimelineRange().start;
  header.className = "timeline-header";
  if (timelineMode === "month") {
    header.classList.add("month-mode");
  }

  const cornerCell = document.createElement("div");
  cornerCell.className = `corner-cell${timelineMode === "month" ? " month-corner" : ""}`;
  if (state.showTaskMeta) {
    cornerCell.innerHTML = `
      <div class="eyebrow">Tasks</div>
      <div class="timeline-subtitle">${scale.length} ${timelineMode} slots</div>
    `;
  }
  header.appendChild(cornerCell);

  if (shouldShowCompactDates()) {
    header.appendChild(renderCompactDateHeaderCell("Start", "start"));
    header.appendChild(renderCompactDateHeaderCell("End", "end"));
  }

  if (timelineMode === "month") {
    renderMonthBands(header, scale, relativeTimelineStart);
    return header;
  }

  scale.forEach((slot) => {
    const cell = document.createElement("div");
    cell.className = "timeline-day";

    if (timelineMode === "day" && isWeekend(slot.start)) {
      cell.classList.add("weekend");
    }

    cell.innerHTML = getTimelineCellMarkup(slot, timelineMode, relativeTimelineStart);
    header.appendChild(cell);
  });

  return header;
}

function renderMonthBands(header, scale, timelineStart) {
  const yearGroups = state.showRelativeTimeline && timelineStart
    ? groupMonthsByRelativeYear(scale, timelineStart)
    : groupMonthsByYear(scale);
  const timelineStartColumn = getTimelineGridStartColumn();

  yearGroups.forEach((group, index) => {
    const yearCell = document.createElement("div");
    yearCell.className = `timeline-year-band ${getYearShadeClass(index)}`;
    yearCell.style.gridColumn = `${timelineStartColumn + group.startIndex} / span ${group.months.length}`;
    yearCell.style.gridRow = "1";
    if (state.showRelativeTimeline && timelineStart) {
      yearCell.textContent = `Year ${group.relativeYear}`;
    } else {
      yearCell.textContent = String(group.year);
    }
    header.appendChild(yearCell);
  });

  scale.forEach((slot, index) => {
    const shadeGroupIndex = state.showRelativeTimeline && timelineStart
      ? Math.floor(index / 12)
      : getYearGroupIndex(yearGroups, slot.start.getFullYear());
    const monthCell = document.createElement("div");
    monthCell.className = `timeline-month-band ${getYearShadeClass(shadeGroupIndex)}`;
    if (shouldRotateMonthLabel()) {
      monthCell.classList.add("rotated");
    }
    monthCell.style.gridColumn = `${timelineStartColumn + index}`;
    monthCell.style.gridRow = "2";
    if (state.showRelativeTimeline && timelineStart) {
      const monthNumber = index + 1;
      monthCell.innerHTML = `
        <div class="timeline-month">${getTimelinePeriodLabel("month")}</div>
        <div class="timeline-date">${monthNumber}</div>
      `;
    } else {
      monthCell.innerHTML = `
        <div class="timeline-month">${slot.start.toLocaleString("en-GB", { month: "short" })}</div>
      `;
    }
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

function renderCompactDateCell(taskRow, field) {
  const cell = document.createElement(taskRow.tasks.length > 1 ? "div" : "button");
  if (cell instanceof HTMLButtonElement) {
    cell.type = "button";
  }
  cell.className = "compact-date-cell";
  cell.textContent = formatCompactDate(taskRow[field], field);

  if (taskRow.tasks.length > 1) {
    cell.title = `Grouped ${field} date`;
    cell.classList.add("is-readonly");
    return cell;
  }

  cell.title = `Edit ${field} date`;
  cell.addEventListener("click", () => {
    state.editingTaskId = taskRow.id;
    state.editingTaskDateField = field;
    render();
  });
  return cell;
}

function renderTaskRow(taskRow, scale, timelineStart, timelineMode, index) {
  const row = document.createElement("div");
  row.className = "task-row";
  row.dataset.taskId = taskRow.id;
  row.dataset.taskIndex = String(index);
  row.style.minHeight = `${getTaskRowHeight(taskRow.tasks.length)}px`;

  if (state.rowDrag?.taskId === taskRow.id) {
    row.classList.add("is-row-dragging");
  }

  const label = document.createElement("div");
  label.className = "task-label";

  if (state.showTaskMeta) {
    const taskInfo = document.createElement("div");
    taskInfo.className = "task-info";
    taskInfo.appendChild(renderTaskTitleRow(taskRow));

    if (state.editingTaskId === taskRow.id && taskRow.tasks.length === 1) {
      taskInfo.appendChild(renderTaskDateEditor(taskRow.tasks[0]));
    } else if (state.showTaskDates && state.detailLevel !== "compressed") {
      const dateButton = document.createElement("button");
      dateButton.type = "button";
      dateButton.className = "task-date-button";
      const durationLabel = `${getDuration(taskRow.start, taskRow.end)} days`;
      dateButton.textContent = `${formatRange(taskRow.start, taskRow.end)} • ${durationLabel}`;
      if (taskRow.tasks.length === 1) {
        dateButton.addEventListener("click", () => {
          state.editingTaskId = taskRow.id;
          render();
        });
      } else {
        dateButton.disabled = true;
        dateButton.title = "Drag individual bars to adjust grouped task dates";
      }
      taskInfo.appendChild(dateButton);
    }

    const colorPickerWrap = document.createElement("label");
    colorPickerWrap.className = "task-color-picker";
    colorPickerWrap.title = "Change task color";

    const taskColorInput = document.createElement("input");
    taskColorInput.type = "color";
    taskColorInput.value = taskRow.color;
    taskColorInput.setAttribute("aria-label", `Change color for ${taskRow.name}`);
    taskColorInput.addEventListener("input", (event) => {
      taskRow.tasks.forEach((task) => {
        task.color = event.target.value;
      });
      render();
    });

    const ownerBadge = renderTaskOwnerControl(taskRow);
    const reorderHandle = document.createElement("button");
    reorderHandle.type = "button";
    reorderHandle.className = "task-reorder-handle";
    reorderHandle.setAttribute("aria-label", `Move ${taskRow.name} up or down`);
    reorderHandle.title = "Drag to reorder task";
    reorderHandle.textContent = "⋮⋮";
    reorderHandle.addEventListener("pointerdown", (event) => handleTaskReorderStart(event, taskRow.id));

    colorPickerWrap.appendChild(taskColorInput);
    label.append(reorderHandle, taskInfo, ownerBadge, colorPickerWrap);
  }
  row.appendChild(label);

  if (shouldShowCompactDates()) {
    row.appendChild(renderCompactDateCell(taskRow, "start"));
    row.appendChild(renderCompactDateCell(taskRow, "end"));
  }

  scale.forEach((slot) => {
    const cell = document.createElement("div");
    cell.className = "day-cell";
    if (timelineMode === "day" && isWeekend(slot.start)) {
      cell.classList.add("weekend");
    }
    if (timelineMode === "month" || timelineMode === "year") {
      cell.classList.add(getYearShadeClass(getYearBandIndex(scale, slot, timelineMode)));
    }
    row.appendChild(cell);
  });

  const layer = document.createElement("div");
  layer.className = "task-bar-layer";

  taskRow.tasks.forEach((task, taskIndex) => {
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
    if (taskRow.tasks.length === 1) {
      bar.style.top = "50%";
      bar.style.transform = task.milestone ? "translateY(-50%) rotate(45deg)" : "translateY(-50%)";
    } else {
      bar.style.top = `${taskIndex * getTaskBarStep()}px`;
      bar.style.transform = task.milestone ? "rotate(45deg)" : "none";
    }

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

    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.className = "task-bar-delete";
    deleteButton.textContent = "x";
    deleteButton.setAttribute("aria-label", `Delete ${task.name}`);
    deleteButton.title = "Delete task bar";
    deleteButton.addEventListener("pointerdown", (event) => {
      event.stopPropagation();
    });
    deleteButton.addEventListener("click", (event) => {
      event.stopPropagation();
      removeTask(task.id);
    });

    if (!task.milestone) {
      bar.append(startHandle, labelText, endHandle, deleteButton);
    } else {
      bar.append(labelText, deleteButton);
    }
    bar.addEventListener("pointerdown", (event) => handlePointerDown(event, task.id));
    layer.appendChild(bar);
  });
  row.appendChild(layer);

  return row;
}

function handlePointerDown(event, taskId) {
  const target = event.target;
  const dayWidth = getDayWidth();

  const dragType = target.dataset.dragType || "move";
  const task = state.tasks.find((entry) => entry.id === taskId);
  const timelineMode = getTimelineMode(getTimelineRange());
  const scale = buildTimelineScale(getTimelineRange().start, getTimelineRange().end, timelineMode);

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
    timelineMode,
    scale,
    originalStartUnits: getDateUnitsWithinScale(toDate(task.start), scale),
    originalEndUnits: getDateUnitsWithinScale(addDays(toDate(task.end), 1), scale, true),
  };

  window.addEventListener("pointermove", handlePointerMove);
  window.addEventListener("pointerup", handlePointerUp);
  window.addEventListener("pointercancel", handlePointerUp);
}

function handlePointerMove(event) {
  if (!state.drag) {
    return;
  }

  const unitDelta = (event.clientX - state.drag.originX) / state.drag.dayWidth;
  const task = state.tasks.find((entry) => entry.id === state.drag.taskId);

  if (!task) {
    return;
  }

  const originalStart = toDate(state.drag.originalStart);
  const originalEnd = toDate(state.drag.originalEnd);

  if (state.drag.dragType === "move") {
    const nextStart = getDateFromScaleUnits(state.drag.originalStartUnits + unitDelta, state.drag.scale);
    const nextEndExclusive = task.milestone
      ? nextStart
      : getDateFromScaleUnits(state.drag.originalEndUnits + unitDelta, state.drag.scale, true);
    const nextEnd = task.milestone ? nextStart : addDays(nextEndExclusive, -1);
    if (isWithinPlan(nextStart, nextEnd)) {
      task.start = formatInputDate(nextStart);
      task.end = formatInputDate(nextEnd);
    }
  }

  if (state.drag.dragType === "resize-start") {
    const nextStart = getDateFromScaleUnits(state.drag.originalStartUnits + unitDelta, state.drag.scale);
    if (nextStart <= toDate(task.end) && isWithinPlan(nextStart, toDate(task.end))) {
      task.start = formatInputDate(nextStart);
    }
  }

  if (state.drag.dragType === "resize-end") {
    const nextEndExclusive = getDateFromScaleUnits(state.drag.originalEndUnits + unitDelta, state.drag.scale, true);
    const nextEnd = addDays(nextEndExclusive, -1);
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

  const taskRows = buildTaskRows();
  const taskIndex = taskRows.findIndex((taskRow) => taskRow.id === taskId);
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
  const taskRows = buildTaskRows();
  const currentIndex = taskRows.findIndex((taskRow) => taskRow.id === taskId);
  if (currentIndex === -1 || currentIndex === nextIndex) {
    return;
  }

  const movedRow = taskRows[currentIndex];
  const nextRows = taskRows.filter((taskRow) => taskRow.id !== taskId);
  nextRows.splice(nextIndex, 0, movedRow);
  state.tasks = nextRows.flatMap((taskRow) => taskRow.tasks);
}

function buildTaskRows() {
  const rows = [];

  state.tasks.forEach((task) => {
    const existingRow = rows.find((row) => row.name === task.name);
    if (!existingRow) {
      rows.push(createTaskRow(task));
      return;
    }

    existingRow.tasks.push(task);
    existingRow.start = formatInputDate(new Date(Math.min(toDate(existingRow.start), toDate(task.start))));
    existingRow.end = formatInputDate(new Date(Math.max(toDate(existingRow.end), toDate(task.end))));
    existingRow.complete = existingRow.tasks.every((entry) => entry.complete);
    existingRow.owner = getSharedTaskOwner(existingRow.tasks);
  });

  return rows;
}

function createTaskRow(task) {
  return {
    id: task.name,
    name: task.name,
    start: task.start,
    end: task.end,
    color: task.color,
    owner: task.owner || "--",
    complete: task.complete,
    tasks: [task],
  };
}

function getSharedTaskOwner(tasks) {
  const owners = [...new Set(tasks.map((task) => task.owner || "--"))];
  return owners.length === 1 ? owners[0] : `${tasks.length}x`;
}

function getTaskBarHeight() {
  if (state.detailLevel === "compressed") {
    return 14;
  }
  if (state.detailLevel === "summary") {
    return 20;
  }
  return 34;
}

function getTaskBarGap() {
  if (state.detailLevel === "compressed") {
    return 4;
  }
  if (state.detailLevel === "summary") {
    return 5;
  }
  return 6;
}

function getTaskBarStep() {
  return getTaskBarHeight() + getTaskBarGap();
}

function getTaskRowHeight(taskCount) {
  const baseHeight = state.detailLevel === "compressed"
    ? 28
    : state.detailLevel === "summary"
      ? 38
      : 66;
  const topBottomInset = state.detailLevel === "compressed"
    ? 8
    : state.detailLevel === "summary"
      ? 12
      : 20;
  const stackedHeight = topBottomInset + (taskCount * getTaskBarHeight()) + ((taskCount - 1) * getTaskBarGap());
  return Math.max(baseHeight, stackedHeight);
}

function getLabelColumnWidth() {
  return state.showTaskMeta ? state.labelColumnWidth : 28;
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

function applyZoomInput() {
  const parsedZoom = Number.parseInt(zoomLabel.value, 10);

  if (Number.isNaN(parsedZoom)) {
    syncZoomUi();
    return;
  }

  state.zoom = clamp(parsedZoom, getMinimumZoom(), 360);
  syncZoomUi();
  render();
}

function syncLayoutVars() {
  document.documentElement.style.setProperty("--label-column", `${getLabelColumnWidth()}px`);
  document.documentElement.style.setProperty("--compact-start-column", `${state.compactStartColumnWidth}px`);
  document.documentElement.style.setProperty("--compact-end-column", `${state.compactEndColumnWidth}px`);
}

function syncCompactDateControl() {
  showCompactDatesInput.checked = state.showCompactDates;
}

function syncTimelineModeButtons() {
  timelineModeButtons.forEach((button) => {
    const isActive = button.dataset.timelineMode === state.timelineModeOverride;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
  });
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
  const minZoom = getMinimumZoom();
  zoomLabel.min = String(minZoom);
  zoomLabel.max = "360";
  zoomLabel.value = String(state.zoom);
  zoomOutButton.disabled = state.zoom <= minZoom;
  zoomInButton.disabled = state.zoom >= 360;
  syncTimelineModeButtons();
}

function syncPlanSummary() {
  planRangeLabel.textContent = `Plan: ${formatRange(state.planStart, state.planEnd)}`;
}

function syncPlanHeading() {
  if (!state.editingPlanTitle) {
    planHeading.textContent = state.planTitle || "Schedule Overview";
    planHeading.classList.remove("plan-heading-button");
    planHeading.removeAttribute("role");
    planHeading.removeAttribute("tabindex");
    planHeading.removeAttribute("aria-label");
    return;
  }

  const editor = document.createElement("input");
  editor.type = "text";
  editor.className = "plan-heading-editor";
  editor.value = state.planTitle;
  editor.maxLength = 80;
  editor.placeholder = "Schedule Overview";

  const saveTitle = () => {
    state.planTitle = editor.value.trim();
    planTitleInput.value = state.planTitle;
    state.editingPlanTitle = false;
    syncPlanHeading();
  };

  editor.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      saveTitle();
    }

    if (event.key === "Escape") {
      state.editingPlanTitle = false;
      syncPlanHeading();
    }
  });

  editor.addEventListener("blur", saveTitle);
  planHeading.textContent = "";
  planHeading.appendChild(editor);
  planHeading.classList.add("plan-heading-button");
  planHeading.setAttribute("role", "button");
  planHeading.setAttribute("tabindex", "0");
  planHeading.setAttribute("aria-label", "Edit plan title");
  queueMicrotask(() => {
    editor.focus();
    editor.select();
  });
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
  if (state.timelineModeOverride) {
    return state.timelineModeOverride;
  }
  return getAutomaticTimelineMode(range);
}

function getAutomaticTimelineMode(range = getTimelineRange()) {
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

function setTimelineModeOverride(mode) {
  state.timelineModeOverride = state.timelineModeOverride === mode ? null : mode;
  syncZoomUi();
  render();
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

async function exportChartToPpt() {
  try {
    const svgBlob = createChartExportSvgBlob();
    const imageBlob = await createBoardImageBlob(svgBlob);
    const imageSize = await getImageDimensionsFromBlob(imageBlob);
    const pptBlob = await buildSingleSlidePptx(imageBlob, imageSize.width, imageSize.height);
    downloadChartBlob(pptBlob, "gantt-chart.pptx");
  } catch (error) {
    window.alert("Could not export the chart to PowerPoint.");
  }
}

async function copyChartImageToClipboard() {
  try {
    const svgBlob = createBoardImageSvgBlob(board);
    const imageBlob = await createBoardImageBlob(svgBlob);

    if (!window.ClipboardItem || !navigator.clipboard?.write) {
      downloadChartBlob(imageBlob, "gantt-chart.png");
      window.alert("Image copy is unavailable in this browser, so the chart image was downloaded instead.");
      return;
    }

    await navigator.clipboard.write([
      new ClipboardItem({
        "image/png": imageBlob,
      }),
    ]);
  } catch (error) {
    try {
      const svgBlob = createBoardImageSvgBlob(board);
      downloadChartBlob(svgBlob, "gantt-chart.svg");
      window.alert("Could not copy the chart image to the clipboard, so the chart image was downloaded as an SVG instead.");
    } catch (fallbackError) {
      window.alert("Could not copy or download the chart image.");
    }
  }
}

function createBoardImageSvgBlob(element) {
  const width = Math.ceil(element.scrollWidth);
  const height = Math.ceil(element.scrollHeight);
  const clone = element.cloneNode(true);

  inlineComputedStyles(element, clone);
  clone.style.width = `${width}px`;
  clone.style.height = `${height}px`;
  clone.style.overflow = "visible";
  clone.style.maxWidth = "none";

  const wrapper = document.createElement("div");
  wrapper.setAttribute("xmlns", "http://www.w3.org/1999/xhtml");
  wrapper.style.width = `${width}px`;
  wrapper.style.height = `${height}px`;
  wrapper.appendChild(clone);

  const foreignObject = document.createElementNS("http://www.w3.org/2000/svg", "foreignObject");
  foreignObject.setAttribute("width", "100%");
  foreignObject.setAttribute("height", "100%");
  foreignObject.appendChild(wrapper);

  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svg.setAttribute("width", String(width));
  svg.setAttribute("height", String(height));
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  svg.appendChild(foreignObject);

  const svgMarkup = new XMLSerializer().serializeToString(svg);
  return new Blob([svgMarkup], { type: "image/svg+xml;charset=utf-8" });
}

function createChartExportSvgBlob() {
  const range = getTimelineRange();
  const timelineMode = getTimelineMode(range);
  const scale = buildTimelineScale(range.start, range.end, timelineMode);
  const taskRows = buildTaskRows();
  const dayWidth = getDayWidth();
  const labelWidth = getLabelColumnWidth();
  const showCompactDates = shouldShowCompactDates();
  const compactStartWidth = showCompactDates ? state.compactStartColumnWidth : 0;
  const compactEndWidth = showCompactDates ? state.compactEndColumnWidth : 0;
  const timelineX = labelWidth + compactStartWidth + compactEndWidth;
  const timelineWidth = scale.length * dayWidth;
  const headerHeight = timelineMode === "month" ? 92 : 56;
  const totalHeight = headerHeight + taskRows.reduce((sum, taskRow) => sum + getTaskRowHeight(taskRow.tasks.length), 0);
  const totalWidth = timelineX + timelineWidth;
  const scheme = getChartExportScheme();
  const timelineStart = range.start;
  const yearGroups = timelineMode === "month"
    ? (state.showRelativeTimeline ? groupMonthsByRelativeYear(scale, timelineStart) : groupMonthsByYear(scale))
    : [];

  const svgParts = [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
    `<svg xmlns="http://www.w3.org/2000/svg" width="${Math.ceil(totalWidth)}" height="${Math.ceil(totalHeight)}" viewBox="0 0 ${Math.ceil(totalWidth)} ${Math.ceil(totalHeight)}">`,
    `<rect width="${Math.ceil(totalWidth)}" height="${Math.ceil(totalHeight)}" rx="24" fill="${scheme.surface}"/>`,
    `<rect x="0" y="0" width="${labelWidth}" height="${headerHeight}" fill="${scheme.labelSurface}"/>`,
  ];

  if (showCompactDates) {
    svgParts.push(
      `<rect x="${labelWidth}" y="0" width="${compactStartWidth}" height="${headerHeight}" fill="${scheme.headerSurface}"/>`,
      `<rect x="${labelWidth + compactStartWidth}" y="0" width="${compactEndWidth}" height="${headerHeight}" fill="${scheme.headerSurface}"/>`
    );
  }

  svgParts.push(`<rect x="${timelineX}" y="0" width="${timelineWidth}" height="${headerHeight}" fill="${scheme.headerSurface}"/>`);

  if (state.showTaskMeta) {
    svgParts.push(
      `<text x="20" y="22" font-family="Arial, sans-serif" font-size="11" font-weight="700" letter-spacing="1.4" fill="${scheme.muted}">TASKS</text>`,
      `<text x="20" y="46" font-family="Arial, sans-serif" font-size="16" font-weight="700" fill="${scheme.text}">${escapeHtml(`${scale.length} ${timelineMode} slots`)}</text>`
    );
  }

  if (showCompactDates) {
    svgParts.push(
      `<text x="${labelWidth + (compactStartWidth / 2)}" y="${(headerHeight / 2) + 6}" text-anchor="middle" font-family="Arial, sans-serif" font-size="12" font-weight="700" fill="${scheme.muted}">Start</text>`,
      `<text x="${labelWidth + compactStartWidth + (compactEndWidth / 2)}" y="${(headerHeight / 2) + 6}" text-anchor="middle" font-family="Arial, sans-serif" font-size="12" font-weight="700" fill="${scheme.muted}">End</text>`
    );
  }

  if (timelineMode === "month") {
    yearGroups.forEach((group, index) => {
      const x = timelineX + (group.startIndex * dayWidth);
      const width = group.months.length * dayWidth;
      const fill = index % 2 === 0 ? scheme.slotEven : scheme.slotOdd;
      const yearLabel = state.showRelativeTimeline ? `Year ${group.relativeYear}` : String(group.year);
      svgParts.push(
        `<rect x="${x}" y="0" width="${width}" height="44" fill="${fill}"/>`,
        `<text x="${x + (width / 2)}" y="27" text-anchor="middle" font-family="Arial, sans-serif" font-size="15" font-weight="700" fill="${scheme.text}">${escapeHtml(yearLabel)}</text>`
      );
    });

    scale.forEach((slot, index) => {
      const shadeIndex = state.showRelativeTimeline ? Math.floor(index / 12) : getYearGroupIndex(yearGroups, slot.start.getFullYear());
      const x = timelineX + (index * dayWidth);
      const fill = shadeIndex % 2 === 0 ? scheme.slotEven : scheme.slotOdd;
      const topLabel = state.showRelativeTimeline
        ? getTimelinePeriodLabel("month")
        : slot.start.toLocaleString("en-GB", { month: "short" });
      const bottomLabel = state.showRelativeTimeline ? String(index + 1) : "";
      svgParts.push(
        `<rect x="${x}" y="44" width="${dayWidth}" height="48" fill="${fill}" stroke="${scheme.lineSoft}" stroke-width="1"/>`,
        `<text x="${x + (dayWidth / 2)}" y="${state.showRelativeTimeline ? 63 : 72}" text-anchor="middle" font-family="Arial, sans-serif" font-size="12" fill="${scheme.muted}">${escapeHtml(topLabel)}</text>`,
        bottomLabel
          ? `<text x="${x + (dayWidth / 2)}" y="80" text-anchor="middle" font-family="Arial, sans-serif" font-size="13" font-weight="700" fill="${scheme.text}">${escapeHtml(bottomLabel)}</text>`
          : ""
      );
    });
  } else {
    scale.forEach((slot, index) => {
      const x = timelineX + (index * dayWidth);
      const fill = getExportSlotFill(slot, scale, timelineMode, scheme);
      const labels = getExportTimelineLabels(slot, timelineMode, timelineStart);
      svgParts.push(
        `<rect x="${x}" y="0" width="${dayWidth}" height="${headerHeight}" fill="${fill}" stroke="${scheme.lineSoft}" stroke-width="1"/>`,
        `<text x="${x + (dayWidth / 2)}" y="24" text-anchor="middle" font-family="Arial, sans-serif" font-size="11" fill="${scheme.muted}">${escapeHtml(labels.top)}</text>`,
        `<text x="${x + (dayWidth / 2)}" y="42" text-anchor="middle" font-family="Arial, sans-serif" font-size="13" font-weight="700" fill="${scheme.text}">${escapeHtml(labels.bottom)}</text>`
      );
    });
  }

  let rowY = headerHeight;
  taskRows.forEach((taskRow) => {
    const rowHeight = getTaskRowHeight(taskRow.tasks.length);
    const barLayerTop = getExportBarLayerTop();
    const taskBarHeight = getTaskBarHeight();

    if (state.showTaskMeta) {
      svgParts.push(`<rect x="0" y="${rowY}" width="${labelWidth}" height="${rowHeight}" fill="${scheme.labelSurface}" stroke="${scheme.lineSoft}" stroke-width="1"/>`);
      svgParts.push(
        `<text x="20" y="${rowY + 24}" font-family="Arial, sans-serif" font-size="14" font-weight="700" fill="${scheme.text}">${escapeHtml(taskRow.name)}</text>`
      );
      if (state.showTaskDates && state.detailLevel !== "compressed") {
        const durationLabel = `${getDuration(taskRow.start, taskRow.end)} days`;
        svgParts.push(
          `<text x="20" y="${rowY + 44}" font-family="Arial, sans-serif" font-size="11" fill="${scheme.muted}">${escapeHtml(`${formatRange(taskRow.start, taskRow.end)} • ${durationLabel}`)}</text>`
        );
      }
    }

    if (showCompactDates) {
      svgParts.push(
        `<rect x="${labelWidth}" y="${rowY}" width="${compactStartWidth}" height="${rowHeight}" fill="${scheme.compactSurface}" stroke="${scheme.lineSoft}" stroke-width="1"/>`,
        `<rect x="${labelWidth + compactStartWidth}" y="${rowY}" width="${compactEndWidth}" height="${rowHeight}" fill="${scheme.compactSurface}" stroke="${scheme.lineSoft}" stroke-width="1"/>`,
        `<text x="${labelWidth + (compactStartWidth / 2)}" y="${rowY + (rowHeight / 2) + 4}" text-anchor="middle" font-family="Arial, sans-serif" font-size="11" fill="${scheme.text}">${escapeHtml(formatCompactDate(taskRow.start, "start"))}</text>`,
        `<text x="${labelWidth + compactStartWidth + (compactEndWidth / 2)}" y="${rowY + (rowHeight / 2) + 4}" text-anchor="middle" font-family="Arial, sans-serif" font-size="11" fill="${scheme.text}">${escapeHtml(formatCompactDate(taskRow.end, "end"))}</text>`
      );
    }

    scale.forEach((slot, index) => {
      const x = timelineX + (index * dayWidth);
      const fill = getExportSlotFill(slot, scale, timelineMode, scheme);
      svgParts.push(
        `<rect x="${x}" y="${rowY}" width="${dayWidth}" height="${rowHeight}" fill="${fill}" stroke="${scheme.lineSubtle}" stroke-width="1"/>`
      );
    });

    taskRow.tasks.forEach((task, taskIndex) => {
      const taskPosition = getTaskPosition(task, scale);
      const x = task.milestone
        ? timelineX + (taskPosition.centerUnits * dayWidth) - 12
        : timelineX + (taskPosition.startUnits * dayWidth);
      const width = task.milestone ? 24 : Math.max(taskPosition.widthUnits * dayWidth, 6);
      const y = taskRow.tasks.length === 1
        ? rowY + ((rowHeight - taskBarHeight) / 2)
        : rowY + barLayerTop + (taskIndex * getTaskBarStep());
      const color = getTaskDisplayColor(task);
      const labelHidden = shouldHideTaskLabel(task, scale, timelineMode);
      const labelColor = getContrastTextColor(color);

      if (task.milestone) {
        const cx = x + 12;
        const cy = y + 12;
        svgParts.push(
          `<polygon points="${cx},${y} ${x + 24},${cy} ${cx},${y + 24} ${x},${cy}" fill="${color}"/>`
        );
      } else {
        svgParts.push(
          `<rect x="${x}" y="${y}" width="${width}" height="${taskBarHeight}" rx="${taskBarHeight / 2}" fill="${color}"/>`,
          `<circle cx="${x + 10}" cy="${y + (taskBarHeight / 2)}" r="5" fill="rgba(255,255,255,0.92)"/>`,
          `<circle cx="${x + width - 10}" cy="${y + (taskBarHeight / 2)}" r="5" fill="rgba(255,255,255,0.92)"/>`
        );
      }

      if (!labelHidden) {
        svgParts.push(
          `<text x="${x + (width / 2)}" y="${y + (taskBarHeight / 2) + 5}" text-anchor="middle" font-family="Arial, sans-serif" font-size="12" font-weight="700" fill="${labelColor}">${escapeHtml(task.name)}</text>`
        );
      }
    });

    rowY += rowHeight;
  });

  svgParts.push(
    `<rect x="0.5" y="0.5" width="${Math.ceil(totalWidth) - 1}" height="${Math.ceil(totalHeight) - 1}" rx="24" fill="none" stroke="${scheme.border}" stroke-width="1"/>`,
    `</svg>`
  );

  return new Blob([svgParts.join("")], { type: "image/svg+xml;charset=utf-8" });
}

function getChartExportScheme() {
  if (state.chartColorScheme === "cool-grey") {
    return {
      surface: "#eceff1",
      headerSurface: "#f1f4f6",
      labelSurface: "#f1f4f6",
      compactSurface: "#e9edf0",
      slotEven: "#f5f7f8",
      slotOdd: "#e6eaed",
      weekend: "#dde2e6",
      border: "#d4cbc0",
      lineSoft: "rgba(35, 27, 22, 0.08)",
      lineSubtle: "rgba(35, 27, 22, 0.05)",
      text: "#231b16",
      muted: "#74665b",
    };
  }

  if (state.chartColorScheme === "striped-grey") {
    return {
      surface: "#e7eaed",
      headerSurface: "#eef1f3",
      labelSurface: "#eef1f3",
      compactSurface: "#e4e8ec",
      slotEven: "#f2f4f6",
      slotOdd: "#d8dde2",
      weekend: "#cdd3d9",
      border: "#d4cbc0",
      lineSoft: "rgba(35, 27, 22, 0.08)",
      lineSubtle: "rgba(35, 27, 22, 0.05)",
      text: "#231b16",
      muted: "#74665b",
    };
  }

  return {
    surface: "#fbf7f0",
    headerSurface: "#fffaf3",
    labelSurface: "#fffaf3",
    compactSurface: "#fff4e1",
    slotEven: "#fffdf9",
    slotOdd: "#fff8ec",
    weekend: "#fff1dd",
    border: "#d4cbc0",
    lineSoft: "rgba(35, 27, 22, 0.08)",
    lineSubtle: "rgba(35, 27, 22, 0.05)",
    text: "#231b16",
    muted: "#74665b",
  };
}

function getExportSlotFill(slot, scale, timelineMode, scheme) {
  if (timelineMode === "day" && isWeekend(slot.start)) {
    return scheme.weekend;
  }

  if (timelineMode === "month" || timelineMode === "year") {
    const bandIndex = getYearBandIndex(scale, slot, timelineMode);
    return bandIndex % 2 === 0 ? scheme.slotEven : scheme.slotOdd;
  }

  return scheme.surface;
}

function getExportTimelineLabels(slot, timelineMode, timelineStart) {
  if (state.showRelativeTimeline) {
    if (timelineMode === "year") {
      const yearNumber = Math.floor(getRelativeMonthOffset(getVisibleSlotAnchor(slot), timelineStart) / 12) + 1;
      return { top: "Year", bottom: String(yearNumber) };
    }

    if (timelineMode === "week") {
      const weekNumber = Math.floor(diffDays(getStartOfWeek(timelineStart), slot.start) / 7) + 1;
      return { top: getTimelinePeriodLabel("week"), bottom: String(weekNumber) };
    }

    const dayNumber = diffDays(timelineStart, slot.start) + 1;
    return { top: getTimelinePeriodLabel("day"), bottom: String(dayNumber) };
  }

  if (timelineMode === "year") {
    return { top: "Year", bottom: String(slot.start.getFullYear()) };
  }

  if (timelineMode === "week") {
    return {
      top: getTimelinePeriodLabel("week"),
      bottom: `${slot.start.getDate()} ${slot.start.toLocaleString("en-GB", { month: "short" })}`,
    };
  }

  return {
    top: slot.start.toLocaleString("en-GB", { month: "short" }),
    bottom: String(slot.start.getDate()),
  };
}

function getExportBarLayerTop() {
  if (state.detailLevel === "compressed") {
    return 4;
  }
  if (state.detailLevel === "summary") {
    return 6;
  }
  return 10;
}

async function createBoardImageBlob(svgBlob) {
  const svgText = await svgBlob.text();
  const widthMatch = svgText.match(/viewBox="0 0 (\d+) (\d+)"/);
  const width = widthMatch ? Number.parseInt(widthMatch[1], 10) : 0;
  const height = widthMatch ? Number.parseInt(widthMatch[2], 10) : 0;
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");

  if (!context) {
    throw new Error("Canvas rendering is unavailable.");
  }

  const image = await loadSvgSnapshotImage(svgBlob, svgText);
  context.drawImage(image, 0, 0, width, height);

  return await new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("PNG generation failed."));
        return;
      }
      resolve(blob);
    }, "image/png");
  });
}

function inlineComputedStyles(sourceNode, targetNode) {
  if (!(sourceNode instanceof Element) || !(targetNode instanceof Element)) {
    return;
  }

  const computedStyle = getComputedStyle(sourceNode);
  for (const propertyName of computedStyle) {
    targetNode.style.setProperty(
      propertyName,
      computedStyle.getPropertyValue(propertyName),
      computedStyle.getPropertyPriority(propertyName)
    );
  }

  const sourceChildren = [...sourceNode.children];
  const targetChildren = [...targetNode.children];
  sourceChildren.forEach((sourceChild, index) => {
    inlineComputedStyles(sourceChild, targetChildren[index]);
  });
}

function loadImage(url) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("Image load failed."));
    image.src = url;
  });
}

async function loadSvgSnapshotImage(svgBlob, svgText) {
  const blobUrl = URL.createObjectURL(svgBlob);

  try {
    return await loadImage(blobUrl);
  } catch (blobError) {
    const dataUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svgText)}`;

    try {
      return await loadImage(dataUrl);
    } catch (dataUrlError) {
      throw new Error("SVG snapshot rendering failed.");
    }
  } finally {
    URL.revokeObjectURL(blobUrl);
  }
}

async function getImageDimensionsFromBlob(blob) {
  const url = URL.createObjectURL(blob);
  try {
    const image = await loadImage(url);
    return {
      width: image.naturalWidth || image.width,
      height: image.naturalHeight || image.height,
    };
  } finally {
    URL.revokeObjectURL(url);
  }
}

function downloadChartBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function buildSingleSlidePptx(imageBlob, imageWidth, imageHeight) {
  const slideWidth = 12192000;
  const slideHeight = 6858000;
  const margin = 228600;
  const contentWidth = slideWidth - (margin * 2);
  const contentHeight = slideHeight - (margin * 2);
  const imageRatio = imageWidth / Math.max(imageHeight, 1);
  const contentRatio = contentWidth / contentHeight;

  let targetWidth = contentWidth;
  let targetHeight = Math.round(contentWidth / imageRatio);
  if (imageRatio < contentRatio) {
    targetHeight = contentHeight;
    targetWidth = Math.round(contentHeight * imageRatio);
  }

  const offsetX = Math.round((slideWidth - targetWidth) / 2);
  const offsetY = Math.round((slideHeight - targetHeight) / 2);
  const imageBytes = new Uint8Array(await imageBlob.arrayBuffer());
  const imageExtension = getPptImageExtension(imageBlob.type);
  const imageContentType = getPptImageContentType(imageBlob.type);
  const nowIso = new Date().toISOString();

  const files = [
    {
      name: "[Content_Types].xml",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="${imageExtension}" ContentType="${imageContentType}"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`),
    },
    {
      name: "docProps/core.xml",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Gantt Chart</dc:title>
  <dc:creator>OpenAI Codex</dc:creator>
  <cp:lastModifiedBy>OpenAI Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${nowIso}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${nowIso}</dcterms:modified>
</cp:coreProperties>`),
    },
    {
      name: "docProps/app.xml",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Interactive Gantt Chart</Application>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Slides>1</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <MMClips>0</MMClips>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>1</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr">
      <vt:lpstr>Office Theme</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>`),
    },
    {
      name: "ppt/presentation.xml",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1" autoCompressPictures="0">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="${slideWidth}" cy="${slideHeight}"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle/>
</p:presentation>`),
    },
    {
      name: "ppt/_rels/presentation.xml.rels",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>
</Relationships>`),
    },
    {
      name: "ppt/presProps.xml",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`),
    },
    {
      name: "ppt/viewProps.xml",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:normalViewPr>
    <p:restoredLeft sz="15620"/>
    <p:restoredTop sz="94660"/>
  </p:normalViewPr>
  <p:slideViewPr>
    <p:cSldViewPr snapToGrid="1" snapToObjects="1"/>
  </p:slideViewPr>
  <p:notesTextViewPr>
    <p:cViewPr varScale="1">
      <p:scale sx="100" sy="100"/>
      <p:origin x="0" y="0"/>
    </p:cViewPr>
  </p:notesTextViewPr>
  <p:gridSpacing cx="72008" cy="72008"/>
</p:viewPr>`),
    },
    {
      name: "ppt/tableStyles.xml",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`),
    },
    {
      name: "ppt/slideMasters/slideMaster1.xml",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:schemeClr val="bg1"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle/>
    <p:bodyStyle/>
    <p:otherStyle/>
  </p:txStyles>
</p:sldMaster>`),
    },
    {
      name: "ppt/slideMasters/_rels/slideMaster1.xml.rels",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`),
    },
    {
      name: "ppt/slideLayouts/slideLayout1.xml",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`),
    },
    {
      name: "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`),
    },
    {
      name: "ppt/theme/theme1.xml",
      data: textEncoder.encode(getMinimalThemeXml()),
    },
    {
      name: "ppt/slides/slide1.xml",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:pic>
        <p:nvPicPr>
          <p:cNvPr id="2" name="Gantt Chart"/>
          <p:cNvPicPr/>
          <p:nvPr/>
        </p:nvPicPr>
        <p:blipFill>
          <a:blip r:embed="rId2"/>
          <a:stretch><a:fillRect/></a:stretch>
        </p:blipFill>
        <p:spPr>
          <a:xfrm>
            <a:off x="${offsetX}" y="${offsetY}"/>
            <a:ext cx="${targetWidth}" cy="${targetHeight}"/>
          </a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
      </p:pic>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`),
    },
    {
      name: "ppt/slides/_rels/slide1.xml.rels",
      data: textEncoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.${imageExtension}"/>
</Relationships>`),
    },
    {
      name: `ppt/media/image1.${imageExtension}`,
      data: imageBytes,
    },
  ];

  return createZipBlob(files);
}

function getPptImageExtension(mimeType) {
  if (mimeType === "image/svg+xml;charset=utf-8" || mimeType === "image/svg+xml") {
    return "svg";
  }

  if (mimeType === "image/jpeg") {
    return "jpeg";
  }

  return "png";
}

function getPptImageContentType(mimeType) {
  if (mimeType === "image/svg+xml;charset=utf-8" || mimeType === "image/svg+xml") {
    return "image/svg+xml";
  }

  if (mimeType === "image/jpeg") {
    return "image/jpeg";
  }

  return "image/png";
}

function getMinimalThemeXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2>
      <a:lt2><a:srgbClr val="EEECE1"/></a:lt2>
      <a:accent1><a:srgbClr val="4F81BD"/></a:accent1>
      <a:accent2><a:srgbClr val="C0504D"/></a:accent2>
      <a:accent3><a:srgbClr val="9BBB59"/></a:accent3>
      <a:accent4><a:srgbClr val="8064A2"/></a:accent4>
      <a:accent5><a:srgbClr val="4BACC6"/></a:accent5>
      <a:accent6><a:srgbClr val="F79646"/></a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="800080"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Arial"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Arial"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>`;
}

function createZipBlob(files) {
  const localRecords = [];
  const centralRecords = [];
  let offset = 0;

  files.forEach((file) => {
    const nameBytes = textEncoder.encode(file.name);
    const data = file.data instanceof Uint8Array ? file.data : new Uint8Array(file.data);
    const crc = crc32(data);
    const localHeader = createZipLocalHeader(nameBytes, data.length, crc);
    localRecords.push(localHeader, nameBytes, data);

    const centralHeader = createZipCentralHeader(nameBytes, data.length, crc, offset);
    centralRecords.push(centralHeader, nameBytes);

    offset += localHeader.length + nameBytes.length + data.length;
  });

  const centralDirectorySize = centralRecords.reduce((sum, part) => sum + part.length, 0);
  const endRecord = createZipEndRecord(files.length, centralDirectorySize, offset);

  return new Blob([...localRecords, ...centralRecords, endRecord], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
}

function createZipLocalHeader(nameBytes, size, crc) {
  const header = new Uint8Array(30);
  const view = new DataView(header.buffer);
  view.setUint32(0, 0x04034b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, 0, true);
  view.setUint16(10, 0, true);
  view.setUint16(12, 0, true);
  view.setUint32(14, crc, true);
  view.setUint32(18, size, true);
  view.setUint32(22, size, true);
  view.setUint16(26, nameBytes.length, true);
  view.setUint16(28, 0, true);
  return header;
}

function createZipCentralHeader(nameBytes, size, crc, offset) {
  const header = new Uint8Array(46);
  const view = new DataView(header.buffer);
  view.setUint32(0, 0x02014b50, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 20, true);
  view.setUint16(8, 0, true);
  view.setUint16(10, 0, true);
  view.setUint16(12, 0, true);
  view.setUint16(14, 0, true);
  view.setUint32(16, crc, true);
  view.setUint32(20, size, true);
  view.setUint32(24, size, true);
  view.setUint16(28, nameBytes.length, true);
  view.setUint16(30, 0, true);
  view.setUint16(32, 0, true);
  view.setUint16(34, 0, true);
  view.setUint16(36, 0, true);
  view.setUint32(38, 0, true);
  view.setUint32(42, offset, true);
  return header;
}

function createZipEndRecord(entryCount, centralDirectorySize, centralDirectoryOffset) {
  const record = new Uint8Array(22);
  const view = new DataView(record.buffer);
  view.setUint32(0, 0x06054b50, true);
  view.setUint16(4, 0, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, entryCount, true);
  view.setUint16(10, entryCount, true);
  view.setUint32(12, centralDirectorySize, true);
  view.setUint32(16, centralDirectoryOffset, true);
  view.setUint16(20, 0, true);
  return record;
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (let index = 0; index < bytes.length; index += 1) {
    crc ^= bytes[index];
    for (let bit = 0; bit < 8; bit += 1) {
      const mask = -(crc & 1);
      crc = (crc >>> 1) ^ (0xedb88320 & mask);
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

const textEncoder = new TextEncoder();

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

function renderTaskTitleRow(taskRow) {
  const titleRow = document.createElement("div");
  titleRow.className = "task-title-row";

  const completeButton = document.createElement("button");
  completeButton.type = "button";
  completeButton.className = `task-complete-button${taskRow.complete ? " is-complete" : ""}`;
  completeButton.textContent = taskRow.complete ? "✓" : "";
  completeButton.setAttribute("aria-label", taskRow.complete ? `Mark ${taskRow.name} incomplete` : `Mark ${taskRow.name} complete`);
  completeButton.addEventListener("click", () => {
    const nextComplete = !taskRow.complete;
    taskRow.tasks.forEach((task) => {
      task.complete = nextComplete;
    });
    render();
  });
  titleRow.appendChild(completeButton);

  if (state.editingTaskNameId === taskRow.id) {
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.className = "task-name-editor";
    nameInput.value = taskRow.name;
    nameInput.maxLength = 50;
    nameInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        saveTaskName(taskRow, nameInput.value);
      }
      if (event.key === "Escape") {
        state.editingTaskNameId = null;
        render();
      }
    });
    nameInput.addEventListener("blur", () => {
      saveTaskName(taskRow, nameInput.value);
    });
    titleRow.appendChild(nameInput);
    queueMicrotask(() => nameInput.focus());
  } else {
    const titleButton = document.createElement("button");
    titleButton.type = "button";
    titleButton.className = "task-name-button";
    if (taskRow.complete) {
      titleButton.classList.add("is-complete");
    }
    titleButton.textContent = taskRow.name;
    titleButton.title = state.detailLevel === "summary" ? formatShortRange(taskRow.start, taskRow.end) : taskRow.name;
    titleButton.addEventListener("click", () => {
      state.editingTaskNameId = taskRow.id;
      render();
    });
    titleRow.appendChild(titleButton);
  }

  const removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.className = "task-remove-button";
  removeButton.textContent = "x";
  removeButton.setAttribute("aria-label", `Remove ${taskRow.name}`);
  removeButton.addEventListener("click", () => {
    removeTask(taskRow.id);
  });
  titleRow.appendChild(removeButton);

  return titleRow;
}

function renderTaskOwnerControl(taskRow) {
  if (state.editingTaskOwnerId === taskRow.id) {
    const ownerEditor = document.createElement("input");
    ownerEditor.type = "text";
    ownerEditor.className = "task-owner-editor";
    ownerEditor.value = taskRow.owner === "--" || taskRow.owner.endsWith("x") ? "" : taskRow.owner;
    ownerEditor.maxLength = 4;
    ownerEditor.setAttribute("aria-label", `Edit owner for ${taskRow.name}`);
    ownerEditor.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        saveTaskOwner(taskRow, ownerEditor.value);
      }
      if (event.key === "Escape") {
        state.editingTaskOwnerId = null;
        render();
      }
    });
    ownerEditor.addEventListener("blur", () => {
      saveTaskOwner(taskRow, ownerEditor.value);
    });
    queueMicrotask(() => ownerEditor.focus());
    return ownerEditor;
  }

  const ownerBadge = document.createElement("button");
  ownerBadge.type = "button";
  ownerBadge.className = "task-owner-badge";
  ownerBadge.textContent = taskRow.owner || "--";
  ownerBadge.title = taskRow.owner && !taskRow.owner.endsWith("x") ? `Owner ${taskRow.owner}` : "Set owner";
  ownerBadge.addEventListener("click", () => {
    state.editingTaskOwnerId = taskRow.id;
    render();
  });
  return ownerBadge;
}

function saveTaskName(taskRow, nextName) {
  const trimmed = nextName.trim();
  if (trimmed) {
    taskRow.tasks.forEach((task) => {
      task.name = trimmed;
    });
  }
  state.editingTaskNameId = null;
  render();
}

function saveTaskOwner(taskRow, nextOwner) {
  const owner = normaliseOwner(nextOwner);
  taskRow.tasks.forEach((task) => {
    task.owner = owner;
  });
  state.editingTaskOwnerId = null;
  render();
}

function removeTask(taskId) {
  state.tasks = state.tasks.filter((task) => task.id !== taskId && task.name !== taskId);
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

function getDateFromScaleUnits(units, scale, endExclusive = false) {
  if (!scale.length) {
    return new Date();
  }

  if (units <= 0) {
    return new Date(scale[0].start);
  }

  if (units >= scale.length) {
    const lastSlotEndExclusive = addDays(scale[scale.length - 1].end, 1);
    return endExclusive ? lastSlotEndExclusive : addDays(lastSlotEndExclusive, -1);
  }

  const slotIndex = Math.min(Math.floor(units), scale.length - 1);
  const slot = scale[slotIndex];
  const slotEndExclusive = addDays(slot.end, 1);
  const slotDays = Math.max(diffDays(slot.start, slotEndExclusive), 1);
  const fractionalUnits = units - slotIndex;
  const dayOffset = Math.round(fractionalUnits * slotDays);
  const nextDate = addDays(slot.start, dayOffset);

  if (endExclusive) {
    return nextDate > slotEndExclusive ? slotEndExclusive : nextDate;
  }

  return nextDate >= slotEndExclusive ? addDays(slotEndExclusive, -1) : nextDate;
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
    const monthEnd = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0);
    months.push({ start: new Date(cursor), end: monthEnd });
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
      start: yearStart,
      end: yearEnd,
    });
    year += 1;
  }

  return years;
}

function getTimelineCellMarkup(slot, timelineMode, timelineStart = slot.start) {
  if (state.showRelativeTimeline) {
    return getRelativeTimelineCellMarkup(slot, timelineMode, timelineStart);
  }

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
      <div class="timeline-month">${getTimelinePeriodLabel("week")}</div>
      <div class="timeline-date">${slot.start.getDate()} ${slot.start.toLocaleString("en-GB", { month: "short" })}</div>
    `;
  }

  return `
    <div class="timeline-month">${slot.start.toLocaleString("en-GB", { month: "short" })}</div>
    <div class="timeline-date">${slot.start.getDate()}</div>
  `;
}

function getRelativeTimelineCellMarkup(slot, timelineMode, timelineStart) {
  if (timelineMode === "year") {
    const yearNumber = Math.floor(getRelativeMonthOffset(getVisibleSlotAnchor(slot), timelineStart) / 12) + 1;
    return `
      <div class="timeline-month">Year</div>
      <div class="timeline-date">${yearNumber}</div>
    `;
  }

  if (timelineMode === "month") {
    const monthNumber = getRelativeMonthOffset(getVisibleSlotEndAnchor(slot), timelineStart) + 1;
    return `
      <div class="timeline-month">${getTimelinePeriodLabel("month")}</div>
      <div class="timeline-date">${monthNumber}</div>
    `;
  }

  if (timelineMode === "week") {
    const weekNumber = Math.floor(diffDays(getStartOfWeek(timelineStart), slot.start) / 7) + 1;
    return `
      <div class="timeline-month">${getTimelinePeriodLabel("week")}</div>
      <div class="timeline-date">${weekNumber}</div>
    `;
  }

  const dayNumber = diffDays(timelineStart, slot.start) + 1;
  return `
    <div class="timeline-month">${getTimelinePeriodLabel("day")}</div>
    <div class="timeline-date">${dayNumber}</div>
  `;
}

function getTimelinePeriodLabel(period) {
  if (getDayWidth() < 48) {
    if (period === "day") {
      return "D";
    }
    if (period === "month") {
      return "M";
    }
    if (period === "week") {
      return "W";
    }
  }

  if (period === "day") {
    return "Day";
  }
  if (period === "month") {
    return "Month";
  }
  if (period === "week") {
    return "Week";
  }
  return period;
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

function groupMonthsByRelativeYear(scale, timelineStart) {
  const groups = [];

  scale.forEach((slot, index) => {
    const monthIndex = index;
    const relativeYear = Math.floor(monthIndex / 12) + 1;
    const lastGroup = groups[groups.length - 1];

    if (!lastGroup || lastGroup.relativeYear !== relativeYear) {
      groups.push({
        year: slot.start.getFullYear(),
        relativeYear,
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

function getYearBandIndex(scale, slot, timelineMode) {
  if (timelineMode === "year") {
    if (state.showRelativeTimeline) {
      const yearIndex = Math.floor(
        diffMonths(getTimelineRange().start, slot.start) / 12
      );
      return Math.max(yearIndex, 0);
    }
    return getYearGroupIndex(groupYears(scale), slot.start.getFullYear());
  }
  if (timelineMode === "month" && state.showRelativeTimeline) {
    const monthIndex = scale.findIndex((entry) => entry.start.getTime() === slot.start.getTime());
    return Math.floor(Math.max(monthIndex, 0) / 12);
  }
  return getYearGroupIndex(groupMonthsByYear(scale), slot.start.getFullYear());
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
