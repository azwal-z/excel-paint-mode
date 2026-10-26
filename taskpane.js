// Paint Mode: Ctrl+Shift+Z fills the current selection with the chosen colour.
// Ctrl+Shift+W joins rows that spilled over in the selected column back into the row above.

const NO_FILL = "none";

// Excel's default Office theme: background/text colours, then Accent 1-6.
const THEME_COLORS = [
  "#FFFFFF", "#000000", "#E7E6E6", "#44546A", "#4472C4",
  "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47",
];

const STANDARD_COLORS = [
  "#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050",
  "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0",
];

const VERSION = "v7";

let currentColor = loadColor();

// Registered at load time (as in Microsoft's sample), not inside onReady.
let associateError = null;
try {
  Office.actions.associate("APPLYFILL", paintSelection);
  Office.actions.associate("JOINROWS", joinSplitRows);
} catch (e) {
  associateError = e;
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  document.getElementById("version").textContent = "Paint Mode " + VERSION;
  showShortcutState();

  // Excel's shortcuts only fire when the grid has focus; catch them here too for when the panel does.
  document.addEventListener("keydown", (e) => {
    if (!e.ctrlKey || !e.shiftKey) return;
    const key = e.key.toLowerCase();
    if (key === "z") {
      e.preventDefault();
      paintSelection();
    } else if (key === "w") {
      e.preventDefault();
      joinSplitRows();
    }
  });

  buildPalettes();
  document.getElementById("noFill").addEventListener("click", () => setColor(NO_FILL));
  document.getElementById("custom").addEventListener("input", (e) => setColor(e.target.value.toUpperCase()));
  render();
});

// Shows the key Excel has actually assigned to this add-in, which can differ from shortcuts.json
// if Excel's own command won a conflict.
async function showShortcutState() {
  const fillHint = document.getElementById("hint");
  const joinHint = document.getElementById("joinHint");
  if (associateError) {
    fillHint.textContent = "Keyboard shortcuts aren't supported in this version of Excel (" + associateError.message + ").";
    joinHint.textContent = "";
    return;
  }
  try {
    const keys = await Office.actions.getShortcuts();
    describeShortcut(fillHint, keys.APPLYFILL, "Ctrl+Shift+Z", "Pick a colour, then press KEY on any cell");
    describeShortcut(joinHint, keys.JOINROWS, "Ctrl+Shift+W",
      "Select the cells of a column whose text spilled onto extra rows, then press KEY. " +
      "Rows with values only in that column are added to the row above and deleted.");
  } catch (e) {
    fillHint.textContent += " (couldn't confirm shortcuts: " + e.message + ")";
  }
}

// key is null when Excel's own command won a conflict, and missing when Excel is still
// using an older cached shortcuts.json that doesn't have this action yet.
function describeShortcut(el, key, expected, text) {
  if (key) {
    el.textContent = text.replace("KEY", key);
    return;
  }
  el.textContent = key === null
    ? "Excel has given " + expected + " to its own command, not Paint Mode. Type \"Reset Office Add-ins shortcut preferences\" " +
      "in Excel's search box at the top and run it, then press " + expected + " on a cell and choose Paint Mode."
    : "Excel hasn't loaded the " + expected + " shortcut yet. Close Excel, run the Paint Mode installer again, then reopen Excel.";
  el.classList.add("warn");
}

async function paintSelection() {
  try {
    await Excel.run(async (context) => {
      // getSelectedRanges also covers Ctrl+click multi-selections.
      const ranges = context.workbook.getSelectedRanges();
      ranges.load("address");
      if (currentColor === NO_FILL) {
        ranges.format.fill.clear();
      } else {
        ranges.format.fill.color = currentColor;
      }
      await context.sync();
      showStatus("Filled " + ranges.address.replace(/[^,!]*!/g, "") + " at " + new Date().toLocaleTimeString());
    });
    showError("");
  } catch (e) {
    // e.g. a chart or shape is selected, so there is no cell range to paint.
    showError(e.message);
  }
}

async function joinSplitRows() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const selection = context.workbook.getSelectedRange(); // throws for Ctrl+click multi-selections
      const used = sheet.getUsedRange(true);
      selection.load("rowIndex,rowCount,columnIndex,columnCount");
      used.load("rowIndex,rowCount,columnIndex,columnCount");
      await context.sync();

      // Trim whole-column selections to the data.
      const usedBottom = used.rowIndex + used.rowCount - 1;
      const usedRight = used.columnIndex + used.columnCount - 1;
      const top = Math.max(selection.rowIndex, used.rowIndex);
      const bottom = Math.min(selection.rowIndex + selection.rowCount - 1, usedBottom);
      const selLeft = Math.max(selection.columnIndex, used.columnIndex);
      const selRight = Math.min(selection.columnIndex + selection.columnCount - 1, usedRight);
      if (top > bottom || selLeft > selRight) {
        showStatus("Nothing to join: the selected cells are empty.");
        showError("");
        return;
      }
      if (selLeft === used.columnIndex && selRight === usedRight) {
        throw new Error("Select cells in just the column that spilled over, not every column.");
      }

      // Read whole rows, since a row is only deleted if everything outside the selected column is
      // blank. The row above the selection is read too, so selecting just the spilled-over cells works.
      const first = Math.max(top - 1, used.rowIndex);
      const block = sheet.getRangeByIndexes(first, used.columnIndex, bottom - first + 1, used.columnCount);
      block.load("values");
      await context.sync();

      const { writes, deletes } = planJoin(block.values, selLeft - used.columnIndex, selRight - used.columnIndex);
      if (deletes.length === 0) {
        showStatus("Nothing to join: no selected row has values only in the selected column.");
        showError("");
        return;
      }

      for (const w of writes) {
        // Keep text like "= x" or "- item" from being read as a formula.
        block.getCell(w.row, w.col).values = [[/^[=+\-@]/.test(w.text) ? "'" + w.text : w.text]];
      }
      const runs = [];
      for (const r of deletes) {
        const last = runs[runs.length - 1];
        if (last && last.start + last.count === r) last.count++;
        else runs.push({ start: r, count: 1 });
      }
      // Bottom-up, so each deletion leaves the rows above it where they were.
      for (const run of runs.reverse()) {
        sheet.getRangeByIndexes(first + run.start, 0, run.count, 1).getEntireRow().delete(Excel.DeleteShiftDirection.up);
      }
      await context.sync();

      const joinedInto = new Set(writes.map((w) => w.row)).size;
      showStatus("Joined " + plural(deletes.length, "row") + " into " + plural(joinedInto, "row") +
        " at " + new Date().toLocaleTimeString());
    });
    showError("");
  } catch (e) {
    showError(e.message);
  }
}

// values: whole rows of the sheet; columns selFrom..selTo are the selected column(s).
// A row with values only in the selected columns is spill-over: its text is appended (space-separated)
// to the nearest row above it, and the row is deleted. A completely blank row ends a record.
// Returns the cells to rewrite and the row offsets to delete, top to bottom.
function planJoin(values, selFrom, selTo) {
  const isBlank = (v) => v === null || String(v).trim() === "";
  const inSelection = (c) => c >= selFrom && c <= selTo;
  const writes = new Map();
  const deletes = [];
  let anchor = -1;
  values.forEach((row, r) => {
    const insideBlank = row.every((v, c) => !inSelection(c) || isBlank(v));
    const outsideBlank = row.every((v, c) => inSelection(c) || isBlank(v));
    if (insideBlank && outsideBlank) {
      anchor = -1;
      return;
    }
    if (!outsideBlank || anchor < 0) {
      anchor = r;
      return;
    }
    for (let c = selFrom; c <= selTo; c++) {
      if (isBlank(row[c])) continue;
      const key = anchor + "," + c;
      const base = writes.has(key) ? writes.get(key).text : String(values[anchor][c]).trim();
      const piece = String(row[c]).trim();
      writes.set(key, { row: anchor, col: c, text: base ? base + " " + piece : piece });
    }
    deletes.push(r);
  });
  return { writes: [...writes.values()], deletes };
}

function plural(n, word) {
  return n + " " + word + (n === 1 ? "" : "s");
}

function buildPalettes() {
  // Theme grid is laid out like Excel's: base colours on top, then five shade rows beneath each.
  const columns = THEME_COLORS.map((base) => [base, ...shadesOf(base)]);
  const themeGrid = document.getElementById("themeGrid");
  for (let row = 0; row < 6; row++) {
    for (const column of columns) themeGrid.appendChild(makeSwatch(column[row]));
  }

  const standardGrid = document.getElementById("standardGrid");
  for (const color of STANDARD_COLORS) standardGrid.appendChild(makeSwatch(color));
}

function makeSwatch(color) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "swatch";
  btn.dataset.color = color;
  btn.style.background = color;
  btn.title = color;
  btn.addEventListener("click", () => setColor(color));
  return btn;
}

// Same tint/shade steps Excel's colour picker uses.
function shadesOf(hex) {
  const [h, s, l] = hexToHsl(hex);
  let steps;
  if (l > 0.99) steps = [-0.05, -0.15, -0.25, -0.35, -0.5];   // white
  else if (l < 0.01) steps = [0.5, 0.35, 0.25, 0.15, 0.05];   // black
  else if (l > 0.85) steps = [-0.1, -0.25, -0.5, -0.75, -0.9]; // near-white background
  else steps = [0.8, 0.6, 0.4, -0.25, -0.5];
  return steps.map((t) => hslToHex(h, s, t > 0 ? l + (1 - l) * t : l * (1 + t)));
}

function hexToHsl(hex) {
  const r = parseInt(hex.slice(1, 3), 16) / 255;
  const g = parseInt(hex.slice(3, 5), 16) / 255;
  const b = parseInt(hex.slice(5, 7), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToHex(h, s, l) {
  const hueToRgb = (p, q, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let r, g, b;
  if (s === 0) {
    r = g = b = l;
  } else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hueToRgb(p, q, h + 1 / 3);
    g = hueToRgb(p, q, h);
    b = hueToRgb(p, q, h - 1 / 3);
  }
  const toHex = (x) => Math.round(x * 255).toString(16).padStart(2, "0");
  return ("#" + toHex(r) + toHex(g) + toHex(b)).toUpperCase();
}

function setColor(color) {
  currentColor = color;
  try {
    localStorage.setItem("paintModeColor", color);
  } catch (e) {
    // Storage unavailable; the colour just won't be remembered next time.
  }
  render();
}

function loadColor() {
  try {
    return localStorage.getItem("paintModeColor") || "#FFFF00";
  } catch (e) {
    return "#FFFF00";
  }
}

function render() {
  for (const swatch of document.querySelectorAll(".swatch")) {
    swatch.classList.toggle("selected", swatch.dataset.color === currentColor);
  }

  const chip = document.getElementById("currentChip");
  const label = document.getElementById("currentLabel");
  if (currentColor === NO_FILL) {
    chip.className = "nofill";
    chip.style.background = "";
    label.textContent = "Fill colour: No fill";
  } else {
    chip.className = "";
    chip.style.background = currentColor;
    label.textContent = "Fill colour: " + currentColor;
    document.getElementById("custom").value = currentColor;
  }
}

function showStatus(message) {
  document.getElementById("status").textContent = message;
}

function showError(message) {
  document.getElementById("error").textContent = message;
}
