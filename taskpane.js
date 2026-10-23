// Paint Mode: Ctrl+Shift+Z fills the current selection with the chosen colour.

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

let currentColor = loadColor();

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  try {
    Office.actions.associate("APPLYFILL", paintSelection);
  } catch (e) {
    document.getElementById("hint").textContent =
      "Keyboard shortcuts aren't supported in this version of Excel.";
  }

  // Excel's shortcut only fires when the grid has focus; catch it here too for when the panel does.
  document.addEventListener("keydown", (e) => {
    if (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === "z") {
      e.preventDefault();
      paintSelection();
    }
  });

  buildPalettes();
  document.getElementById("noFill").addEventListener("click", () => setColor(NO_FILL));
  document.getElementById("custom").addEventListener("input", (e) => setColor(e.target.value.toUpperCase()));
  render();
});

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
