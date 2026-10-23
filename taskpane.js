// Paint Mode: while on, every cell/range you select gets filled with the chosen colour.

const NO_FILL = "none";

// Excel's "Standard Colors" row plus a few common highlight tints.
const PALETTE = [
  "#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050",
  "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0",
  "#FFC7CE", "#FFEB9C", "#C6EFCE", "#DDEBF7", "#D9D9D9",
];

let paintOn = false;
let currentColor = loadColor();

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  try {
    Office.actions.associate("TOGGLEPAINTMODE", togglePaintMode);
  } catch (e) {
    document.getElementById("hint").textContent =
      "Keyboard shortcuts aren't supported in this version of Excel. Use the button above.";
  }

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSelectionChanged
  );

  buildPalette();
  document.getElementById("toggle").addEventListener("click", togglePaintMode);
  document.getElementById("custom").addEventListener("input", (e) => setColor(e.target.value.toUpperCase()));
  render();
});

function togglePaintMode() {
  paintOn = !paintOn;
  render();
}

async function onSelectionChanged() {
  if (!paintOn) return;
  try {
    await Excel.run(async (context) => {
      // getSelectedRanges also covers Ctrl+click multi-selections.
      const fill = context.workbook.getSelectedRanges().format.fill;
      if (currentColor === NO_FILL) {
        fill.clear();
      } else {
        fill.color = currentColor;
      }
      await context.sync();
    });
    showError("");
  } catch (e) {
    // e.g. a chart or shape is selected, so there is no cell range to paint.
    showError(e.message);
  }
}

function buildPalette() {
  const palette = document.getElementById("palette");
  for (const color of [...PALETTE, NO_FILL]) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "swatch";
    btn.dataset.color = color;
    if (color === NO_FILL) {
      btn.classList.add("nofill");
      btn.title = "No fill (eraser)";
    } else {
      btn.style.background = color;
      btn.title = color;
    }
    btn.addEventListener("click", () => setColor(color));
    palette.appendChild(btn);
  }
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
  const toggle = document.getElementById("toggle");
  toggle.textContent = paintOn ? "Paint mode: ON" : "Paint mode: OFF";
  toggle.classList.toggle("on", paintOn);

  for (const swatch of document.querySelectorAll(".swatch")) {
    swatch.classList.toggle("selected", swatch.dataset.color === currentColor);
  }

  const chip = document.getElementById("currentChip");
  if (currentColor === NO_FILL) {
    chip.className = "nofill";
    chip.style.background = "";
    document.getElementById("currentLabel").textContent = "No fill (eraser)";
  } else {
    chip.className = "";
    chip.style.background = currentColor;
    document.getElementById("currentLabel").textContent = currentColor;
    document.getElementById("custom").value = currentColor;
  }
}

function showError(message) {
  document.getElementById("error").textContent = message;
}
