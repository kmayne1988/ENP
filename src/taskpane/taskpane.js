let contextSheetName = null;
let pinnedSheets = JSON.parse(localStorage.getItem("pinnedSheets") || "[]");

Office.onReady(() => {
  Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;

    await listSheets();

    // Auto update on sheet add/delete/rename
    sheets.onAdded.add(() => listSheets());
    sheets.onDeleted.add(() => listSheets());
    sheets.onNameChanged.add(() => listSheets());

    await context.sync();
  });

  // Restore dark mode if previously set
  if (localStorage.getItem("darkMode") === "true") {
    document.body.classList.add("dark-mode");
    document.getElementById("dark-mode-toggle").checked = true;
  }
});

//
// Dark mode toggle
//
function toggleDarkMode() {
  document.body.classList.toggle("dark-mode");
  localStorage.setItem("darkMode", document.body.classList.contains("dark-mode"));
}
window.toggleDarkMode = toggleDarkMode;

//
// Main render
//
window.listSheets = async function () {
  await Excel.run(async (context) => {
    const showHidden = document.getElementById("show-hidden")?.checked;
    const onlyHidden = document.getElementById("only-hidden")?.checked;

    const sheets = context.workbook.worksheets;
    sheets.load("items/name,items/visibility,items/tabColor");
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("name");
    await context.sync();

    const listContainer = document.getElementById("sheet-list");
    listContainer.innerHTML = "";

    const pinned = sheets.items.filter((s) => pinnedSheets.includes(s.name));
    const unpinned = sheets.items.filter((s) => !pinnedSheets.includes(s.name));

    if (pinned.length > 0) {
      const pinnedHeader = document.createElement("div");
      pinnedHeader.textContent = "Pinned Sheets";
      pinnedHeader.className = "section-header";
      listContainer.appendChild(pinnedHeader);
      pinned.forEach((sheet) => renderSheetButton(sheet, activeSheet, showHidden, onlyHidden, listContainer));
    }

    if (unpinned.length > 0) {
      const otherHeader = document.createElement("div");
      otherHeader.textContent = "All Sheets";
      otherHeader.className = "section-header";
      listContainer.appendChild(otherHeader);
      unpinned.forEach((sheet) => renderSheetButton(sheet, activeSheet, showHidden, onlyHidden, listContainer));
    }
  });
};

//
// Render one sheet button
//
function renderSheetButton(sheet, activeSheet, showHidden, onlyHidden, container) {
  if (onlyHidden && sheet.visibility !== Excel.SheetVisibility.hidden) return;
  if (!onlyHidden && !showHidden && sheet.visibility === Excel.SheetVisibility.hidden) return;

  const btn = document.createElement("button");
  btn.textContent = sheet.visibility === Excel.SheetVisibility.hidden
    ? `👻 ${sheet.name}`
    : sheet.name;

  btn.className = "sheet-btn";
  if (pinnedSheets.includes(sheet.name)) btn.classList.add("pinned");
  if (sheet.name === activeSheet.name) btn.classList.add("active");
  if (sheet.visibility === Excel.SheetVisibility.hidden) btn.classList.add("hidden");

  // Tab color stripe
  if (sheet.tabColor) {
    btn.style.borderLeft = `6px solid ${sheet.tabColor}`;
  } else {
    btn.style.borderLeft = "6px solid transparent";
  }

  // Left click → activate or unhide+activate
  btn.onclick = async () => {
    if (sheet.visibility === Excel.SheetVisibility.hidden) {
      if (confirm(`"${sheet.name}" is hidden. Unhide it?`)) {
        await Excel.run(async (ctx) => {
          const s = ctx.workbook.worksheets.getItem(sheet.name);
          s.visibility = Excel.SheetVisibility.visible;
          s.activate();
          await ctx.sync();
        });
        await listSheets();
      }
    } else {
      await Excel.run(async (ctx) => {
        const s = ctx.workbook.worksheets.getItem(sheet.name);
        s.activate();
        await ctx.sync();
      });
      await listSheets();
    }
  };

  // Right click → context menu
  btn.oncontextmenu = (ev) => {
    ev.preventDefault();
    contextSheetName = sheet.name;
    const menu = document.getElementById("context-menu");
    menu.style.display = "block";
    menu.style.top = ev.clientY + "px";
    menu.style.left = ev.clientX + "px";
  };

  container.appendChild(btn);
}

//
// Search filter
//
window.filterSheets = function () {
  const query = document.getElementById("sheet-search").value.toLowerCase();
  const buttons = document.querySelectorAll("#sheet-list .sheet-btn");
  buttons.forEach((btn) => {
    btn.style.display = btn.textContent.toLowerCase().includes(query) ? "block" : "none";
  });
};

//
// Context menu actions
//
window.contextAction = async function (action) {
  if (!contextSheetName) return;

  await Excel.run(async (context) => {
    let sheet;
    try {
      sheet = context.workbook.worksheets.getItem(contextSheetName);
      sheet.load("name, visibility");
      await context.sync();
    } catch {
      console.warn("Sheet not found:", contextSheetName);
      return;
    }

    switch (action) {
      case "activate":
        sheet.activate();
        break;

      case "hide":
        sheet.visibility = Excel.SheetVisibility.hidden;
        break;

      case "unhide":
        sheet.visibility = Excel.SheetVisibility.visible;
        break;

      case "pin":
        if (!pinnedSheets.includes(sheet.name)) {
          pinnedSheets.push(sheet.name);
        } else {
          pinnedSheets = pinnedSheets.filter((s) => s !== sheet.name);
        }
        localStorage.setItem("pinnedSheets", JSON.stringify(pinnedSheets));
        break;
    }

    await context.sync();
  });

  await listSheets();

  document.getElementById("context-menu").style.display = "none";
};

//
// Hide context menu when clicking elsewhere
//
document.addEventListener("click", () => {
  document.getElementById("context-menu").style.display = "none";
});
