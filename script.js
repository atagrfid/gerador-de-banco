// script.js
document.addEventListener("DOMContentLoaded", () => {
  // Referências aos elementos
  const undoBtn = document.getElementById("undoBtn");
  const redoBtn = document.getElementById("redoBtn");
  const importSheetBtn = document.getElementById("importSheetBtn");
  const spreadsheetContainer = document.getElementById("spreadsheet-container");
  const fillHandle = document.getElementById("fill-handle");
  const addAutocompleteColBtn = document.getElementById("addAutocompleteCol");
  const pasteListBtn = document.getElementById("pasteList");
  const deleteRowsBtn = document.getElementById("deleteRowsBtn");
  const deleteColsBtn = document.getElementById("deleteColsBtn");
  const clearSheetBtn = document.getElementById("clearSheetBtn");
  const generateEPCBtn = document.getElementById("generateEPC");
  const copyEPCsBtn = document.getElementById("copyEPCsBtn");
  const exportTxtBtn = document.getElementById("exportTxtBtn");
  const exportCsvBtn = document.getElementById("exportCsvBtn");
  const exportXlsxBtn = document.getElementById("exportXlsxBtn");
  const pasteModalOverlay = document.getElementById("paste-modal-overlay");
  const modalColumnSelect = document.getElementById("modal-column-select");
  const modalListInput = document.getElementById("modal-list-input");
  const modalConfirmBtn = document.getElementById("modal-confirm-btn");
  const modalCancelBtn = document.getElementById("modal-cancel-btn");
  const closePasteModalBtn = pasteModalOverlay.querySelector(".close-btn");
  const importModalOverlay = document.getElementById("import-modal-overlay");
  const fileInput = document.getElementById("file-input");
  const fileNameDisplay = document.getElementById("file-name-display");
  const headerCheckbox = document.getElementById("header-checkbox");
  const modalImportConfirmBtn = document.getElementById(
    "modal-import-confirm-btn"
  );
  const modalImportCancelBtn = document.getElementById(
    "modal-import-cancel-btn"
  );
  const closeImportModalBtn = importModalOverlay.querySelector(".close-btn");
  const emptyState = document.getElementById("empty-state");
  const toast = document.getElementById("toast-notification");

  // Variáveis de estado
  const AUTOCOMPLETE_COL_NAME = "[AUTOCOMPLETE]";
  let headers = [];
  let data = [];
  let draggedColumnIndex = null;
  let isSelecting = false;
  let selectionStartCell = null;
  let selectedCells = new Set();
  let isFilling = false;
  let fillStartRange = null;
  let isCopyMode = false;
  let selectedFile = null;

  // Lógica de Histórico
  let history = [];
  let historyIndex = -1;

  // --- FUNÇÕES DE DESFAZER/REFAZER E TECLADO (COM CORREÇÃO) ---
  const updateUndoRedoButtons = () => {
    undoBtn.disabled = historyIndex <= 0;
    redoBtn.disabled = historyIndex >= history.length - 1;
  };
  const saveState = () => {
    history = history.slice(0, historyIndex + 1);
    const currentState = {
      headers: JSON.parse(JSON.stringify(headers)),
      data: JSON.parse(JSON.stringify(data)),
    };
    history.push(currentState);
    historyIndex++;
    updateUndoRedoButtons();
  };
  const loadState = (state) => {
    headers = JSON.parse(JSON.stringify(state.headers));
    data = JSON.parse(JSON.stringify(state.data));
    renderSheet();
  };
  const undo = () => {
    if (historyIndex > 0) {
      historyIndex--;
      loadState(history[historyIndex]);
      updateUndoRedoButtons();
      showToast("Ação desfeita", "info");
    }
  };
  const redo = () => {
    if (historyIndex < history.length - 1) {
      historyIndex++;
      loadState(history[historyIndex]);
      updateUndoRedoButtons();
      showToast("Ação refeita", "info");
    }
  };

  const handleKeyDown = (e) => {
    const activeElement = document.activeElement;
    const isEditingCell = activeElement && activeElement.isContentEditable;
    const isEditingInput =
      activeElement &&
      (activeElement.tagName === "INPUT" ||
        activeElement.tagName === "TEXTAREA");

    if (e.ctrlKey && e.key.toLowerCase() === "z") {
      e.preventDefault();
      undo();
    } else if (e.ctrlKey && e.key.toLowerCase() === "y") {
      e.preventDefault();
      redo();
    } else if (e.key === "Delete" || e.key === "Backspace") {
      // CORREÇÃO: Deixa a tecla funcionar normalmente se estivermos editando qualquer tipo de input
      if (isEditingCell || isEditingInput) {
        return;
      }
      e.preventDefault();
      clearSelectedCellsContent();
    } else if (e.key === "Control") {
      isCopyMode = true;
      fillHandle.classList.add("copy-mode");
    }
  };
  const handleKeyUp = (e) => {
    if (e.key === "Control") {
      isCopyMode = false;
      fillHandle.classList.remove("copy-mode");
    }
  };

  // (O resto do código permanece o mesmo)
  let toastTimeout;
  const showToast = (message, type = "success") => {
    clearTimeout(toastTimeout);
    toast.textContent = message;
    toast.className = "toast show";
    toast.classList.add(type);
    toastTimeout = setTimeout(() => {
      toast.classList.remove("show");
    }, 3000);
  };
  const checkEmptyState = () => {
    const table = spreadsheetContainer.querySelector("table");
    if (headers.length === 0) {
      if (table) table.style.display = "none";
      emptyState.style.display = "flex";
    } else {
      if (table) table.style.display = "";
      emptyState.style.display = "none";
    }
  };
  const addColumns = () => {
    const quantity = parseInt(prompt("Quantas colunas adicionar?", "1"));
    if (isNaN(quantity) || quantity <= 0) return;
    for (let i = 0; i < quantity; i++) {
      headers.push(`Coluna ${headers.length + 1}`);
      data.forEach((row) => row.push(""));
    }
    renderSheet();
    saveState();
  };
  const addRows = () => {
    if (headers.length === 0) {
      showToast("Adicione uma coluna primeiro.", "info");
      return;
    }
    const quantity = parseInt(prompt("Quantas linhas adicionar?", "1"));
    if (isNaN(quantity) || quantity <= 0) return;
    for (let i = 0; i < quantity; i++) {
      data.push(Array(headers.length).fill(""));
    }
    renderSheet();
    saveState();
  };
  const renderSheet = () => {
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const tbody = document.createElement("tbody");
    const headerRow = document.createElement("tr");
    if (headers.length > 0) {
      headerRow.innerHTML = `<th class="control-column"><input type="checkbox" id="select-all-checkbox"></th>`;
    } else {
      headerRow.innerHTML = `<th class="control-column">#</th>`;
    }
    headers.forEach((headerText, index) => {
      const th = document.createElement("th");
      const isAutocomplete = headerText === AUTOCOMPLETE_COL_NAME;
      const isEpc = headerText.toLowerCase() === "epc";
      th.draggable = !isEpc;
      th.dataset.colIndex = index;
      if (isAutocomplete) th.classList.add("autocomplete-column", "header");
      th.innerHTML = `<div class="header-content"><div>${
        !isEpc
          ? `<input type="checkbox" class="concat-checkbox" data-col-index="${index}">`
          : ""
      }<span class="header-name">${headerText}</span></div>${
        !isAutocomplete && !isEpc
          ? `<span class="delete-col-btn" data-col-index="${index}">&times;</span>`
          : ""
      }</div>`;
      headerRow.appendChild(th);
    });
    const addColCell = document.createElement("th");
    addColCell.className = "add-btn-cell";
    addColCell.innerHTML = `<button id="add-col-btn-dynamic" class="add-btn" title="Adicionar Coluna(s)">+</button>`;
    headerRow.appendChild(addColCell);
    thead.appendChild(headerRow);
    data.forEach((rowData, rowIndex) => {
      const tr = document.createElement("tr");
      const rowNumberCell = document.createElement("td");
      rowNumberCell.classList.add("control-column");
      rowNumberCell.textContent = rowIndex + 1;
      tr.appendChild(rowNumberCell);
      rowData.forEach((cellData, colIndex) => {
        const td = document.createElement("td");
        const header = headers[colIndex];
        const isAutocomplete = header === AUTOCOMPLETE_COL_NAME;
        const isEpc = header?.toLowerCase() === "epc";
        td.contentEditable = !isAutocomplete && !isEpc;
        td.dataset.row = rowIndex;
        td.dataset.col = colIndex;
        if (isEpc) td.classList.add("epc-column");
        if (isAutocomplete) {
          td.classList.add("autocomplete-column");
          td.textContent = "(zeros)";
        } else {
          td.textContent = cellData || "";
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    const addRowTr = document.createElement("tr");
    const addRowCell = document.createElement("td");
    addRowCell.className = "control-column add-btn-cell";
    addRowCell.innerHTML = `<button id="add-row-btn-dynamic" class="add-btn" title="Adicionar Linha(s)">+</button>`;
    addRowTr.appendChild(addRowCell);
    const emptySpanCell = document.createElement("td");
    emptySpanCell.colSpan = headers.length + 1;
    addRowTr.appendChild(emptySpanCell);
    tbody.appendChild(addRowTr);
    table.appendChild(thead);
    table.appendChild(tbody);
    spreadsheetContainer.innerHTML = "";
    spreadsheetContainer.appendChild(emptyState);
    spreadsheetContainer.appendChild(fillHandle);
    spreadsheetContainer.appendChild(table);
    attachAllListeners();
    checkEmptyState();
  };
  const attachAllListeners = () => {
    const addColBtnDynamic = document.getElementById("add-col-btn-dynamic");
    if (addColBtnDynamic)
      addColBtnDynamic.addEventListener("click", addColumns);
    const addRowBtnDynamic = document.getElementById("add-row-btn-dynamic");
    if (addRowBtnDynamic) addRowBtnDynamic.addEventListener("click", addRows);
    spreadsheetContainer
      .querySelectorAll('td[contenteditable="true"]')
      .forEach((cell) => {
        cell.addEventListener("blur", handleCellEdit);
        cell.addEventListener("mousedown", handleSelectionStart);
        cell.addEventListener("mouseover", handleSelectionOver);
      });
    document.addEventListener("mouseup", handleSelectionEnd);
    spreadsheetContainer
      .querySelectorAll("th:not(.control-column):not(.add-btn-cell)")
      .forEach((th) => {
        const headerName = th.querySelector(".header-name");
        if (headerName?.textContent.toLowerCase() === "epc") return;
        if (headerName?.textContent !== AUTOCOMPLETE_COL_NAME) {
          headerName?.addEventListener("dblclick", () =>
            handleHeaderRename(th)
          );
        }
        th.addEventListener("dragstart", handleDragStart);
        th.addEventListener("dragover", handleDragOver);
        th.addEventListener("dragleave", handleDragLeave);
        th.addEventListener("drop", handleDrop);
      });
    spreadsheetContainer.querySelectorAll(".delete-col-btn").forEach((btn) => {
      btn.addEventListener("click", handleDeleteColumn);
    });
    const selectAllCheckbox = document.getElementById("select-all-checkbox");
    const columnCheckboxes =
      spreadsheetContainer.querySelectorAll(".concat-checkbox");
    if (selectAllCheckbox) {
      selectAllCheckbox.addEventListener("click", () => {
        columnCheckboxes.forEach((cb) => {
          cb.checked = selectAllCheckbox.checked;
        });
      });
    }
    columnCheckboxes.forEach((cb) => {
      cb.addEventListener("click", () => {
        if (!selectAllCheckbox) return;
        selectAllCheckbox.checked = Array.from(columnCheckboxes).every(
          (box) => box.checked
        );
      });
    });
  };
  const updateFillHandlePosition = () => {
    if (selectedCells.size === 0) {
      fillHandle.style.display = "none";
      return;
    }
    let lastCell;
    selectedCells.forEach((cell) => (lastCell = cell));
    const containerRect = spreadsheetContainer.getBoundingClientRect();
    const cellRect = lastCell.getBoundingClientRect();
    fillHandle.style.left = `${
      cellRect.right - containerRect.left - 5 + spreadsheetContainer.scrollLeft
    }px`;
    fillHandle.style.top = `${
      cellRect.bottom - containerRect.top - 5 + spreadsheetContainer.scrollTop
    }px`;
    fillHandle.style.display = "block";
  };
  const startFill = (e) => {
    e.preventDefault();
    isFilling = true;
    fillStartRange = getSelectionRange();
  };
  const updateFill = (targetCell) => {
    if (!isFilling || !targetCell) return;
    spreadsheetContainer
      .querySelectorAll("td.fill-selection")
      .forEach((c) => c.classList.remove("fill-selection"));
    const { row, col } = targetCell.dataset;
    const endRow = parseInt(row);
    const endCol = parseInt(col);
    const fillMinRow = Math.min(fillStartRange.minRow, endRow);
    const fillMaxRow = Math.max(fillStartRange.maxRow, endRow);
    const fillMinCol = Math.min(fillStartRange.minCol, endCol);
    const fillMaxCol = Math.max(fillStartRange.maxCol, endCol);
    for (let r = fillMinRow; r <= fillMaxRow; r++) {
      for (let c = fillMinCol; c <= fillMaxCol; c++) {
        const cell = spreadsheetContainer.querySelector(
          `td[data-row='${r}'][data-col='${c}']`
        );
        if (cell && !cell.classList.contains("selected")) {
          cell.classList.add("fill-selection");
        }
      }
    }
  };
  const handleFillDoubleClick = () => {
    if (selectedCells.size === 0) return;
    let { minRow, maxRow, minCol, maxCol } = getSelectionRange();
    let contextCol = minCol > 0 ? minCol - 1 : maxCol + 1;
    if (contextCol < 0 || contextCol >= headers.length) {
      showToast(
        "Nenhuma coluna adjacente para preenchimento automático.",
        "info"
      );
      return;
    }
    let lastDataRow = -1;
    if (headers[contextCol] === AUTOCOMPLETE_COL_NAME) {
      lastDataRow = data.length - 1;
    } else {
      for (let i = data.length - 1; i >= 0; i--) {
        if (data[i][contextCol] && data[i][contextCol].trim() !== "") {
          lastDataRow = i;
          break;
        }
      }
    }
    if (lastDataRow <= maxRow) {
      showToast("Nenhum dado abaixo para preencher.", "info");
      return;
    }
    const sourceValues = getSourceValues({ minRow, maxRow, minCol, maxCol });
    for (let r = maxRow + 1; r <= lastDataRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const sourceRowIndex = (r - minRow - 1) % sourceValues.length;
        const sourceColIndex = c - minCol;
        const baseValue = sourceValues[sourceRowIndex][sourceColIndex];
        const increment = r - maxRow;
        const newValue = getFillValue(
          baseValue,
          increment,
          isCopyMode ? "copy" : "sequential"
        );
        data[r][c] = newValue;
      }
    }
    renderSheet();
    showToast("Preenchimento automático aplicado!");
    saveState();
  };
  const getFillValue = (baseValue, increment, mode) => {
    if (mode === "copy") {
      return baseValue;
    }
    const match = baseValue.match(/^(.*?)(\d+)$/);
    if (!match) return baseValue;
    const prefix = match[1];
    const numberStr = match[2];
    const newNumber = parseInt(numberStr) + increment;
    return prefix + String(newNumber).padStart(numberStr.length, "0");
  };
  const getSelectionRange = () => {
    let minRow = Infinity,
      maxRow = -1,
      minCol = Infinity,
      maxCol = -1;
    selectedCells.forEach((cell) => {
      const row = parseInt(cell.dataset.row);
      const col = parseInt(cell.dataset.col);
      if (row < minRow) minRow = row;
      if (row > maxRow) maxRow = row;
      if (col < minCol) minCol = col;
      if (col > maxCol) maxCol = col;
    });
    return { minRow, maxRow, minCol, maxCol };
  };
  const getSourceValues = (range) => {
    const sourceValues = [];
    for (let r = range.minRow; r <= range.maxRow; r++) {
      const rowValues = [];
      for (let c = range.minCol; c <= range.maxCol; c++) {
        rowValues.push(data[r][c]);
      }
      sourceValues.push(rowValues);
    }
    return sourceValues;
  };
  const endFill = () => {
    if (!isFilling) return;
    isFilling = false;
    const fillCells =
      spreadsheetContainer.querySelectorAll("td.fill-selection");
    if (fillCells.length === 0) return;
    const sourceValues = getSourceValues(fillStartRange);
    fillCells.forEach((cell) => {
      const { row, col } = cell.dataset;
      const r = parseInt(row);
      const c = parseInt(col);
      const sourceRowIndex = (r - fillStartRange.minRow) % sourceValues.length;
      const sourceColIndex =
        (c - fillStartRange.minCol) % sourceValues[0].length;
      const baseValue = sourceValues[sourceRowIndex][sourceColIndex];
      let increment = 0;
      if (r > fillStartRange.maxRow) increment = r - fillStartRange.maxRow;
      if (c > fillStartRange.maxCol) increment = c - fillStartRange.maxCol;
      if (r < fillStartRange.minRow) increment = r - fillStartRange.minRow;
      if (c < fillStartRange.minCol) increment = c - fillStartRange.minCol;
      const newValue = getFillValue(
        baseValue,
        increment,
        isCopyMode ? "copy" : "sequential"
      );
      data[r][c] = newValue;
    });
    renderSheet();
    saveState();
  };
  const handleSelectionStart = (e) => {
    if (e.target.closest(".control-column")) return;
    isSelecting = true;
    selectionStartCell = e.target;
    if (!e.shiftKey) {
      clearSelectionVisuals();
      selectedCells.clear();
    }
    updateSelection(e.target);
  };
  const handleSelectionOver = (e) => {
    if (isSelecting) {
      updateSelection(e.target);
    } else if (isFilling) {
      updateFill(e.target);
    }
  };
  const handleSelectionEnd = () => {
    isSelecting = false;
    endFill();
  };
  const updateSelection = (endCell) => {
    clearSelectionVisuals();
    selectedCells.clear();
    if (!selectionStartCell) return;
    const startRow = parseInt(selectionStartCell.dataset.row);
    const startCol = parseInt(selectionStartCell.dataset.col);
    const endRow = parseInt(endCell.dataset.row);
    const endCol = parseInt(endCell.dataset.col);
    for (
      let r = Math.min(startRow, endRow);
      r <= Math.max(startRow, endRow);
      r++
    ) {
      for (
        let c = Math.min(startCol, endCol);
        c <= Math.max(startCol, endCol);
        c++
      ) {
        const cell = spreadsheetContainer.querySelector(
          `td[data-row='${r}'][data-col='${c}']`
        );
        if (cell) {
          cell.classList.add("selected");
          selectedCells.add(cell);
        }
      }
    }
    updateFillHandlePosition();
  };
  const clearSelectionVisuals = () => {
    spreadsheetContainer
      .querySelectorAll("td.selected, td.fill-selection")
      .forEach((td) => td.classList.remove("selected", "fill-selection"));
    fillHandle.style.display = "none";
  };
  fillHandle.addEventListener("mousedown", startFill);
  fillHandle.addEventListener("dblclick", handleFillDoubleClick);
  generateEPCBtn.addEventListener("click", () => {
    const checkedCheckboxes = document.querySelectorAll(
      ".concat-checkbox:checked"
    );
    if (checkedCheckboxes.length === 0) {
      showToast("Selecione ao menos uma coluna.", "info");
      return;
    }
    const selectedIndexes = Array.from(checkedCheckboxes).map((cb) =>
      parseInt(cb.dataset.colIndex)
    );
    selectedIndexes.sort((a, b) => a - b);
    const autocompleteColIndex = headers.indexOf(AUTOCOMPLETE_COL_NAME);
    const autocompleteIsSelected =
      selectedIndexes.includes(autocompleteColIndex);
    let errorMessages = [];
    data.forEach((row, rowIndex) => {
      let currentLength = 0;
      for (const index of selectedIndexes) {
        if (index !== autocompleteColIndex) {
          currentLength += (row[index] || "").length;
        }
      }
      if (currentLength > 24) {
        errorMessages.push(
          `- Linha ${
            rowIndex + 1
          }: Comprimento (${currentLength}) excede 24 caracteres.`
        );
      } else if (currentLength < 24 && !autocompleteIsSelected) {
        errorMessages.push(
          `- Linha ${
            rowIndex + 1
          }: Comprimento (${currentLength}) < 24 e [AUTOCOMPLETE] não selecionado.`
        );
      }
    });
    if (errorMessages.length > 0) {
      alert(
        "Falha na geração dos EPCs. Corrija os seguintes erros:\n\n" +
          errorMessages.join("\n")
      );
      return;
    }
    let epcColumnIndex = headers.findIndex((h) => h.toLowerCase() === "epc");
    if (epcColumnIndex === -1) {
      headers.push("EPC");
      epcColumnIndex = headers.length - 1;
      data.forEach((r) => r.push(""));
    }
    data.forEach((row) => {
      let parts = [];
      let currentLength = 0;
      for (const index of selectedIndexes) {
        if (index === autocompleteColIndex) {
          parts.push(null);
        } else {
          const cellValue = row[index] || "";
          parts.push(cellValue);
          currentLength += cellValue.length;
        }
      }
      const paddingNeeded = 24 - currentLength;
      const paddingString = "0".repeat(paddingNeeded);
      const autocompletePartIndex = parts.indexOf(null);
      if (autocompletePartIndex !== -1) {
        parts[autocompletePartIndex] = paddingString;
      }
      const finalEPC = parts.join("");
      row[epcColumnIndex] = finalEPC;
    });
    showToast("EPCs gerados com sucesso!");
    renderSheet();
  });
  const openPasteModal = () => {
    const availableCols = headers.filter(
      (h) => h !== AUTOCOMPLETE_COL_NAME && h.toLowerCase() !== "epc"
    );
    if (availableCols.length === 0) {
      showToast("Adicione colunas de dados primeiro.", "info");
      return;
    }
    modalColumnSelect.innerHTML = "";
    availableCols.forEach((header) => {
      const option = document.createElement("option");
      option.value = headers.indexOf(header);
      option.textContent = header;
      modalColumnSelect.appendChild(option);
    });
    modalListInput.value = "";
    pasteModalOverlay.classList.remove("hidden");
  };
  const closePasteModal = () => {
    pasteModalOverlay.classList.add("hidden");
  };
  const confirmPaste = () => {
    const colIndex = parseInt(modalColumnSelect.value);
    const listData = modalListInput.value.trim();
    if (isNaN(colIndex)) {
      showToast("Coluna inválida.", "error");
      return;
    }
    if (!listData) {
      showToast("A lista de dados está vazia.", "info");
      return;
    }
    const items = listData.split("\n");
    items.forEach((item, index) => {
      if (data[index]) {
        data[index][colIndex] = item.trim();
      } else {
        const newRow = Array(headers.length).fill("");
        newRow[colIndex] = item.trim();
        data.push(newRow);
      }
    });
    renderSheet();
    closePasteModal();
    showToast("Lista colada com sucesso!");
    saveState();
  };
  pasteListBtn.addEventListener("click", openPasteModal);
  modalConfirmBtn.addEventListener("click", confirmPaste);
  modalCancelBtn.addEventListener("click", closePasteModal);
  closePasteModalBtn.addEventListener("click", closePasteModal);
  pasteModalOverlay.addEventListener("click", (e) => {
    if (e.target === pasteModalOverlay) {
      closePasteModal();
    }
  });
  const downloadFile = (filename, content, mimeType) => {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
  const getEpcData = () => {
    const epcColIndex = headers.findIndex((h) => h.toLowerCase() === "epc");
    if (epcColIndex === -1) {
      showToast("Gere os EPCs primeiro.", "info");
      return null;
    }
    const epcs = data.map((row) => row[epcColIndex]).filter(Boolean);
    if (epcs.length === 0) {
      showToast("Nenhum EPC para exportar.", "info");
      return null;
    }
    return epcs;
  };
  const handleCopyEPCs = () => {
    const epcs = getEpcData();
    if (!epcs) return;
    navigator.clipboard
      .writeText(epcs.join("\n"))
      .then(() => {
        showToast("EPCs copiados com sucesso!");
      })
      .catch((err) => {
        showToast("Erro ao copiar.", "error");
      });
  };
  const handleExportTXT = () => {
    const epcs = getEpcData();
    if (!epcs) return;
    downloadFile("epcs.txt", epcs.join("\n"), "text/plain;charset=utf-8;");
    showToast("Arquivo .txt exportado!");
  };
  const handleExportCSV = () => {
    const epcs = getEpcData();
    if (!epcs) return;
    downloadFile("epcs.csv", epcs.join("\n"), "text/csv;charset=utf-8;");
    showToast("Arquivo .csv exportado!");
  };
  const handleExportXLSX = () => {
    const epcs = getEpcData();
    if (!epcs) return;
    const dataToExport = epcs.map((epc) => [epc]);
    const worksheet = XLSX.utils.aoa_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "EPCs");
    XLSX.writeFile(workbook, "epcs.xlsx");
  };
  copyEPCsBtn.addEventListener("click", handleCopyEPCs);
  exportTxtBtn.addEventListener("click", handleExportTXT);
  exportCsvBtn.addEventListener("click", handleExportCSV);
  addAutocompleteColBtn.addEventListener("click", () => {
    if (headers.includes(AUTOCOMPLETE_COL_NAME)) {
      showToast("A coluna Autocomplete já existe.", "info");
      return;
    }
    headers.push(AUTOCOMPLETE_COL_NAME);
    data.forEach((r) => r.push(""));
    renderSheet();
    saveState();
  });
  const handleDeleteColumn = (e) => {
    const colIndex = parseInt(e.target.dataset.colIndex);
    if (confirm(`Excluir a coluna "${headers[colIndex]}"?`)) {
      headers.splice(colIndex, 1);
      data.forEach((row) => row.splice(colIndex, 1));
      renderSheet();
      saveState();
    }
  };
  const handleDeleteSelectedCols = () => {
    if (selectedCells.size === 0) {
      showToast("Selecione colunas para excluir.", "info");
      return;
    }
    const colsToDelete = new Set();
    selectedCells.forEach((cell) => {
      const colIndex = parseInt(cell.dataset.col);
      if (headers[colIndex] !== AUTOCOMPLETE_COL_NAME) {
        colsToDelete.add(colIndex);
      }
    });
    if (
      colsToDelete.size > 0 &&
      confirm(`Excluir as ${colsToDelete.size} colunas selecionadas?`)
    ) {
      const sortedCols = Array.from(colsToDelete).sort((a, b) => b - a);
      sortedCols.forEach((colIndex) => {
        headers.splice(colIndex, 1);
        data.forEach((row) => row.splice(colIndex, 1));
      });
      selectedCells.clear();
      renderSheet();
      saveState();
    }
  };
  deleteColsBtn.addEventListener("click", handleDeleteSelectedCols);
  const handleDeleteSelectedRows = () => {
    if (selectedCells.size === 0) {
      showToast("Selecione linhas para excluir.", "info");
      return;
    }
    const rowsToDelete = new Set();
    selectedCells.forEach((cell) =>
      rowsToDelete.add(parseInt(cell.dataset.row))
    );
    if (confirm(`Excluir as ${rowsToDelete.size} linhas selecionadas?`)) {
      const sortedRows = Array.from(rowsToDelete).sort((a, b) => b - a);
      sortedRows.forEach((rowIndex) => data.splice(rowIndex, 1));
      selectedCells.clear();
      renderSheet();
      saveState();
    }
  };
  deleteRowsBtn.addEventListener("click", handleDeleteSelectedRows);
  const handleClearSheet = () => {
    if (headers.length === 0 && data.length === 0) {
      showToast("A planilha já está vazia.", "info");
      return;
    }
    if (
      confirm(
        "Tem certeza que deseja apagar todos os dados da planilha? Esta ação não pode ser desfeita."
      )
    ) {
      headers = [];
      data = [];
      selectedCells.clear();
      renderSheet();
      showToast("Planilha limpa com sucesso!");
      saveState();
    }
  };
  clearSheetBtn.addEventListener("click", handleClearSheet);
  const handleCellEdit = (e) => {
    const { row, col } = e.target.dataset;
    if (
      data[row] &&
      data[row][col] !== undefined &&
      data[row][col] !== e.target.textContent
    ) {
      data[row][col] = e.target.textContent;
      saveState();
    }
  };
  const handleHeaderRename = (th) => {
    clearSelectionVisuals();
    const headerSpan = th.querySelector(".header-name");
    const originalName = headerSpan.textContent;
    const colIndex = th.dataset.colIndex;
    const input = document.createElement("input");
    input.type = "text";
    input.value = originalName;
    headerSpan.replaceWith(input);
    input.focus();
    const saveRename = () => {
      const newName = input.value.trim();
      if (newName && newName !== originalName && !headers.includes(newName)) {
        headers[colIndex] = newName;
        saveState();
      } else if (headers.includes(newName) && newName !== originalName) {
        showToast("Este nome de coluna já existe.", "error");
      }
      renderSheet();
    };
    input.addEventListener("blur", saveRename);
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") input.blur();
      if (e.key === "Escape") {
        renderSheet();
      }
    });
  };
  function handleDragStart(e) {
    const activeElement = document.activeElement;
    if (
      activeElement &&
      activeElement.isContentEditable &&
      activeElement.tagName === "TD"
    )
      handleCellEdit({ target: activeElement });
    draggedColumnIndex = parseInt(e.currentTarget.dataset.colIndex);
    e.dataTransfer.effectAllowed = "move";
  }
  function handleDragOver(e) {
    e.preventDefault();
    e.currentTarget.classList.add("drag-over");
  }
  function handleDragLeave(e) {
    e.currentTarget.classList.remove("drag-over");
  }
  function handleDrop(e) {
    e.preventDefault();
    e.currentTarget.classList.remove("drag-over");
    const sourceIndex = draggedColumnIndex;
    const targetIndex = parseInt(e.currentTarget.dataset.colIndex);
    if (sourceIndex === targetIndex) return;
    const [movedHeader] = headers.splice(sourceIndex, 1);
    headers.splice(targetIndex, 0, movedHeader);
    data.forEach((row) => {
      const [movedCell] = row.splice(sourceIndex, 1);
      row.splice(targetIndex, 0, movedCell);
    });
    renderSheet();
    saveState();
  }
  document.addEventListener("keydown", handleKeyDown);
  document.addEventListener("keyup", handleKeyUp);
  const clearSelectedCellsContent = () => {
    if (selectedCells.size === 0) return;
    let changed = false;
    selectedCells.forEach((cell) => {
      const { row, col } = cell.dataset;
      if (
        headers[col] !== AUTOCOMPLETE_COL_NAME &&
        headers[col].toLowerCase() !== "epc" &&
        data[row][col] !== ""
      ) {
        data[row][col] = "";
        cell.textContent = "";
        changed = true;
      }
    });
    if (changed) saveState();
  };
  const openImportModal = () => {
    if (
      data.length > 0 &&
      !confirm(
        "Importar um novo arquivo irá substituir os dados atuais. Deseja continuar?"
      )
    ) {
      return;
    }
    selectedFile = null;
    fileInput.value = "";
    fileNameDisplay.textContent = "Nenhum arquivo selecionado.";
    headerCheckbox.checked = true;
    importModalOverlay.classList.remove("hidden");
  };
  const closeImportModal = () => {
    importModalOverlay.classList.add("hidden");
  };
  const handleFileSelected = (event) => {
    selectedFile = event.target.files[0];
    if (selectedFile) {
      fileNameDisplay.textContent = selectedFile.name;
    }
  };
  const confirmImport = () => {
    if (!selectedFile) {
      showToast("Por favor, selecione um arquivo primeiro.", "info");
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const fileData = e.target.result;
        const workbook = XLSX.read(fileData, { type: "binary" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const importedData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        processAndLoadData(importedData, headerCheckbox.checked);
      } catch (error) {
        console.error("Erro ao ler o arquivo:", error);
        showToast("Falha ao importar o arquivo.", "error");
      }
    };
    reader.readAsBinaryString(selectedFile);
    closeImportModal();
  };
  const processAndLoadData = (importedData, hasHeader) => {
    if (!importedData || importedData.length === 0) {
      showToast(
        "O arquivo importado está vazio ou em um formato inválido.",
        "error"
      );
      return;
    }
    let newHeaders = [];
    let newData = [];
    if (hasHeader) {
      newHeaders = importedData[0].map((header) => String(header || ""));
      newData = importedData
        .slice(1)
        .map((row) =>
          row.map((cell) =>
            cell !== null && cell !== undefined ? String(cell) : ""
          )
        );
    } else {
      const numCols = importedData[0].length;
      newHeaders = Array.from({ length: numCols }, (_, i) => `Coluna ${i + 1}`);
      newData = importedData.map((row) =>
        row.map((cell) =>
          cell !== null && cell !== undefined ? String(cell) : ""
        )
      );
    }
    headers = newHeaders;
    data = newData;
    renderSheet();
    saveState();
    showToast("Planilha importada com sucesso!");
  };
  importSheetBtn.addEventListener("click", openImportModal);
  document
    .querySelector(".file-input-label")
    .addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", handleFileSelected);
  modalImportConfirmBtn.addEventListener("click", confirmImport);
  modalImportCancelBtn.addEventListener("click", closeImportModal);
  closeImportModalBtn.addEventListener("click", closeImportModal);

  // Inicialização
  saveState();
  renderSheet();
});
