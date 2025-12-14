"use client";

import React, { useState } from "react";
import * as XLSX from "xlsx-js-style";
import { X, Menu, Layout, Zap, Sheet, Plus } from "lucide-react"; // Import Icons
import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}
import SheetPreview from "@/components/SheetPreview";
import ControlPanel from "@/components/ControlPanel";
import { Chart as ChartJS, CategoryScale, LinearScale, BarElement, LineElement, PointElement, Title, Tooltip, Legend } from 'chart.js';
import { Bar, Line, Scatter } from 'react-chartjs-2';
import * as ss from 'simple-statistics';

ChartJS.register(CategoryScale, LinearScale, BarElement, LineElement, PointElement, Title, Tooltip, Legend);

import { applyRowMath, applyColMath, applyStyleUtil, applyMove, applyColConst, applyLogicIf, applyRangeMath } from "@/utils/recipeActions";

// Helper for default empty grid (20 rows x 10 cols)
const createDefaultGrid = () => {
  return Array.from({ length: 20 }, () => Array(10).fill(""));
};

export default function Home() {
  // Multi-Sheet State: Initialize with Default Grid
  const [worksheets, setWorksheets] = useState<Record<string, any[][]>>({ "Sheet1": createDefaultGrid() });
  const [currentSheetName, setCurrentSheetName] = useState<string>("Sheet1");

  // Derived State (for compatibility)
  const excelData = worksheets[currentSheetName] || [];

  const [fileName, setFileName] = useState<string | null>(null);
  const [isFileLoaded, setIsFileLoaded] = useState(true); // Default to True (Start Editing Immediately)
  const [alignment, setAlignment] = useState<"left" | "center" | "right" | "auto">("auto");

  // Theme State
  const [isDarkMode, setIsDarkMode] = useState(false);

  // UX State: History, Clipboard, Selection (History tracks Workbook State)
  const [history, setHistory] = useState<{ sheets: Record<string, any[][]>; current: string }[]>([]);
  const [future, setFuture] = useState<{ sheets: Record<string, any[][]>; current: string }[]>([]);
  const [clipboard, setClipboard] = useState<any[][] | null>(null);
  const [selection, setSelection] = useState<{ start: { r: number; c: number }; end: { r: number; c: number } } | null>(null);
  const [analysisResult, setAnalysisResult] = useState<any>(null); // { type, data, stats, options }
  const [isAnalysisOpen, setIsAnalysisOpen] = useState(false);

  // Mobile State
  const [isMobileSidebarOpen, setIsMobileSidebarOpen] = useState(false);
  // Removed isMobilePanelOpen (Lite Mode Logic)

  // File System Handles
  const [fileHandle, setFileHandle] = useState<FileSystemFileHandle | null>(null);

  // Target Pick Logic
  const [isPickingTarget, setIsPickingTarget] = useState(false);
  const [pendingAction, setPendingAction] = useState<any>(null);

  // Single Click: Just Visual Selection (Handled by SheetPreview internal or onSelectionChange)
  const handleTargetPick = (r: number, c: number) => {
    // Just select the cell to show user where they clicked
    setSelection({ start: { r, c }, end: { r, c } });
  };

  // Double Click: Confirm Selection
  const handleTargetConfirm = (r: number, c: number) => {
    if (!pendingAction) return;
    setIsPickingTarget(false);

    const newAction = {
      ...pendingAction,
      option: {
        ...pendingAction.option,
        resultMode: 'pick_execute',
        targetStart: { r, c }
      }
    };
    handleAction(newAction);
    setPendingAction(null);
  };

  // Wrapper for Setting Data (Updates Current Sheet)
  const updateData = (newData: any[][], saveHistory = true) => {
    const newWorksheets = { ...worksheets, [currentSheetName]: newData };

    if (saveHistory) {
      setHistory(prev => [...prev.slice(-19), { sheets: worksheets, current: currentSheetName }]);
      setFuture([]);
    }
    setWorksheets(newWorksheets);
  };

  // Wrapper for Setting Workbook (Updates All Sheets)
  const updateWorkbook = (newWorksheets: Record<string, any[][]>, newCurrentSheet: string, saveHistory = true) => {
    if (saveHistory) {
      setHistory(prev => [...prev.slice(-19), { sheets: worksheets, current: currentSheetName }]);
      setFuture([]); // Clear future
    }
    setWorksheets(newWorksheets);
    setCurrentSheetName(newCurrentSheet);
  };

  const handleUndo = () => {
    if (history.length === 0) return;
    const previous = history[history.length - 1];
    const newHistory = history.slice(0, -1);

    setFuture(prev => [{ sheets: worksheets, current: currentSheetName }, ...prev]);
    setWorksheets(previous.sheets);
    setCurrentSheetName(previous.current);
    setHistory(newHistory);
  };

  const handleRedo = () => {
    if (future.length === 0) return;
    const next = future[0];
    const newFuture = future.slice(1);

    setHistory(prev => [...prev, { sheets: worksheets, current: currentSheetName }]);
    setWorksheets(next.sheets);
    setCurrentSheetName(next.current);
    setFuture(newFuture);
  };

  const handleCopy = () => {
    if (!selection || excelData.length === 0) return;

    // Determine range
    const r1 = Math.min(selection.start.r, selection.end.r);
    const r2 = Math.max(selection.start.r, selection.end.r);
    const c1 = Math.min(selection.start.c, selection.end.c);
    const c2 = Math.max(selection.start.c, selection.end.c);

    const copiedData = [];
    for (let r = r1; r <= r2; r++) {
      const row = [];
      for (let c = c1; c <= c2; c++) {
        row.push(excelData[r]?.[c] ?? null);
      }
      copiedData.push(row);
    }

    setClipboard(copiedData);
    console.log("Copied to internal clipboard", copiedData);
  };

  const handlePaste = () => {
    if (!clipboard || !selection) return;

    const rStart = Math.min(selection.start.r, selection.end.r);
    const cStart = Math.min(selection.start.c, selection.end.c);

    const newData = excelData.map(row => [...(row || [])]);

    clipboard.forEach((row, rIndex) => {
      const targetR = rStart + rIndex;
      if (!newData[targetR]) newData[targetR] = [];

      row.forEach((cellValue, cIndex) => {
        const targetC = cStart + cIndex;
        newData[targetR][targetC] = cellValue;
      });
    });

    updateData(newData);
  };

  const handleOpenFile = async () => {
    try {
      if (!('showOpenFilePicker' in window)) {
        alert("이 브라우저는 파일 열기 API를 지원하지 않습니다. 파일을 드래그해서 넣어주세요.");
        return;
      }

      const [handle] = await window.showOpenFilePicker({
        types: [{
          description: 'Excel Files',
          accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
        }],
        multiple: false
      });

      const file = await handle.getFile();
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });

      const newWorksheets: Record<string, any[][]> = {};
      workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        newWorksheets[sheetName] = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
      });

      const firstSheet = workbook.SheetNames[0];

      setFileHandle(handle);
      setFileName(file.name);
      updateWorkbook(newWorksheets, firstSheet, true);

    } catch (err: any) {
      if (err.name !== 'AbortError') {
        console.error("File open failed:", err);
        alert("파일을 여는 데 실패했습니다.");
      }
    }
  };

  const handleInstantSave = async () => {
    if (!fileHandle) {
      // Fallback to Export (Save As) if no file is linked
      if (confirm("열린 파일이 없습니다. 새 파일로 저장하시겠습니까?")) {
        handleDownload();
      }
      return;
    }

    try {
      // Create a writable stream to the file
      // @ts-ignore - FileSystemFileHandle type might be missing in strict setup
      const writable = await fileHandle.createWritable();

      // Convert current state to workbook
      const wb = XLSX.utils.book_new();
      Object.keys(worksheets).forEach(sheet => {
        const ws = XLSX.utils.aoa_to_sheet(worksheets[sheet]);
        XLSX.utils.book_append_sheet(wb, ws, sheet);
      });

      // Write data
      const data = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      await writable.write(data);
      await writable.close();

      alert("저장되었습니다! (Overwrite)");
    } catch (err) {
      console.error("Save failed:", err);
      alert("저장에 실패했습니다.");
    }
  };

  const handleFileSelect = async (file: File) => {
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });

      const newWorksheets: Record<string, any[][]> = {};
      workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        newWorksheets[sheetName] = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
      });

      const firstSheet = workbook.SheetNames[0];

      setFileHandle(null);
      updateWorkbook(newWorksheets, firstSheet, true);
      setFileName(file.name);
    } catch (error) {
      console.error("Error reading excel file:", error);
      alert("파일을 읽는 도중 오류가 발생했습니다.");
    }
  };

  // Helper for default empty grid
  // (Moved to top-level)


  const handleStartEmpty = () => {
    updateWorkbook({ "Sheet1": createDefaultGrid() }, "Sheet1", false);
    setIsFileLoaded(true);
  };

  const handleClear = () => {
    updateWorkbook({ "Sheet1": createDefaultGrid() }, "Sheet1");
    setFileName(null);
    setFileHandle(null);
  };

  // --- Sheet Management ---
  const handleAddSheet = () => {
    const existingNames = Object.keys(worksheets);
    let newIndex = existingNames.length + 1;
    let newName = `Sheet${newIndex}`;
    while (existingNames.includes(newName)) {
      newIndex++;
      newName = `Sheet${newIndex}`;
    }

    updateWorkbook({ ...worksheets, [newName]: [] }, newName);
  };

  const handleDeleteSheet = (sheetName: string) => {
    const keys = Object.keys(worksheets);
    if (keys.length <= 1) {
      alert("최소 한 개의 시트는 있어야 해요!");
      return;
    }

    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    const { [sheetName]: _deleted, ...remaining } = worksheets;
    const remainingKeys = Object.keys(remaining);
    // If we deleted current sheet, switch to first available
    const nextSheet = currentSheetName === sheetName ? remainingKeys[0] : currentSheetName;

    updateWorkbook(remaining, nextSheet);
  };

  const handleChangeSheet = (sheetName: string) => {
    setCurrentSheetName(sheetName);
    setSelection(null); // Clear selection on switch
  };

  const handleSaveAnalysisToSheet = () => {
    if (!analysisResult || analysisResult.type !== 'stat') {
      alert("현재 통계 결과만 시트로 저장할 수 있습니다."); // Feature limitation for now
      return;
    }

    const { stats } = analysisResult;
    // Create Table Data
    const reportData = [
      ["항목", "값"],
      ["데이터 개수", stats.count],
      ["합계", stats.sum],
      ["평균", stats.mean],
      ["중앙값", stats.median],
      ["최솟값", stats.min],
      ["최댓값", stats.max],
      ["표준편차", stats.stdDev],
      [],
      ["생성 일시", new Date().toLocaleString()]
    ];

    const existingNames = Object.keys(worksheets);
    let newIndex = 1;
    let newName = `Analysis_${newIndex}`;
    while (existingNames.includes(newName)) {
      newIndex++;
      newName = `Analysis_${newIndex}`;
    }

    alert(`분석 결과를 '${newName}' 시트에 저장했습니다! ✅`);
    updateWorkbook({ ...worksheets, [newName]: reportData }, newName);
    setIsAnalysisOpen(false);
  };

  const handleCleanEmptyRows = () => {
    if (excelData.length === 0) {
      alert("데이터가 없어서 청소할 게 없어요! 😅");
      return;
    }
    const cleanedData = excelData.filter(row => {
      if (!row || row.length === 0) return false;
      return row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== "");
    });
    const removedCount = excelData.length - cleanedData.length;
    if (removedCount === 0) {
      alert("이미 깨끗해서 청소할 빈 줄이 없어요! ✨");
    } else {
      updateData(cleanedData);
      alert(`총 ${removedCount}개의 빈 줄을 찾아 싹 청소했어요! 🧹`);
    }
  };

  // Helper to generate workbook buffer
  // Excel Formula Helpers
  const indexToA1Notation = (r: number, c: number) => {
    let label = "";
    let i = c;
    while (i >= 0) {
      label = String.fromCharCode((i % 26) + 65) + label;
      i = Math.floor(i / 26) - 1;
    }
    return `${label}${r + 1}`;
  };

  const generateExcelFormula = (formula: any, currentR: number, currentC: number) => {
    const { type, startC, endC, startR, endR, operator } = formula;

    // Map Operators to Excel Functions
    // Fallback to value if no mapping found
    let excelFunc = "";
    if (operator === '+') excelFunc = "SUM";
    else if (operator === 'avg') excelFunc = "AVERAGE"; // Internal 'avg' might be passed as string if my recipe used it
    // Wait, 'handleMath...' uses standard operators (+, -, *, /) or avg/count logic from 'recipe'?
    // Actually 'handleMathColCol' usually receives basic operators (+, -, *, /).
    // Recipe types like 'sum', 'avg' in 'handleCalculate' are different from 'handleMathColCol'.
    // 'handleMathColCol' constructs 'row_sum' with 'operator'.
    // If operator is simple (+), we use SUM.
    // If operator is *, we use PRODUCT?
    // If operator is -, Excel SUM doesn't subtract.

    if (['+', 'sum'].includes(operator)) excelFunc = "SUM";
    else if (['*', 'mul'].includes(operator)) excelFunc = "PRODUCT";
    else if (['avg', 'mean'].includes(operator)) excelFunc = "AVERAGE";
    else if (['count', 'cnt'].includes(operator)) excelFunc = "COUNT";
    else if (['max'].includes(operator)) excelFunc = "MAX";
    else if (['min'].includes(operator)) excelFunc = "MIN";

    if (type === 'logic_if') {
      const sourceRef = XLSX.utils.encode_cell({ r: formula.targetR, c: formula.targetC });
      const val = isNaN(Number(formula.compareVal)) ? `"${formula.compareVal}"` : formula.compareVal;
      return `IF(${sourceRef}${formula.operator}${val}, "${formula.trueVal}", "${formula.falseVal}")`;
    }
    else if (type === 'logic_vlookup') {
      const sourceRef = XLSX.utils.encode_cell({ r: formula.targetR, c: formula.targetC });
      return `VLOOKUP(${sourceRef}, ${formula.rangeStr}, ${formula.colIndex}, 0)`;
    }

    if (!excelFunc) return null; // No mapping, use static value

    if (type === 'row_sum') {
      const startRef = XLSX.utils.encode_cell({ r: currentR, c: startC });
      const endRef = XLSX.utils.encode_cell({ r: currentR, c: endC });
      return `${excelFunc}(${startRef}:${endRef})`;
    }
    else if (type === 'col_sum') {
      const startRef = XLSX.utils.encode_cell({ r: startR, c: currentC });
      const endRef = XLSX.utils.encode_cell({ r: endR, c: currentC });
      return `${excelFunc}(${startRef}:${endRef})`;
    }
    else if (type === 'range_agg') {
      const startRef = XLSX.utils.encode_cell({ r: startR, c: startC });
      const endRef = XLSX.utils.encode_cell({ r: endR, c: endC });
      return `${excelFunc}(${startRef}:${endRef})`;
    }

    return null;
  };

  const getWorkbookBuffer = () => {
    const wb = XLSX.utils.book_new();

    Object.keys(worksheets).forEach(sheetName => {
      const sheetData = worksheets[sheetName];
      const ws = XLSX.utils.aoa_to_sheet(sheetData);

      // Apply Styles & Formulas
      const range = XLSX.utils.decode_range(ws['!ref'] || "A1");
      for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
          const cell = ws[cellAddress];

          // Access original data logic to find formula
          const originalCell = sheetData[R]?.[C];

          if (!cell && !originalCell) continue;

          // Logic: SheetJS might not have created a cell if value was null/undefined
          // But if we want to add style/formula to empty cell, we need to create it.
          // However, formula usually implies value exists.

          if (!ws[cellAddress]) {
            // Only create if we have something to add? 
            // If originalCell had formula but no value? Unlikely.
            if (originalCell) ws[cellAddress] = { v: "", t: 's' };
          }
          const activeCell = ws[cellAddress];
          if (!activeCell) continue;

          // 1. Inject Formula (Priority)
          if (originalCell && typeof originalCell === 'object' && originalCell.formula) {
            const excelFormula = generateExcelFormula(originalCell.formula, R, C);
            if (excelFormula) {
              activeCell.f = excelFormula;
              activeCell.t = 'n'; // Ensure type is number so Excel calculates it
            }
          }

          // 2. Apply Styles
          if (!activeCell.s) activeCell.s = {};
          if (!activeCell.s.alignment) activeCell.s.alignment = {};

          // Check for original style preference (e.g. from highlight)
          if (originalCell && typeof originalCell === 'object' && originalCell.s) {
            activeCell.s = { ...activeCell.s, ...originalCell.s };
          }

          if (alignment !== 'auto') {
            activeCell.s.alignment.horizontal = alignment;
          } else {
            // Preserve existing alignment if set by highlight, else default
            if (!activeCell.s.alignment.horizontal) {
              activeCell.s.alignment.horizontal = (activeCell.t === 'n') ? 'right' : 'left';
            }
          }

          if (!activeCell.s.border) {
            activeCell.s.border = {
              top: { style: "thin", color: { rgb: "C0C0C0" } },
              bottom: { style: "thin", color: { rgb: "C0C0C0" } },
              left: { style: "thin", color: { rgb: "C0C0C0" } },
              right: { style: "thin", color: { rgb: "C0C0C0" } }
            };
          }
        }
      }
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    // Generate buffer
    return XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  };

  // Save Logic (Hybrid)
  const handleSave = async () => {
    // Check if workbook is empty? Not really needed as we always have at least one sheet.

    // CASE A: We have a handle -> Overwrite
    if (fileHandle) {
      try {
        const writable = await fileHandle.createWritable();
        const buffer = getWorkbookBuffer();
        await writable.write(buffer);
        await writable.close();
        alert("✅ 원본 파일에 저장했습니다!");
      } catch (err) {
        console.error("Save failed:", err);
        alert("⛔ 저장에 실패했습니다!\n\n혹시 원본 엑셀 파일이 **다른 프로그램(Excel 등)에서 열려 있나요?**\n파일을 닫고 다시 시도해 주세요.\n\n(확인을 누르면 '새 파일'로 다운로드하여 데이터를 보호합니다.)");
        handleDownload();
      }
    }
    // CASE B: No handle -> Download
    else {
      handleDownload();
    }
  };

  const handleDownload = () => {
    // Re-use logic for simple download
    const buffer = getWorkbookBuffer(); // reuse buffer logic
    // Create Blob and trigger download
    const blob = new Blob([buffer], { type: "application/octet-stream" });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName ? fileName.replace(/\.xlsx?$/i, "") + "_exported.xlsx" : "workbook.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  };


  // --- Core Logic Features ---

  const handleRemoveDuplicates = () => {
    if (excelData.length === 0) return;

    // Use JSON stringify to check for distinct rows
    // Note: This is simple and effective for basic data types.
    // For handling dates/objects strictly, deeper comparison might be needed but stringify is usually fine for Excel data.
    const uniqueMap = new Map();
    const uniqueData: any[][] = [];
    let duplicatesCount = 0;

    excelData.forEach(row => {
      const signature = JSON.stringify(row);
      if (!uniqueMap.has(signature)) {
        uniqueMap.set(signature, true);
        uniqueData.push(row);
      } else {
        duplicatesCount++;
      }
    });

    if (duplicatesCount === 0) {
      alert("중복된 행이 하나도 없어요! 👍");
    } else {
      updateData(uniqueData);
      alert(`중복된 행 ${duplicatesCount}개를 찾아 삭제했어요! 🗑️`);
    }
  };

  const handleTrimSpaces = (option?: { allSpaces: boolean }) => {
    if (excelData.length === 0) return;

    let trimCount = 0;
    const newData = excelData.map(row => {
      if (!row) return row;
      return row.map(cell => {
        if (typeof cell === 'string') {
          let trimmed = cell;
          if (option?.allSpaces) {
            // Remove ALL spaces
            trimmed = cell.replace(/\s+/g, '');
          } else {
            // Check if user has explicit object wrapper, if so, we might need to update .v
            // (Assuming simple strings for now or stringifying)
            trimmed = cell.trim();
          }

          if (trimmed !== cell) trimCount++;
          return trimmed;
        }
        return cell;
      });
    });

    if (trimCount === 0) {
      alert("공백이 없어서 청소할 게 없네요! ✨");
    } else {
      updateData(newData);
      alert(`총 ${trimCount}개의 칸에서 ${option?.allSpaces ? '모든' : '불필요한'} 공백을 닦아냈어요! 🧼`);
    }
  };

  const handleAddCommas = () => {
    if (excelData.length === 0) return;

    let formatCount = 0;
    const newData = excelData.map(row => {
      if (!row) return row;
      return row.map(cell => {
        // Check if cell is a number or a string looking like a number
        if (cell !== null && cell !== undefined && cell !== '') {
          const numValue = Number(cell);
          if (!isNaN(numValue) && typeof cell === 'number') {
            // It's a raw number, format it
            formatCount++;
            return numValue.toLocaleString();
          }
          // Optional: Attempt to format string numbers? Let's stick to safe numbers first or numeric strings.
          // If user wants string numbers formatted, we can double check.
          // For now, let's format actual numbers. 
          // If the spreadsheet loaded numbers as strings (e.g. "1234"), we might want to check.
        }
        return cell;
      });
    });

    if (formatCount === 0) {
      alert("숫자가 없어서 쉼표를 찍을 게 없어요.");
    } else {
      updateData(newData);
      alert(`${formatCount}개의 숫자에 쉼표(,)를 예쁘게 찍었어요! ✍️`);
    }
  };

  const handleHeaderStyle = () => {
    if (excelData.length === 0) return;

    // This logic is slightly different: we don't change DATA, we change STYLING metadata.
    // However, our current simple model stores data in state `excelData` which is mostly values.
    // If we want to persist styles, we might need a more complex state or apply it on export.
    // BUT, `xlsx-js-style` allows cell objects.
    // Let's wrap the first row's values into objects with .s (style) property if they aren't already.

    const newData = [...excelData];
    const headerRow = newData[0];

    if (!headerRow) return;

    newData[0] = headerRow.map((cell: any) => {
      // If cell is primitive, make it object.
      let cellObj = (typeof cell === 'object' && cell !== null) ? { ...cell } : { v: cell, t: 's' };

      // Apply Style
      if (!cellObj.s) cellObj.s = {};
      cellObj.s.font = { bold: true, color: { rgb: "000000" } };
      cellObj.s.fill = { fgColor: { rgb: "EFEFEF" } };
      cellObj.s.alignment = { horizontal: "center", vertical: "center" };

      return cellObj;
    });

    updateData(newData);
    alert("첫 번째 줄을 헤더로 예쁘게 꾸몄어요! 🎨");
  };

  // --- Recipe Logic ---
  const handleCalculate = (recipe: { type: 'sum' | 'count' | 'avg'; filter?: { operator: string; value: number | string }; format?: boolean }) => {
    if (!selection || excelData.length === 0) {
      alert("계산할 영역을 먼저 선택해주세요!");
      return;
    }

    const { type, filter, format } = recipe;
    const { start, end } = selection;

    // Calc Range (same range logic)
    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);

    let sum = 0;
    let count = 0;
    let validCellCount = 0;

    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const cell = excelData[r]?.[c];
        let val: any = cell;

        // Unwrap object if needed
        if (cell && typeof cell === 'object' && 'v' in cell) {
          val = cell.v;
        }

        // Skip empty/null for calc
        if (val === null || val === undefined || val === '') continue;

        const numVal = Number(val);
        const isNum = !isNaN(numVal);

        // Filter Check
        let pass = true;
        if (filter && isNum) {
          // Apply simple operators
          if (filter.operator === '>' && !(numVal > Number(filter.value))) pass = false;
          if (filter.operator === '<' && !(numVal < Number(filter.value))) pass = false;
          if (filter.operator === '>=' && !(numVal >= Number(filter.value))) pass = false;
          if (filter.operator === '<=' && !(numVal <= Number(filter.value))) pass = false;
          if (filter.operator === '=' && !(numVal === Number(filter.value))) pass = false;
        }

        if (pass) {
          if (type === 'sum' || type === 'avg') {
            if (isNum) {
              sum += numVal;
              count++;
            }
          } else if (type === 'count') {
            count++;
          }
        }
      }
    }

    let resultMsg = "";
    if (type === 'sum') {
      resultMsg = `합계: ${format ? sum.toLocaleString() : sum}`;
    } else if (type === 'count') {
      resultMsg = `개수: ${format ? count.toLocaleString() : count}개`;
    } else if (type === 'avg') {
      const avg = count > 0 ? sum / count : 0;
      resultMsg = `평균: ${format ? avg.toLocaleString(undefined, { maximumFractionDigits: 2 }) : avg}`;
    }

    alert(`계산 결과 🧮\n${resultMsg}`);
  };

  // --- Text Logic ---
  const handleTextAction = (recipe: { type: 'join' | 'split' | 'extract'; option?: { delimiter?: string; count?: number; mode?: 'left' | 'right' }; keepOriginal?: boolean }) => {
    if (!selection || excelData.length === 0) {
      alert("적용할 영역을 먼저 선택해주세요!");
      return;
    }

    const { type, option, keepOriginal } = recipe;
    const { start, end } = selection;

    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);

    const newData = excelData.map(row => [...(row || [])]);

    // Make sure we have enough space for result if keepOriginal is true?
    // Simplified: "Keep Original" -> Result goes to next available column (c2 + 1) for each row? 
    // Or just "Don't overwrite source, make new column"? 
    // Let's implement: "Result overwrites source" (default) VS "Result appears in next column" (keepOriginal)

    let modifiedCount = 0;

    if (type === 'join') {
      const delimiter = option?.delimiter ?? "";
      // Join Row by Row? Or all cells? Usually Row by Row is safer for massive list. 
      // Let's do: For each row in range, join the selected cols.
      for (let r = r1; r <= r2; r++) {
        if (!newData[r]) newData[r] = [];

        const parts = [];
        for (let c = c1; c <= c2; c++) {
          const cell = newData[r][c];
          let val = cell;
          if (cell && typeof cell === 'object' && 'v' in cell) val = cell.v;
          if (val !== null && val !== undefined) parts.push(String(val));
        }

        const joined = parts.join(delimiter);

        if (keepOriginal) {
          // Write to c2 + 1
          newData[r][c2 + 1] = joined;
        } else {
          // Overwrite first cell, clear others? Or just overwrite first and leave rest?
          // Standard behavior: Merge cells usually keeps top-left.
          newData[r][c1] = joined;
          for (let c = c1 + 1; c <= c2; c++) newData[r][c] = null; // Clear others in merge
        }
        modifiedCount++;
      }
    }
    else if (type === 'split') {
      const delimiter = option?.delimiter ?? ",";
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          // Concept: Split "A,B" -> "A" in c, "B" in c+1
          // If KeepOriginal -> "A,B" in c, "A" in c+1, "B" in c+2
          const cell = newData[r]?.[c];
          let val = cell;
          if (cell && typeof cell === 'object' && 'v' in cell) val = cell.v;

          if (typeof val === 'string' && val.includes(delimiter)) {
            const parts = val.split(delimiter);
            const targetStartCol = keepOriginal ? c + 1 : c;

            parts.forEach((part, idx) => {
              newData[r][targetStartCol + idx] = part.trim();
            });
            modifiedCount++;
          }
        }
      }
    }
    else if (type === 'extract') {
      const count = option?.count ?? 1;
      const mode = option?.mode ?? 'left';

      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          const cell = newData[r]?.[c];
          let val = cell;
          if (cell && typeof cell === 'object' && 'v' in cell) val = cell.v;

          if (typeof val === 'string') {
            let extracted = "";
            if (mode === 'left') extracted = val.substring(0, count);
            else extracted = val.substring(val.length - count);

            if (keepOriginal) {
              newData[r][c + 1] = extracted;
            } else {
              newData[r][c] = extracted;
            }
            modifiedCount++;
          }
        }
      }
    }

    updateData(newData);
    alert(`${modifiedCount}개의 행에 텍스트 마법을 부렸어요! 🧙‍♂️`);
  };

  // --- Style Logic ---
  const handleHighlight = (option?: { operator: string; value: string; color: string }) => {
    if (!selection) {
      alert("강조할 영역을 먼저 선택해주세요! 🎨");
      return;
    }

    // Color mapping
    const colorMap: Record<string, any> = {
      yellow: { rgb: "FFFF00" },
      red: { rgb: "FFCDD2" }, // Light Red
      green: { rgb: "C8E6C9" }, // Light Green
      blue: { rgb: "BBDEFB" }   // Light Blue
    };

    const targetColor = colorMap[option?.color || 'yellow'] || colorMap['yellow'];
    const operator = option?.operator || '>';
    const targetVal = option?.value || '';

    // Apply to Copy of Data
    // NOTE: Since ExcelData state stores values, and we want to change STYLES,
    // we must convert affected cells to objects { v: val, s: style } if they aren't already.
    // However, our `updateData` and simple render might only check `v`.
    // We need to ensure `SheetPreview` renders styles or at least we store them.
    // SheetPreview renders `cell.v` or `cell`.

    const newData = excelData.map(row => [...(row || [])]);
    const { start, end } = selection;
    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);

    let count = 0;

    for (let r = r1; r <= r2; r++) {
      // Safety: Ensure row exists
      if (!newData[r]) newData[r] = [];

      for (let c = c1; c <= c2; c++) {
        const cell = newData[r][c];
        let val = cell;

        // Extract plain value
        if (cell && typeof cell === 'object' && 'v' in cell) val = cell.v;

        // Skip empty?
        if (val === null || val === undefined || val === '') continue;

        const numVal = Number(val);
        const isNum = !isNaN(numVal) && typeof val !== 'boolean';
        // Note: Number(true) is 1, so stricter check needed if we care about types, 
        // but Excel often treats booleans as distinct. Let's assume standard lax Check or strict?
        // Let's use `!isNaN(numVal)` but also `val !== ''` (handled above).

        const limit = Number(targetVal);
        const isLimitNum = !isNaN(limit) && targetVal !== '';

        // Check Logic
        let match = false;

        if (operator === 'contains') {
          if (String(val).includes(targetVal)) match = true;
        } else if (isNum && isLimitNum) {
          if (operator === '>' && numVal > limit) match = true;
          if (operator === '<' && numVal < limit) match = true;
          if (operator === '=' && numVal === limit) match = true;
        }
        // String equality fallback for '='
        else if (operator === '=' && String(val) === targetVal) {
          match = true;
        }

        if (match) {
          // Robust Style Merging
          // 1. Normalize to Object
          let cellObj: any = (typeof cell === 'object' && cell !== null) ? { ...cell } : { v: cell, t: isNum ? 'n' : 's' };

          // 2. Ensure 's' (Style) object exists
          if (!cellObj.s) cellObj.s = {};
          if (!cellObj.s.fill) cellObj.s.fill = {};

          // 3. Apply New Style (Merge)
          cellObj.s.fill.fgColor = targetColor;

          newData[r][c] = cellObj;
          count++;
        }
      }
    }

    if (count > 0) {
      updateData(newData);
      alert(`조건에 맞는 ${count}개의 칸을 색칠했어요! 🖍️`);
    } else {
      alert("조건에 맞는 칸이 없네요. 😅");
    }
  };

  const handleAnalyze = (recipe: { type: string; option?: { label?: boolean } }) => {
    if (!selection) {
      alert("분석할 영역을 먼저 선택해주세요! 📊");
      return;
    }

    const { type, option } = recipe;
    const { start, end } = selection;

    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);

    // Extract Data
    const useLabel = option?.label ?? false;
    const dataStartCol = useLabel ? c1 + 1 : c1;

    // Collect Labels
    let labels: string[] = [];
    if (useLabel) {
      for (let r = r1; r <= r2; r++) {
        const cell = excelData[r]?.[c1];
        let val = (cell && typeof cell === 'object' && 'v' in cell) ? cell.v : cell;
        labels.push(val ? String(val) : `Row ${r + 1}`);
      }
    } else {
      for (let r = r1; r <= r2; r++) labels.push(`Row ${r + 1}`);
    }

    // Collect Numerical Series
    const seriesData: number[][] = [];
    const seriesNames: string[] = [];

    for (let c = dataStartCol; c <= c2; c++) {
      seriesNames.push(`Series ${c - dataStartCol + 1}`);
      const colData: number[] = [];
      for (let r = r1; r <= r2; r++) {
        const cell = excelData[r]?.[c];
        let val = (cell && typeof cell === 'object' && 'v' in cell) ? cell.v : cell;
        const num = Number(val);
        colData.push(isNaN(num) ? 0 : num);
      }
      seriesData.push(colData);
    }

    // --- LOGIC DISTRIBUTOR ---

    if (type === 'stat_basic') {
      const allNumbers = seriesData.flat();
      if (allNumbers.length === 0) { alert("통계 낼 숫자가 없어요!"); return; }

      const stats = {
        count: allNumbers.length,
        sum: ss.sum(allNumbers),
        mean: ss.mean(allNumbers),
        min: ss.min(allNumbers),
        max: ss.max(allNumbers),
        median: ss.median(allNumbers),
        stdDev: ss.standardDeviation(allNumbers)
      };
      setAnalysisResult({ type: 'stat', stats });
      setIsAnalysisOpen(true);
    }
    else if (type === 'chart_bar' || type === 'chart_line') {
      const datasets = seriesData.map((data, i) => ({
        label: seriesNames[i],
        data: data,
        backgroundColor: i === 0 ? 'rgba(75, 192, 192, 0.5)' : 'rgba(255, 99, 132, 0.5)',
        borderColor: i === 0 ? 'rgba(75, 192, 192, 1)' : 'rgba(255, 99, 132, 1)',
        borderWidth: 1
      }));

      setAnalysisResult({
        type: type === 'chart_bar' ? 'bar' : 'line',
        data: { labels, datasets },
        options: { responsive: true, plugins: { legend: { position: 'top' as const }, title: { display: true, text: 'Chart Analysis' } } }
      });
      setIsAnalysisOpen(true);
    }
    else if (type === 'chart_scatter') {
      if (seriesData.length < 2) {
        alert("산점도/회귀분석을 하려면 최소 2개의 숫자 열이 필요해요! (X, Y)");
        return;
      }
      const xData = seriesData[0];
      const yData = seriesData[1];
      const scatterData = xData.map((x, i) => ({ x, y: yData[i] }));

      // Regression
      const regressionPoints = scatterData.map(p => [p.x, p.y]);
      const regression = ss.linearRegression(regressionPoints);
      const regressionLine = ss.linearRegressionLine(regression);
      const rSquared = ss.rSquared(regressionPoints, regressionLine);

      const minX = ss.min(xData);
      const maxX = ss.max(xData);
      const trendData = [
        { x: minX, y: regressionLine(minX) },
        { x: maxX, y: regressionLine(maxX) }
      ];

      setAnalysisResult({
        type: 'scatter',
        data: {
          datasets: [
            {
              label: 'Data',
              data: scatterData,
              backgroundColor: 'rgba(53, 162, 235, 0.5)',
            },
            {
              type: 'line',
              label: `Regression (R²=${rSquared.toFixed(2)})`,
              data: trendData,
              borderColor: 'rgba(255, 99, 132, 1)',
              borderWidth: 2,
              pointRadius: 0
            }
          ]
        },
        options: { scales: { x: { type: 'linear', position: 'bottom' } } }
      });
      setIsAnalysisOpen(true);
    }
  };

  // --- Data Logic (Sort, Filter, Replace) ---
  const handleSort = (order: 'asc' | 'desc', option?: { header: boolean }) => {
    if (!selection) {
      alert("정렬할 기준 열(Column)을 선택해주세요! 📊");
      return;
    }
    const colIdx = selection.start.c;
    const hasHeader = option?.header ?? true;

    // Deep copy current sheet data
    let sortedData = [...excelData];
    let headerRow = null;

    if (hasHeader && sortedData.length > 0) {
      headerRow = sortedData.shift(); // Remove header temporarily to protect it
    }

    sortedData.sort((a, b) => {
      const valA = a[colIdx];
      const valB = b[colIdx];

      // Extract values if object
      const vA = (valA && typeof valA === 'object' && 'v' in valA) ? valA.v : valA;
      const vB = (valB && typeof valB === 'object' && 'v' in valB) ? valB.v : valB;

      if (vA === vB) return 0;
      if (vA === null || vA === undefined) return 1; // Empty last
      if (vB === null || vB === undefined) return -1;

      // Number comparison
      if (!isNaN(Number(vA)) && !isNaN(Number(vB)) && vA !== "" && vB !== "") {
        return order === 'asc' ? Number(vA) - Number(vB) : Number(vB) - Number(vA);
      }

      // String comparison
      const strA = String(vA).toLowerCase();
      const strB = String(vB).toLowerCase();
      if (strA < strB) return order === 'asc' ? -1 : 1;
      if (strA > strB) return order === 'asc' ? 1 : -1;
      return 0;
    });

    if (hasHeader && headerRow) {
      sortedData.unshift(headerRow);
    }

    updateData(sortedData);
    // alert(`정렬이 완료되었습니다!`);
  };

  const handleFilter = (condition: string, option?: { header: boolean }) => {
    if (!selection) {
      alert("필터링할 기준 열을 선택해주세요! 🌪️");
      return;
    }
    if (!condition) {
      alert("조건을 입력해주세요 (예: >50, 서울)");
      return;
    }

    const colIdx = selection.start.c;
    const hasHeader = option?.header ?? true;
    let filteredData = [...excelData];
    let headerRow = null;

    if (hasHeader && filteredData.length > 0) {
      headerRow = filteredData.shift();
    }

    // Filter Logic
    const matchingRows = filteredData.filter(row => {
      const cell = row[colIdx];
      const val = (cell && typeof cell === 'object' && 'v' in cell) ? cell.v : cell;

      // Empty check
      if (val === null || val === undefined || val === "") return false;

      const strVal = String(val);

      // Numeric Condition (>50, <=100)
      if (condition.match(/^[<>=]+/)) {
        // Simple parser
        let operator = condition.match(/^[<>=]+/)?.[0];
        let numStr = condition.substring(operator?.length || 0);

        const conditionNum = Number(numStr);
        const cellNum = Number(val);

        if (isNaN(conditionNum) || isNaN(cellNum)) return false; // Fail if not numbers

        if (operator === '>') return cellNum > conditionNum;
        if (operator === '<') return cellNum < conditionNum;
        if (operator === '>=') return cellNum >= conditionNum;
        if (operator === '<=') return cellNum <= conditionNum;
        if (operator === '=') return cellNum === conditionNum;
      }

      // String Condition
      return strVal.toLowerCase().includes(condition.toLowerCase());
    });

    if (hasHeader && headerRow) {
      matchingRows.unshift(headerRow);
    }

    updateData(matchingRows);
    alert(`조건에 맞는 ${matchingRows.length - (hasHeader ? 1 : 0)}개의 행만 남겼어요! 🗑️`);
  };

  const handleReplace = (find: string, replace: string) => {
    if (!find) {
      alert("찾을 내용을 입력해주세요!");
      return;
    }
    let newData = [...excelData];
    let count = 0;

    // Apply to Selection or Whole Sheet?
    // User requested "Selection Based" but fall back to whole if needed.
    // Let's iterate whole data for "Replace All" behavior if no range provided, 
    // but here we usually have a selection or focused cell.
    // If range selection exists (and spans > 1 cell), replace inside it.
    // Else, replace ALL in sheet.

    let r1 = 0, c1 = 0, r2 = excelData.length - 1, c2 = excelData[0]?.length - 1;

    if (selection && (selection.start.r !== selection.end?.r || selection.start.c !== selection.end?.c)) {
      // Range Selected
      r1 = Math.min(selection.start.r, selection.end.r);
      r2 = Math.max(selection.end.r, selection.start.r);
      c1 = Math.min(selection.start.c, selection.end.c);
      c2 = Math.max(selection.end.c, selection.start.c);
    } else {
      // Single cell selected or no selection -> Ask user? Or Replace All?
      // User request: "If selection exists, inside it. If not, whole sheet text replacement."
      // I'll stick to full sheet if single cell selected is treated as "No Range".
      // But be careful. Replace All is powerful.
      // Alert user? No, just do it.
    }

    newData = newData.map((row, r) => {
      if (r < r1 || r > r2) return row; // Skip rows outside range
      return row.map((cell, c) => {
        if (c < c1 || c > c2) return cell; // Skip cols outside range

        let val = cell;
        if (cell && typeof cell === 'object' && 'v' in cell) val = cell.v;

        if (val !== null && val !== undefined) {
          const strVal = String(val);
          if (strVal.includes(find)) {
            const newValStr = strVal.replaceAll(find, replace);
            count++;

            // Smart Type Detection (Prevent data pollution)
            let finalVal: string | number = newValStr;
            let type = 's';

            // If the result looks like a number (and isn't empty), save as Number
            if (newValStr.trim() !== "" && !isNaN(Number(newValStr))) {
              finalVal = Number(newValStr);
              type = 'n';
            }

            // If original was object, preserve properties but update value & type
            if (cell && typeof cell === 'object' && 'v' in cell) {
              return { ...cell, v: finalVal, t: type };
            }
            return finalVal;
          }
        } return cell;
      });
    });

    updateData(newData);
    alert(`${count}개의 '${find}'를 '${replace}'로 바꿨어요! 🔄`);
  };

  // --- Context Menu Handlers ---
  const handleInsertRow = (rowIndex: number, position: 'above' | 'below') => {
    let newData = [...excelData];
    const insertIndex = position === 'above' ? rowIndex : rowIndex + 1;
    newData.splice(insertIndex, 0, []); // Insert empty row
    updateData(newData);
  };

  const handleDeleteRow = (rowIndex: number) => {
    if (confirm("정말 이 행을 삭제하시겠습니까?")) {
      let newData = [...excelData];
      newData.splice(rowIndex, 1);
      updateData(newData);
    }
  };

  const handleInsertCol = (colIndex: number, position: 'left' | 'right') => {
    let newData = excelData.map(row => [...(row || [])]);
    const insertIndex = position === 'left' ? colIndex : colIndex + 1;

    newData.forEach(row => {
      row.splice(insertIndex, 0, null); // Insert null/empty cell
    });
    updateData(newData);
  };

  const handleDeleteCol = (colIndex: number) => {
    if (confirm("정말 이 열을 삭제하시겠습니까?")) {
      let newData = excelData.map(row => [...(row || [])]);
      newData.forEach(row => {
        row.splice(colIndex, 1);
      });
      updateData(newData);
    }
  };

  const handleClearContent = (rowIndex: number, colIndex: number) => {
    let newData = [...excelData];
    if (newData[rowIndex]) {
      newData[rowIndex][colIndex] = null;
    }
    updateData(newData);
  };

  // --- Math Logic ---
  // --- Math Logic ---
  // --- Math Logic ---
  // --- Math Logic ---
  // Formula Calculation Helper
  const calculateFormula = (formula: any, context: { excelData: any[][], rowIndex: number, colIndex: number }) => {
    const { type, startC, endC, startR, endR, operator, value } = formula;
    const { excelData, rowIndex, colIndex } = context;
    const newData = excelData; // Read-only access to current data

    let res = 0;
    let count = 0;

    if (type === 'row_sum') {
      let first = true;
      for (let c = startC; c <= endC; c++) {
        const cell = newData[rowIndex]?.[c];
        // Robust Extraction
        const val = (cell && typeof cell === 'object') ? cell.v : cell;

        // Robust Parsing (Handle string numbers, 0, empty string)
        const num = parseFloat(String(val));

        if (!isNaN(num)) {
          if (first) { res = num; first = false; }
          else {
            if (operator === '+') res += num;
            else if (operator === '-') res -= num;
            else if (operator === '*') res *= num;
            else if (operator === '/') res = num !== 0 ? res / num : 0;
          }
          count++;
        }
      }
      if (first) return 0;
      return res;
    }
    else if (type === 'col_sum') {
      let first = true;
      for (let r = startR; r <= endR; r++) {
        const cell = newData[r]?.[colIndex];
        const val = (cell && typeof cell === 'object') ? cell.v : cell;
        const num = parseFloat(String(val));

        if (!isNaN(num)) {
          if (first) { res = num; first = false; }
          else {
            if (operator === '+') res += num;
            else if (operator === '-') res -= num;
            else if (operator === '*') res *= num;
            else if (operator === '/') res = num !== 0 ? res / num : 0;
          }
          count++;
        }
      }
      if (first) return 0;
      return res;
    }
    else if (type === 'range_agg') {
      // Range Aggregation (Unified Builder)
      // Usage: Sum(A1:C2) -> Result.
      let values: number[] = [];
      for (let r = startR; r <= endR; r++) {
        for (let c = startC; c <= endC; c++) {
          const cell = newData[r]?.[c];
          const val = (cell && typeof cell === 'object') ? cell.v : cell;
          const num = parseFloat(String(val));
          if (!isNaN(num)) values.push(num);
        }
      }

      if (operator === '+' || operator === 'sum') {
        res = values.reduce((a, b) => a + b, 0);
      } else if (operator === 'avg' || operator === 'average') {
        res = values.reduce((a, b) => a + b, 0) / (values.length || 1);
      } else if (operator === 'count') {
        res = values.length;
      } else if (operator === 'max') {
        res = Math.max(...values);
      } else if (operator === 'min') {
        res = Math.min(...values);
      } else if (operator === '*') { // optional
        res = values.reduce((a, b) => a * b, 1);
      }
      return res;
    }
    else if (type === 'const_op') {
      // Constant Operation (e.g., A1 + 10)
      // The dragging logic is trickier here. "A1" is relative?
      // Let's assume math_col_const applies to the cell *directly left* or similar?
      // Current implementation of 'math_col_const' applies to the cell ITSELF (A1 = A1 + 10).
      // That is a direct mutation, not a formula referencing another cell.
      // So we likely don't store a formula for "Add 10 to self" in this context unless we change it to "B1 = A1 + 10".
      // The current `handleMathColConst` does: Cell = Cell + Const.
      // We will SKIP formula generation for Mutable actions for now, or handle explicitly if requested.
      // User request is specifically about "Sum/Calc Result" (A+B -> C).
      return (formula.currentValue || 0); // Fallback
    }
    else if (type === 'logic_if') {
      const { targetR, targetC, operator, compareVal, trueVal, falseVal } = formula;
      const cell = newData[targetR]?.[targetC];
      const sourceVal = (cell && typeof cell === 'object' && 'v' in cell) ? cell.v : (cell ?? "");

      const sNum = parseFloat(String(sourceVal));
      const vNum = parseFloat(compareVal);
      const isNum = !isNaN(sNum) && !isNaN(vNum);

      let conditionMet = false;
      if (isNum) {
        if (operator === '>') conditionMet = sNum > vNum;
        if (operator === '>=') conditionMet = sNum >= vNum;
        if (operator === '<') conditionMet = sNum < vNum;
        if (operator === '<=') conditionMet = sNum <= vNum;
        if (operator === '=') conditionMet = sNum === vNum;
        if (operator === '!=') conditionMet = sNum !== vNum;
      } else {
        if (operator === '=') conditionMet = String(sourceVal) == compareVal;
        if (operator === '!=') conditionMet = String(sourceVal) != compareVal;
      }
      return conditionMet ? trueVal : falseVal;
    }
    else if (type === 'logic_vlookup') {
      const { targetR, targetC, rangeStr, colIndex } = formula;
      const cell = newData[targetR]?.[targetC];
      const lookupVal = (cell && typeof cell === 'object' && 'v' in cell) ? cell.v : (cell ?? "");

      // Parse Range A1:C5
      const rMatch = rangeStr.match(/([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)/);
      if (!rMatch) return "#N/A";

      const c1Str = rMatch[1]; const r1 = parseInt(rMatch[2]) - 1;
      const c2Str = rMatch[3]; const r2 = parseInt(rMatch[4]) - 1;

      let c1 = 0; for (let i = 0; i < c1Str.length; i++) c1 = c1 * 26 + (c1Str.charCodeAt(i) - 64); c1 -= 1;

      // Search
      for (let r = r1; r <= r2; r++) {
        const keyCell = newData[r]?.[c1];
        const key = (keyCell && typeof keyCell === 'object' && 'v' in keyCell) ? keyCell.v : (keyCell ?? "");
        if (String(key) == String(lookupVal)) {
          const resCell = newData[r]?.[c1 + colIndex - 1];
          return (resCell && typeof resCell === 'object' && 'v' in resCell) ? resCell.v : (resCell ?? "");
        }
      }
      return "#N/A";
    }

    return 0;
  };

  // Helper: Evaluator for String Formulas (IF, VLOOKUP)
  const evaluateFormulaString = (formula: string, context: { excelData: any[][], rowIndex: number, colIndex: number }) => {
    const { excelData } = context;
    if (!formula || !formula.startsWith('=')) return "#ERROR";

    // Helper: cell ref to value (A1 -> val)
    const getVal = (ref: string) => {
      // Check if ref is "A1" format
      const match = ref.match(/^([A-Z]+)([0-9]+)$/);
      if (match) {
        const cStr = match[1];
        const rVal = parseInt(match[2]) - 1;
        let cVal = 0;
        for (let i = 0; i < cStr.length; i++) cVal = cVal * 26 + (cStr.charCodeAt(i) - 64);
        cVal -= 1;

        const cell = excelData[rVal]?.[cVal];
        const v = (cell && typeof cell === 'object' && 'v' in cell) ? cell.v : (cell ?? "");
        return isNaN(Number(v)) ? v : Number(v);
      }
      // Check if it's string "abc" w/ quotes
      if (ref.startsWith('"') && ref.endsWith('"')) return ref.slice(1, -1);
      // Check number
      if (!isNaN(Number(ref))) return Number(ref);
      return ref;
    };

    if (formula.startsWith('=IF(')) {
      // Parse: =IF(cond, true, false)
      // Simplistic split by comma (doesn't handle nested parens well but OK for hotfix)
      const inner = formula.slice(4, -1);
      const parts = inner.split(',').map(s => s.trim());
      if (parts.length < 3) return "#N/A";

      // Condition Parse: A1>10 or A1="Y"
      const cond = parts[0];
      const trueValRaw = parts[1];
      const falseValRaw = parts.slice(2).join(','); // rest

      // Split condition by operators
      const ops = ['>=', '<=', '!=', '=', '>', '<'];
      let op = '';
      for (let o of ops) { if (cond.includes(o)) { op = o; break; } }

      if (!op) return "#ERR";
      const [leftRaw, rightRaw] = cond.split(op).map(s => s.trim());
      const left = getVal(leftRaw);
      const right = getVal(rightRaw);

      let res = false;
      if (op === '>') res = left > right;
      if (op === '>=') res = left >= right;
      if (op === '<') res = left < right;
      if (op === '<=') res = left <= right;
      if (op === '=') res = String(left) == String(right);
      if (op === '!=') res = String(left) != String(right);

      return res ? getVal(trueValRaw) : getVal(falseValRaw);
    }

    if (formula.startsWith('=VLOOKUP(')) {
      // =VLOOKUP(lookup, range, col, 0)
      const inner = formula.slice(9, -1);
      const parts = inner.split(',').map(s => s.trim());
      if (parts.length < 3) return "#N/A";

      const lookupVal = getVal(parts[0]);
      const rangeStr = parts[1]; // A1:C10
      const colIdx = parseInt(parts[2]);

      // Parse Range
      const rMatch = rangeStr.match(/([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)/);
      if (!rMatch) return "#REF!";

      const c1Str = rMatch[1]; const r1 = parseInt(rMatch[2]) - 1;
      const c2Str = rMatch[3]; const r2 = parseInt(rMatch[4]) - 1;

      let c1 = 0; for (let i = 0; i < c1Str.length; i++) c1 = c1 * 26 + (c1Str.charCodeAt(i) - 64); c1 -= 1;

      // Search
      for (let r = r1; r <= r2; r++) {
        const keyCell = excelData[r]?.[c1];
        const key = (keyCell && typeof keyCell === 'object' && 'v' in keyCell) ? keyCell.v : (keyCell ?? "");
        if (String(key) == String(lookupVal)) {
          const resCell = excelData[r]?.[c1 + colIdx - 1];
          const res = (resCell && typeof resCell === 'object' && 'v' in resCell) ? resCell.v : (resCell ?? "");
          return res;
        }
      }
      return "#N/A";
    }

    return "#NAME?";
  };


  const handleMathColCol = (operator: string, resultMode: 'overwrite' | 'new' | 'pick_execute' = 'overwrite', targetStart?: { r: number, c: number }) => {
    if (!selection) { alert("영역을 선택해주세요!"); return; }
    const { start, end } = selection;
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);
    if (c2 - c1 < 1) {
      alert("열끼리 계산하려면 최소 2개의 열(Column)을 선택해야 해요! (A+B+...)");
      return;
    }

    let newData = excelData.map(row => [...(row || [])]);
    let count = 0;
    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);

    for (let r = r1; r <= r2; r++) {
      // Structure Formula Logic
      // We are operating on columns c1 to c2 at row r.
      // This is a 'row_sum' type (Horizontal input, Vertical output usually, or Single Cell output)
      // Actually it's "Row Operation": inputs are in the same row.

      const formula = {
        type: 'row_sum',
        startC: c1,
        endC: c2,
        operator: operator
      };

      // Calculate Initial Value
      const res = calculateFormula(formula, { excelData: newData, rowIndex: r, colIndex: -1 }); // colIndex irrelevant for row_sum

      // Result Location Logic
      let targetR = r;
      let targetC = resultMode === 'new' ? c2 + 1 : c2;
      if (resultMode === 'overwrite') targetC = c2;

      if (resultMode === 'pick_execute' && targetStart) {
        targetR = targetStart.r + (r - r1);
        targetC = targetStart.c;
      }

      // Ensure Matrix Expansion
      if (!newData[targetR]) {
        for (let i = newData.length; i <= targetR; i++) newData[i] = [];
      }
      if (!newData[targetR]) newData[targetR] = [];

      const targetCell = newData[targetR][targetC];
      // Save Value AND Formula
      // Preserve Style if exists
      const cellStyle = (targetCell && typeof targetCell === 'object' && targetCell.s) ? targetCell.s : undefined;

      newData[targetR][targetC] = {
        v: res,
        t: 'n',
        formula: formula,
        s: cellStyle
      };
      count++;
    }
    updateData(newData);
    alert(`${count}개의 행을 계산했어요! 🧮 (수식 기억 완료)`);
  };

  const handleMathRowRow = (operator: string, resultMode: 'overwrite' | 'new' | 'pick_execute' = 'overwrite', targetStart?: { r: number, c: number }) => {
    if (!selection) { alert("영역을 선택해주세요!"); return; }
    const { start, end } = selection;
    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);

    if (r2 - r1 < 1) {
      alert("행끼리 계산하려면 최소 2개의 행(Row)을 선택해야 해요! (위+아래+...)");
      return;
    }

    let newData = excelData.map(row => [...(row || [])]);
    let count = 0;
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);

    // Target Row logic (Base)
    let targetR = resultMode === 'new' ? (r2 + 1) : r2; // Overwrite = Last Row (r2)

    // Pick Logic overrides
    if (resultMode === 'pick_execute' && targetStart) {
      targetR = targetStart.r;
    } else {
      // "New" Mode: Insert Row
      if (resultMode === 'new') {
        if (!newData[targetR]) newData[targetR] = [];
        else newData.splice(targetR, 0, []);
      }
    }

    if (!newData[targetR]) {
      for (let i = newData.length; i <= targetR; i++) newData[i] = [];
    }

    for (let c = c1; c <= c2; c++) {
      // Formula Construction for Column Calculation
      // Inputs are rows r1..r2 at Column c. 
      const formula = {
        type: 'col_sum',
        startR: r1,
        endR: r2,
        operator: operator
      };

      const res = calculateFormula(formula, { excelData: newData, rowIndex: -1, colIndex: c });

      let targetC = c;
      if (resultMode === 'pick_execute' && targetStart) {
        targetC = targetStart.c + (c - c1);
      }

      if (!newData[targetR]) newData[targetR] = [];
      const targetRow = newData[targetR];

      const cellStyle = (targetRow[targetC] && typeof targetRow[targetC] === 'object' && targetRow[targetC].s)
        ? targetRow[targetC].s
        : undefined;

      targetRow[targetC] = {
        v: res,
        t: 'n',
        formula: formula,
        s: cellStyle
      };
      count++;
    }
    updateData(newData);
    alert(`${count}개의 열을 계산했어요! 🔽 (수식 기억 완료)`);
  };

  const handleMathColConst = (operator: string, value: string, resultMode: 'overwrite' | 'new' | 'pick_execute' = 'overwrite', targetStart?: { r: number, c: number }) => {
    if (!selection) { alert("영역을 선택해주세요!"); return; }
    const numVal = Number(value);
    if (isNaN(numVal) || value === '') { alert("올바른 숫자를 입력해주세요!"); return; }

    let newData = excelData.map(row => [...(row || [])]);
    let count = 0;
    const { start, end } = selection;
    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);
    const width = c2 - c1 + 1;

    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const cell = newData[r]?.[c];
        let val = (cell && typeof cell === 'object' && 'v' in cell) ? cell.v : cell;
        const currentNum = Number(val);

        if (!isNaN(currentNum) && val !== "" && val !== null) {
          let res = 0;
          if (operator === '+') res = currentNum + numVal;
          if (operator === '-') res = currentNum - numVal;
          if (operator === '*') res = currentNum * numVal;
          if (operator === '/') res = numVal !== 0 ? currentNum / numVal : 0;

          let targetR = r;
          let targetC = resultMode === 'new' ? (c + width) : c;

          if (resultMode === 'pick_execute' && targetStart) {
            targetR = targetStart.r + (r - r1);
            targetC = targetStart.c + (c - c1);
          }

          // Ensure expansion
          if (!newData[targetR]) {
            for (let i = newData.length; i <= targetR; i++) newData[i] = [];
          }
          if (!newData[targetR]) newData[targetR] = [];

          const targetCell = newData[targetR][targetC];
          if (targetCell && typeof targetCell === 'object') {
            newData[targetR][targetC] = { ...targetCell, v: res, t: 'n' };
          } else {
            newData[targetR][targetC] = res;
          }
          count++;
        }
      }
    }
    updateData(newData);
    alert(`${count}개의 칸을 계산했어요! ✨`);
  };

  const handleAutofill = (source: { start: { r: number; c: number }; end: { r: number; c: number } }, target: { start: { r: number; c: number }; end: { r: number; c: number } }) => {
    let newData = excelData.map(row => [...(row || [])]);

    const srcR1 = Math.min(source.start.r, source.end.r);
    const srcRMax = Math.max(source.start.r, source.end.r);
    const srcC1 = Math.min(source.start.c, source.end.c);
    const srcCMax = Math.max(source.start.c, source.end.c);

    const trgR1 = Math.min(target.start.r, target.end.r);
    const trgR2 = Math.max(target.start.r, target.end.r);
    const trgC1 = Math.min(target.start.c, target.end.c);
    const trgC2 = Math.max(target.start.c, target.end.c);

    const expandVert = trgR2 > srcRMax;
    const expandHorz = trgC2 > srcCMax;

    // Vertical Fill
    if (expandVert) {
      for (let c = trgC1; c <= trgC2; c++) {
        const colValues = [];
        const sourceFormulas: any[] = [];
        for (let r = srcR1; r <= srcRMax; r++) {
          const cell = newData[r]?.[c];
          let v = (cell && typeof cell === 'object' && 'v' in cell) ? cell.v : cell;
          colValues.push(v);

          if (cell && typeof cell === 'object' && cell.formula) {
            sourceFormulas.push(cell.formula);
          } else if (cell && typeof cell === 'object' && cell.f) {
            sourceFormulas.push(cell.f); // Support String Formulas
          } else {
            sourceFormulas.push(null);
          }
        }

        let isNumeric = colValues.length >= 2 && colValues.every(v => v !== "" && v !== null && !isNaN(Number(v)));
        let step = 0;
        if (isNumeric) {
          const nums = colValues.map(Number);
          const firstStep = nums[1] - nums[0];
          let isProgression = true;
          for (let i = 1; i < nums.length; i++) {
            if (Math.abs((nums[i] - nums[i - 1]) - firstStep) > 0.0001) isProgression = false;
          }
          if (isProgression) step = firstStep;
          else isNumeric = false;
        }

        for (let r = srcRMax + 1; r <= trgR2; r++) {
          if (!newData[r]) newData[r] = [];

          const patternIdx = (r - srcR1) % colValues.length;
          const formulaInfo = sourceFormulas[patternIdx];

          let newVal;
          let newFormula = null;

          // PRIORITY 1: Formula
          if (formulaInfo) {
            // Formula Cloning & Index Shifting (Relative Reference)
            // Formula Cloning & Index Shifting (Relative Reference)
            let shiftedFormula = JSON.parse(JSON.stringify(formulaInfo));


            // Vertical Shift: Row + (TargetRow - SourceRow)
            // But we already know 'r' is the target row.
            // Wait, we need the OFFSET from the Source Formula's context.
            // The Source Formula was defined at 'srcR1' (or patternIdx relative).
            // Actually, if we copy A1's formula to A2.
            // The offset is A2 - A1 = 1 row.
            // So we add 1 to startR/endR.

            // Calculate Shift Amount
            // patternIdx maps to a specific source row index? No, sourceFormulas is just values.
            // We need to know the 'original row' this formula came from.
            // 'srcR1 + patternIdx' is the row where formulaInfo originated.
            const originalRow = srcR1 + patternIdx;
            const rowOffset = r - originalRow;

            if (shiftedFormula.type === 'range_agg') {
              shiftedFormula.startR += rowOffset;
              shiftedFormula.endR += rowOffset;
              // Col remains same for vertical drag
            } else if (shiftedFormula.type === 'row_sum') {
              // row_sum operates at `rowIndex`, so startC/endC don't change, but context row changes.
              // calculateFormula uses 'r' (passed as rowIndex). So this is auto-handled by context!
              // BUT we still need to persist the SAME formula object type.
              // UNLESS row_sum stored 'row index' inside it? 
              // Creating 'row_sum' formula: { type, startC, endC, op }. No R inside.
              // So row_sum relies on context. It is naturally relative.
            } else if (shiftedFormula.type === 'logic_if' || shiftedFormula.type === 'logic_vlookup') {
              shiftedFormula.targetR += rowOffset;
              // shiftedFormula.targetC += 0; // Vertical drag doesn't change column ref
            }

            // Recalculate
            newVal = calculateFormula(shiftedFormula, { excelData: newData, rowIndex: r, colIndex: c });
            newFormula = shiftedFormula; // Persist SHIFTED formula
          }
          else if (typeof formulaInfo === 'string') {
            // Priority 1-B: String Formula (Logic Functions)
            // Shift References (e.g. A1 -> A2)
            const rowOffset = r - ((srcR1 + patternIdx)); // Current Row - Original Source Row

            // Helper to shift A1 string
            const shiftRef = (ref: string) => {
              const match = ref.match(/^([A-Z]+)([0-9]+)$/);
              if (!match) return ref;
              const colStr = match[1];
              const rowVal = parseInt(match[2]);

              // Keep column same (vertical fill)
              // Shift Row
              const newRow = rowVal + rowOffset;
              return `${colStr}${newRow}`;
            };

            // Regex to find refs: [A-Z]+[0-9]+
            // We need to be careful not to shift things inside quotes, but for now simple regex is okay.
            const newF = formulaInfo.replace(/([A-Z]+[0-9]+)/g, (match) => {
              // Determine if it is a cell ref. Simple heuristic: matches A1 syntax.
              // Should verify if it's a valid col? assume yes for now.
              return shiftRef(match);
            });


            newFormula = newF;
            // Calculate Value immediately using the new Evaluator
            newVal = evaluateFormulaString(newFormula, { excelData: newData, rowIndex: r, colIndex: c });
          }
          // PRIORITY 2: Numeric Series
          else if (isNumeric && step !== 0) {
            const lastVal = Number(colValues[colValues.length - 1]);
            const dist = r - srcRMax;
            newVal = lastVal + (step * dist);
          }
          // PRIORITY 3: Copy Pattern
          else {
            newVal = colValues[patternIdx];
          }

          const templateRow = srcR1 + ((r - srcR1) % colValues.length);
          const templateCell = newData[templateRow]?.[c];

          let targetCell: any = newVal;
          // Preserve Style
          if (templateCell && typeof templateCell === 'object') {
            targetCell = { ...templateCell, v: newVal };
            if (newFormula) targetCell.formula = newFormula;
          } else if (typeof newFormula === 'string') {
            targetCell = { v: newVal || "#REF!", t: 's', f: newFormula };
          } else if (newFormula) {
            targetCell = { v: newVal, t: 'n', formula: newFormula };
          }

          newData[r][c] = targetCell;
        }
      }
    }
    // Horizontal Fill
    else if (expandHorz) {
      for (let r = trgR1; r <= trgR2; r++) {
        const rowValues = [];
        const sourceFormulas: any[] = [];
        for (let c = srcC1; c <= srcCMax; c++) {
          const cell = newData[r]?.[c];
          let v = (cell && typeof cell === 'object' && 'v' in cell) ? cell.v : cell;
          rowValues.push(v);
          if (cell && typeof cell === 'object' && cell.formula) {
            sourceFormulas.push(cell.formula);
          } else if (cell && typeof cell === 'object' && cell.f) {
            sourceFormulas.push(cell.f);
          } else {
            sourceFormulas.push(null);
          }
        }

        let isNumeric = rowValues.length >= 2 && rowValues.every(v => v !== "" && v !== null && !isNaN(Number(v)));
        let step = 0;
        if (isNumeric) {
          const nums = rowValues.map(Number); // Fix: use rowValues map
          const firstStep = nums[1] - nums[0];
          // ... (Simple check: reuse logic or simplify)
          // For now, let's trust basic series copy or just copy
          step = firstStep; // Simplified series for horizontal
        }

        for (let c = srcCMax + 1; c <= trgC2; c++) {
          if (!newData[r]) newData[r] = [];

          const patternIdx = (c - srcC1) % rowValues.length;
          const formulaInfo = sourceFormulas[patternIdx];

          let newVal;
          let newFormula = null;

          if (formulaInfo) {
            // Horizontal Shift
            let shiftedFormula = JSON.parse(JSON.stringify(formulaInfo));
            const originalCol = srcC1 + patternIdx;
            const colOffset = c - originalCol;

            if (shiftedFormula.type === 'range_agg') {
              shiftedFormula.startC += colOffset;
              shiftedFormula.endC += colOffset;
            } else if (shiftedFormula.type === 'logic_if' || shiftedFormula.type === 'logic_vlookup') {
              shiftedFormula.targetC += colOffset;
            }
            // If col_sum, it uses 'colIndex' context, so it is naturally relative.

            newVal = calculateFormula(shiftedFormula, { excelData: newData, rowIndex: r, colIndex: c });
            newFormula = shiftedFormula;
          }
          else if (typeof formulaInfo === 'string') {
            // Horizontal String Shift
            const colOffset = c - ((srcC1 + patternIdx));

            const shiftRef = (ref: string) => {
              const match = ref.match(/^([A-Z]+)([0-9]+)$/);
              if (!match) return ref;
              const colStr = match[1];
              const rowVal = parseInt(match[2]);

              // Col Shift
              let colNum = 0;
              for (let i = 0; i < colStr.length; i++) colNum = colNum * 26 + (colStr.charCodeAt(i) - 64);
              colNum += colOffset;

              let label = "";
              let temp = colNum - 1;
              while (temp >= 0) {
                label = String.fromCharCode((temp % 26) + 65) + label;
                temp = Math.floor(temp / 26) - 1;
              }
              return `${label}${rowVal}`;
            };


            const newF = formulaInfo.replace(/([A-Z]+[0-9]+)/g, (match) => shiftRef(match));
            newFormula = newF;
            newVal = evaluateFormulaString(newFormula, { excelData: newData, rowIndex: r, colIndex: c });
          }
          else if (isNumeric && step !== 0) {
            const lastVal = Number(rowValues[rowValues.length - 1]);
            const dist = c - srcCMax;
            newVal = lastVal + (step * dist);
          } else {
            newVal = rowValues[patternIdx];
          }

          const templateCol = srcC1 + ((c - srcC1) % rowValues.length);
          const templateCell = newData[r]?.[templateCol];

          let targetCell: any = newVal;
          if (templateCell && typeof templateCell === 'object') {
            targetCell = { ...templateCell, v: newVal };
            if (newFormula) targetCell.formula = newFormula;
          } else if (typeof newFormula === 'string') {
            targetCell = { v: newVal || "#REF!", t: 's', f: newFormula };
          } else if (newFormula) {
            targetCell = { v: newVal, t: 'n', formula: newFormula };
          }
          newData[r][c] = targetCell;
        }
      }
    }

    updateData(newData);
  };



  // Helper: Parse A1 to {r,c}
  const parseA1Notation = (a1: string) => {
    const match = a1.match(/([A-Z]+)(\d+)/);
    if (!match) return null;
    const colStr = match[1];
    const rowStr = match[2];

    let c = 0;
    for (let i = 0; i < colStr.length; i++) {
      c = c * 26 + (colStr.charCodeAt(i) - 64);
    }
    const r = parseInt(rowStr) - 1;
    return { r, c: c - 1 };
  };

  const handleUnifiedCalculation = (payload: { start: string, end: string, connector: string, operation: string, resultMode: string }) => {
    const { start, end, connector, operation, resultMode } = payload;
    const s = parseA1Notation(start);
    // End might be empty for single cell opertaions? 
    // If connector is comma, end is required. If range, end is required.
    const e = end ? parseA1Notation(end) : s;

    if (!s || !e) {
      alert("올바른 셀 주소를 입력해주세요 (예: A1)");
      return;
    }

    // Range Identification
    const r1 = Math.min(s.r, e.r);
    const r2 = Math.max(s.r, e.r);
    const c1 = Math.min(s.c, e.c);
    const c2 = Math.max(s.c, e.c);

    // determine logic
    // Unified Builder always creates a "Formula" that aggregates the Inputs.
    // BUT, we need to know where to put the result.
    // If resultMode is 'pick', we use 'selection' (target). 
    // If 'overwrite', we put it... where?
    // Usually "Math (Rows)" put result in next column. "Math (Unified)" might need explicit target if not implied.
    // However, to keep it simple:
    // If 'overwrite', put it in the Top-Left of Selection? OR Right of Range?
    // OLD Logic: 'overwrite' meant Right of each row for row_sums.
    // NEW Logic: Unified Builder -> Standard Excel Behavior?
    // "Sum(A1:B2)" -> Result is 1 cell.
    // We need a Target Cell.

    // Let's use the current 'selection' as the Target if 'pick' mode, OR default to:
    // If Range is horizontal-ish (rows=1), put to right?
    // If Range is vertical-ish (cols=1), put to bottom?
    // If Block, put to bottom-right?

    let targetR = r2;
    let targetC = c2 + 1; // Default right

    // Logic for Autofill-ready Formula:
    // We want to generate a numeric value AND a formula object.

    // Calculate Value & Formula
    let val = 0;
    let formulaObj: any = null;

    // 1. Calculate Value (Immediate)
    let count = 0;
    let values: number[] = [];

    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const cell = excelData[r]?.[c];
        // Robust Extraction & Parsing
        const rawVal = (cell && typeof cell === 'object') ? cell.v : cell;
        const v = parseFloat(String(rawVal));
        if (!isNaN(v)) values.push(v);
      }
    }

    if (operation === 'sum' || operation === '+') {
      val = values.reduce((a, b) => a + b, 0);
      formulaObj = { type: 'range_agg', startR: r1, endR: r2, startC: c1, endC: c2, operator: '+' };
    } else if (operation === 'avg') {
      val = values.reduce((a, b) => a + b, 0) / (values.length || 1);
      formulaObj = { type: 'range_agg', startR: r1, endR: r2, startC: c1, endC: c2, operator: 'avg' };
    } else if (operation === 'count') {
      val = values.length;
      formulaObj = { type: 'range_agg', startR: r1, endR: r2, startC: c1, endC: c2, operator: 'count' };
    } else if (operation === 'max') {
      val = Math.max(...values);
      formulaObj = { type: 'range_agg', startR: r1, endR: r2, startC: c1, endC: c2, operator: 'max' };
    }

    // 2. Place Result
    // Handle Result Modes
    let newData = excelData.map(row => [...(row || [])]);

    if (resultMode === 'pick' && (payload as any).target) {
      const t = parseA1Notation((payload as any).target);
      if (t) {
        targetR = t.r;
        targetC = t.c;
      }
    }
    else if (resultMode === 'new') {
      // 'new' mode: Originally intended to insert column, but user requested NO SPLICE.
      // Just Overwrite/Append at the calculated slot (targetC).
      // Since targetC was already defaulted to `c2 + 1` (Right of Selection), we just use that.
      // No explicit action needed here, just skip the splice.
    }

    // Ensure row exists
    if (!newData[targetR]) {
      // Pack empty rows
      for (let i = newData.length; i <= targetR; i++) newData[i] = [];
    }

    // Store Result with Formula Metadata
    const cellStyle = {};
    const finalFormula = formulaObj; // Ensure this is not null if operation was valid

    newData[targetR][targetC] = {
      v: val,
      t: 'n', // Numeric Type
      formula: finalFormula, // CRITICAL: Must persist formula object for Export & Autofill
      s: cellStyle
    };

    updateData(newData);
    alert(`계산 완료! ${val} (위치: ${indexToA1Notation(targetR, targetC)})`);
  };

  const handleStyle = (payload: { type: string, value?: any }) => {
    if (!selection) { alert("영역을 선택해주세요!"); return; }

    const newData = excelData.map(row => [...(row || [])]);
    const { start, end } = selection;
    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);

    for (let r = r1; r <= r2; r++) {
      if (!newData[r]) newData[r] = [];
      for (let c = c1; c <= c2; c++) {
        let cell = newData[r][c];

        // Ensure object structure
        if (cell === null || cell === undefined || typeof cell !== 'object') {
          cell = { v: cell, t: typeof cell === 'number' ? 'n' : 's' };
        } else {
          cell = { ...cell };
        }

        // Initialize style object if missing
        if (!cell.s) cell.s = {};
        if (!cell.s.font) cell.s.font = {};
        if (!cell.s.alignment) cell.s.alignment = {};
        if (!cell.s.fill) cell.s.fill = {};

        // Apply Styles
        if (payload.type === 'bold') {
          cell.s.font.bold = !cell.s.font.bold;
        } else if (payload.type === 'italic') {
          cell.s.font.italic = !cell.s.font.italic;
        } else if (payload.type === 'underline') {
          cell.s.font.underline = !cell.s.font.underline;
        } else if (payload.type === 'align') {
          cell.s.alignment.horizontal = payload.value;
        } else if (payload.type === 'color') {
          cell.s.font.color = { rgb: payload.value };
        } else if (payload.type === 'fill') {
          cell.s.fill.fgColor = { rgb: payload.value };
        }

        newData[r][c] = cell;
      }
    }
    updateData(newData);
  };

  const handleDeleteRange = () => {
    if (!selection) return;

    // Optimistic Update
    const newData = excelData.map(row => [...(row || [])]);
    const { start, end } = selection;
    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);

    for (let r = r1; r <= r2; r++) {
      if (!newData[r]) continue;
      for (let c = c1; c <= c2; c++) {
        // Clear value and formula, preserve style
        const cell = newData[r][c];
        if (cell && typeof cell === 'object') {
          newData[r][c] = { ...cell, v: null, f: null, formula: null };
        } else {
          newData[r][c] = null;
        }
      }
    }
    updateData(newData);
  };

  // --- Logic Handlers (IF / VLOOKUP) ---
  const handleLogicIf = (payload: { operator: string, value: string, trueVal: string, falseVal: string }) => {
    if (!selection) {
      alert("셀을 선택해주세요!");
      return;
    }

    const cloneData = (d: any) => JSON.parse(JSON.stringify(d));
    const getAddr = (row: number, col: number) => {
      let label = "";
      let i = col;
      while (i >= 0) {
        label = String.fromCharCode((i % 26) + 65) + label;
        i = Math.floor(i / 26) - 1;
      }
      return `${label}${row + 1}`;
    };

    const { r, c } = selection.start;
    const newData = cloneData(excelData);

    // Result Target: Cell to the Right (c + 1)
    const targetC = c + 1;
    // Source Value: Current Cell
    const sourceCell = newData[r][c];
    const sourceVal = (sourceCell && typeof sourceCell === 'object' && 'v' in sourceCell) ? sourceCell.v : (sourceCell ?? "");

    // 1. Generate Formula String
    // Excel Formula: =IF(A1>60, "Pass", "Fail")
    // Note: We need the ADDRESS of the source cell.
    const sourceRef = getAddr(r, c);

    // Quote strings if needed for formula
    const q = (v: string) => isNaN(Number(v)) ? `"${v}"` : v;

    const formulaStr = `IF(${sourceRef}${payload.operator}${q(payload.value)}, ${q(payload.trueVal)}, ${q(payload.falseVal)})`;

    // 2. JS Evaluation (Simple approximation for immediate view)
    let result = payload.falseVal;

    // Convert for comparison
    const sNum = parseFloat(sourceVal);
    const vNum = parseFloat(payload.value);
    const isNum = !isNaN(sNum) && !isNaN(vNum);

    let conditionMet = false;
    if (isNum) {
      if (payload.operator === '>') conditionMet = sNum > vNum;
      if (payload.operator === '>=') conditionMet = sNum >= vNum;
      if (payload.operator === '<') conditionMet = sNum < vNum;
      if (payload.operator === '<=') conditionMet = sNum <= vNum;
      if (payload.operator === '=') conditionMet = sNum === vNum;
      if (payload.operator === '!=') conditionMet = sNum !== vNum;
    } else {
      // String comparison
      if (payload.operator === '=') conditionMet = String(sourceVal) == payload.value;
      if (payload.operator === '!=') conditionMet = String(sourceVal) != payload.value;
      // > < for strings? Maybe length or lexical? Let's stick to equality for strings in this simple mock.
    }

    if (conditionMet) result = payload.trueVal;

    // 3. Write to Target Cell
    if (!newData[r]) newData[r] = [];

    // Create Object Formula
    const formulaObj = {
      type: 'logic_if',
      targetR: r,
      targetC: c,
      operator: payload.operator,
      compareVal: payload.value,
      trueVal: payload.trueVal,
      falseVal: payload.falseVal
    };

    // Initial Calc
    result = calculateFormula(formulaObj, { excelData: newData, rowIndex: r, colIndex: c });

    newData[r][targetC] = {
      v: result,
      t: isNaN(Number(result)) ? 's' : 'n',
      f: formulaStr, // Keep string for display if needed, but 'formula' obj takes precedence in App
      formula: formulaObj
    };

    updateData(newData);
    // Move selection to result?
    setSelection({ start: { r, c: targetC }, end: { r, c: targetC } });
  };

  const handleLogicVlookup = (payload: { range: string, colIndex: number }) => {
    if (!selection) {
      alert("셀을 선택해주세요!");
      return;
    }
    const cloneData = (d: any) => JSON.parse(JSON.stringify(d));
    const getAddr = (row: number, col: number) => {
      let label = "";
      let i = col;
      while (i >= 0) {
        label = String.fromCharCode((i % 26) + 65) + label;
        i = Math.floor(i / 26) - 1;
      }
      return `${label}${row + 1}`;
    };

    const { r, c } = selection.start;
    const newData = cloneData(excelData);

    const sourceCell = newData[r][c];
    const sourceVal = (sourceCell && typeof sourceCell === 'object' && 'v' in sourceCell) ? sourceCell.v : (sourceCell ?? "");

    const sourceRef = getAddr(r, c);

    // Formula
    const formulaStr = `VLOOKUP(${sourceRef}, ${payload.range}, ${payload.colIndex}, 0)`;

    // JS Eval (Mock VLOOKUP)
    // 1. Parse Range (e.g. A1:C10)
    // We need to convert A1:C10 to r/c bounds. This is tricky without a parser library active in this scope.
    // But we can do a simple regex for "LetterNumber:LetterNumber"
    let result = "#N/A";

    const rangeMatch = payload.range.match(/([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)/);
    if (rangeMatch) {
      const startColStr = rangeMatch[1];
      const startRowStr = rangeMatch[2];
      const endColStr = rangeMatch[3];
      const endRowStr = rangeMatch[4];

      const colToInt = (s: string) => {
        let n = 0;
        for (let i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64;
        return n - 1;
      };

      const startC = colToInt(startColStr);
      const startR = parseInt(startRowStr) - 1;
      const endC = colToInt(endColStr);
      const endR = parseInt(endRowStr) - 1;

      // Search in range in the FIRST column of range
      for (let i = startR; i <= endR; i++) {
        if (newData[i]) {
          const checkCell = newData[i][startC];
          const checkVal = (checkCell && typeof checkCell === 'object' && 'v' in checkCell) ? checkCell.v : (checkCell ?? "");

          if (String(checkVal) == String(sourceVal)) {
            // Found! Get result at startC + colIndex - 1
            const resultC = startC + payload.colIndex - 1;
            const resCell = newData[i][resultC];
            result = (resCell && typeof resCell === 'object' && 'v' in resCell) ? resCell.v : (resCell ?? "");
            break;
          }
        }
      }
    }

    // Write Result to Right
    const targetC = c + 1;
    if (!newData[r]) newData[r] = [];

    const formulaObj = {
      type: 'logic_vlookup',
      targetR: r,
      targetC: c,
      rangeStr: payload.range,
      colIndex: payload.colIndex
    };

    result = calculateFormula(formulaObj, { excelData: newData, rowIndex: r, colIndex: c });

    newData[r][targetC] = {
      v: result,
      t: isNaN(Number(result)) ? 's' : 'n',
      f: formulaStr,
      formula: formulaObj
    };

    updateData(newData);
    setSelection({ start: { r, c: targetC }, end: { r, c: targetC } });
  };




  // --- RECIPE EXECUTION ENGINE ---

  // --- RECIPE EXECUTION ENGINE ---
  const handleRunRecipe = async (recipe: { queue: { type: string, payload: any, desc: string }[] }) => {
    if (!selection) {
      alert("레시피를 실행할 시작 셀(영역)을 선택해주세요!");
      return;
    }

    // Clone Initial State (Deep Clone for safety during loop)
    let currentData = excelData.map(row => (row ? row.map((cell: any) => (typeof cell === 'object' ? { ...cell } : cell)) : []));
    let currentSelection = { ...selection };

    // Execute Queue Sequentially
    for (const action of recipe.queue) {
      try {
        const type = action.type;
        const payload = action.payload;

        if (type === 'unified_calc') {
          // Unified Calc (Range Aggregation)
          const res = applyRangeMath(currentData, currentSelection, payload.operation || payload.payload?.operation);
          currentData = res.newData;
          if (res.newSelection) currentSelection = res.newSelection;
        }
        else if (type === 'math_row_row') {
          // ROW MATH
          const res = applyRowMath(currentData, currentSelection, payload?.option?.operator || 'sum', payload?.option?.resultMode);
          currentData = res.newData;
          if (res.newSelection) currentSelection = res.newSelection;
        } else if (type === 'math_col_col') {
          // COL MATH
          const res = applyColMath(currentData, currentSelection, payload?.option?.operator || 'sum', payload?.option?.resultMode);
          currentData = res.newData;
          if (res.newSelection) currentSelection = res.newSelection;
        }
        else if (type === 'style') {
          // STYLE
          const res = applyStyleUtil(currentData, currentSelection, payload.type, payload.value);
          currentData = res.newData;
          if (res.newSelection) currentSelection = res.newSelection;
        }
        else if (type === 'move') {
          // NAVIGATION
          const res = applyMove(currentSelection, payload);
          if (res) currentSelection = res;
        }
        else if (type === 'logic_if') {
          const condition = { operator: payload.operator, value: payload.value };
          const res = applyLogicIf(currentData, currentSelection, condition, payload.trueVal, payload.falseVal);
          currentData = res.newData;
          if (res.newSelection) currentSelection = res.newSelection;
        }

        // 50ms Visual Delay per step (Optional for "Sequential feeling", maybe skip for performance)
        await new Promise(r => setTimeout(r, 100));
        // Force intermediate render? React state update is batch, so we can't easily animate intermediate steps 
        // UNLESS we updateData() in loop. 
        // User said "Sequential... update temp... setGrid once at end". 
        // BUT "Step 1 -> Step 2". If step 2 depends on Step 1 output, passing `currentData` works.
        // So we don't NEED intermediate renders for logic correctness.
        // We will just wait for loop to finish.
      } catch (e) {
        console.error("Recipe Error:", e);
        alert(`실패: ${action.desc}`);
        break;
      }
    }

    // Final Commit
    updateData(currentData);
    setSelection(currentSelection); // Move the cursor practically at end
    alert("레시피 실행 완료!");
  };

  const handleAction = (action: string | any) => {
    // ... existing dispatch logic ...
    if (typeof action === 'object' && action.type === 'run_recipe') {
      handleRunRecipe(action.recipe);
      return;
    }

    if (typeof action === 'object' && action.type === 'unified_calc') {
      handleUnifiedCalculation(action.payload);
      return;
    }
    if (typeof action === 'object' && action.type === 'style') {
      handleStyle(action.payload);
      return;
    }
    if (typeof action === 'object' && action.type === 'style') {
      handleStyle(action.payload);
      return;
    }
    if (typeof action === 'object' && action.type === 'logic_if') {
      handleLogicIf(action.payload || action.option);
      return;
    }
    if (typeof action === 'object' && action.type === 'logic_vlookup') {
      handleLogicVlookup(action.payload || action.option);
      return;
    }

    // If action is object, it's a recipe
    if (typeof action === 'object' && action.type) {
      if (['sum', 'count', 'avg'].includes(action.type)) handleCalculate(action);
      else if (['join', 'split', 'extract'].includes(action.type)) handleTextAction(action);
      else if (['clean_empty', 'remove_dup', 'trim'].includes(action.type)) {
        if (action.type === 'clean_empty') handleCleanEmptyRows();
        if (action.type === 'remove_dup') handleRemoveDuplicates();
        if (action.type === 'trim') handleTrimSpaces(action.option);
      }
      else if (['comma', 'header_style', 'highlight'].includes(action.type)) {
        if (action.type === 'comma') handleAddCommas();
        if (action.type === 'header_style') handleHeaderStyle();
        if (action.type === 'highlight') handleHighlight(action.option);
      }
      else if (action.type.startsWith('chart') || action.type.startsWith('stat')) {
        handleAnalyze(action);
      }
      else if (action.type === 'math_col_col') {
        if (action.option?.resultMode === 'pick') {
          setIsPickingTarget(true);
          setPendingAction(action);
          alert("🎯 결과를 표시할 칸을 더블 클릭하세요!");
          return;
        }
        handleMathColCol(action.option?.operator || '+', action.option?.resultMode, action.option?.targetStart);
      }
      else if (action.type === 'math_row_row') {
        if (action.option?.resultMode === 'pick') {
          setIsPickingTarget(true);
          setPendingAction(action);
          alert("🎯 결과를 표시할 칸을 더블 클릭하세요!");
          return;
        }
        handleMathRowRow(action.option?.operator || '+', action.option?.resultMode, action.option?.targetStart);
      }
      else if (action.type === 'math_col_const') {
        if (action.option?.resultMode === 'pick') {
          setIsPickingTarget(true);
          setPendingAction(action);
          alert("🎯 결과를 표시할 칸을 더블 클릭하세요!");
          return;
        }
        handleMathColConst(action.option?.operator || '+', action.option?.value, action.option?.resultMode, action.option?.targetStart);
      }
      // [Data] Actions
      else if (['sort_asc', 'sort_desc', 'filter', 'replace'].includes(action.type)) {
        if (action.type === 'sort_asc') handleSort('asc', action.option);
        if (action.type === 'sort_desc') handleSort('desc', action.option);
        if (action.type === 'filter') handleFilter(action.option?.condition, action.option);
        if (action.type === 'replace') handleReplace(action.option?.find, action.option?.replace);
      }
      return;
    }

    switch (action) {
      case "clean_empty_rows": handleCleanEmptyRows(); break;
      case "set_header": handleHeaderStyle(); break;
      case "remove_duplicates": handleRemoveDuplicates(); break;
      case "trim_spaces": handleTrimSpaces(); break;
      case "add_commas": handleAddCommas(); break;
      case "style_header": handleHeaderStyle(); break;
      case "undo": handleUndo(); break;
      case "redo": handleRedo(); break;
      case "copy": handleCopy(); break;
      case "paste": handlePaste(); break;
      case "save": handleSave(); break;
      default: console.log("Unknown action:", action);
    }
  };

  return (
    <main className={`h-screen w-screen overflow-hidden flex flex-col lg:flex-row relative ${isDarkMode ? "bg-slate-900" : "bg-gray-50"}`}>

      {/* --- MOBILE HEADER (Visible only on small screens < lg) --- */}
      <header className="lg:hidden flex-none h-14 bg-slate-900 border-b border-slate-800 flex items-center justify-between px-4 z-40">
        <div className="flex items-center gap-3">
          <button onClick={() => setIsMobileSidebarOpen(true)} className="p-2 text-slate-300 hover:bg-slate-800 rounded-lg">
            <Menu size={24} />
          </button>
          <div className="flex items-center gap-2 text-indigo-400 font-bold">
            <Layout size={20} />
            <span>Kind Sheet</span>
          </div>
        </div>
        <div className="text-xs font-mono text-slate-500 bg-slate-800 px-2 py-1 rounded">
          {currentSheetName}
        </div>
      </header>

      {/* --- DESKTOP/MOBILE SPLIT LAYOUT --- */}
      {/* Container: Flex Col on Mobile, Row on Desktop */}
      <div className="flex-1 flex flex-col lg:flex-row gap-6 p-4 sm:p-6 overflow-hidden relative">

        {/* 1. Grid Area (Takes remaining space on Mobile, 60% on Desktop) */}
        <section className="flex-1 w-full lg:w-[60%] h-full relative overflow-hidden transition-all duration-500 ease-in-out">
          <SheetPreview
            excelData={excelData}
            fileName={fileName}
            onFileSelect={handleFileSelect}
            onClear={handleClear}
            onUpdateData={(newData) => updateData(newData)}
            selection={selection}
            onSelectionChange={setSelection}
            isDarkMode={isDarkMode}
            onOpen={handleOpenFile}
            // Multi-Sheet Props
            sheets={Object.keys(worksheets)}
            currentSheet={currentSheetName}
            onChangeSheet={handleChangeSheet}
            onAddSheet={handleAddSheet}
            onDeleteSheet={handleDeleteSheet}
            onInsertRow={handleInsertRow}
            onDeleteRow={handleDeleteRow}
            onInsertCol={handleInsertCol}
            onDeleteCol={handleDeleteCol}
            onClearContent={handleClearContent}
            onUndo={handleUndo}
            onRedo={handleRedo}
            // Explicit View State
            isFileLoaded={isFileLoaded}
            onStartEmpty={handleStartEmpty}
            // Target Pick Mode
            isPickingTarget={isPickingTarget}
            onCellClick={handleTargetPick}
            onConfirmTarget={handleTargetConfirm}
          />
        </section>

        {/* 2. Control Panel (Bottom Fixed on Mobile, Right Fixed on Desktop) */}
        <section
          className={cn(
            "z-40 shadow-[0_-4px_10px_rgba(0,0,0,0.2)] bg-slate-800",
            // Mobile: Stacked at bottom, auto height (Lite Mode)
            "w-full h-auto",
            // Desktop: Right side, full height, fixed width
            "lg:w-[40%] lg:h-full lg:relative"
          )}
        >
          <ControlPanel
            className="rounded-none lg:rounded-3xl h-full w-full"
            onAction={handleAction}
            onExport={handleDownload}
            onOpen={handleOpenFile}
            onSave={handleInstantSave}
            hasData={excelData.length > 0}
            alignment={alignment}
            onAlignmentChange={setAlignment}
            canUndo={history.length > 0}
            canRedo={future.length > 0}
            selection={selection}
            isDarkMode={isDarkMode}
            onToggleTheme={() => setIsDarkMode(!isDarkMode)}
          />
        </section>
      </div>

      {/* --- MOBILE OVERLAYS --- */}

      {/* Mobile Sidebar (Drawer) */}
      {isMobileSidebarOpen && (
        <div className="fixed inset-0 z-50 flex lg:hidden">
          <div className="fixed inset-0 bg-black/60 backdrop-blur-sm" onClick={() => setIsMobileSidebarOpen(false)}></div>
          <div className="relative w-[80%] max-w-xs h-full bg-slate-900 border-r border-slate-800 shadow-2xl p-6 flex flex-col gap-6 animate-in slide-in-from-left duration-300">
            <div className="flex items-center justify-between">
              <h2 className="text-xl font-bold text-white flex items-center gap-2">
                <Layout size={24} className="text-indigo-500" />
                <span>Sheets</span>
              </h2>
              <button onClick={() => setIsMobileSidebarOpen(false)}><X className="text-slate-500 hover:text-white" /></button>
            </div>

            {/* Sheet List */}
            <div className="flex-1 overflow-y-auto flex flex-col gap-2">
              {Object.keys(worksheets).map(sheet => (
                <button
                  key={sheet}
                  onClick={() => { handleChangeSheet(sheet); setIsMobileSidebarOpen(false); }}
                  className={cn(
                    "p-4 rounded-xl text-left font-bold transition-all flex items-center gap-3",
                    currentSheetName === sheet ? "bg-indigo-600 text-white" : "bg-slate-800 text-slate-400 hover:bg-slate-700"
                  )}
                >
                  <Sheet size={18} />
                  {sheet}
                </button>
              ))}
              <button onClick={handleAddSheet} className="p-4 rounded-xl border-2 border-dashed border-slate-700 text-slate-500 font-bold hover:border-indigo-500 hover:text-indigo-500 transition-colors flex items-center justify-center gap-2">
                <Plus size={18} /> New Sheet
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Analysis Modal */}
      {
        isAnalysisOpen && analysisResult && (
          <div className="absolute inset-0 z-50 bg-slate-900/80 backdrop-blur-sm flex items-center justify-center p-8 animate-in fade-in zoom-in duration-200" onClick={() => setIsAnalysisOpen(false)}>
            <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90%] flex flex-col overflow-hidden" onClick={e => e.stopPropagation()}>
              <div className="p-4 border-b border-gray-200 dark:border-slate-700 flex justify-between items-center bg-gray-50 dark:bg-slate-900">
                <h2 className="text-lg font-bold text-slate-800 dark:text-white flex items-center gap-2">📊 분석 결과</h2>
                <div className="flex items-center gap-2">
                  {analysisResult.type === 'stat' && (
                    <button onClick={handleSaveAnalysisToSheet} className="text-sm px-3 py-1.5 bg-indigo-600 text-white rounded-lg hover:bg-indigo-500 transition-colors shadow-sm font-bold">
                      💾 시트로 저장
                    </button>
                  )}
                  <button onClick={() => setIsAnalysisOpen(false)} className="p-2 hover:bg-slate-200 dark:hover:bg-slate-700 rounded-full transition-colors text-slate-500">
                    <X size={20} />
                  </button>
                </div>
              </div>

              <div className="p-6 overflow-auto flex-1 bg-white dark:bg-slate-800">
                {analysisResult.type === 'stat' && (
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                    {[
                      { label: '데이터 개수', val: analysisResult.stats.count },
                      { label: '합계', val: analysisResult.stats.sum.toLocaleString() },
                      { label: '평균', val: analysisResult.stats.mean.toLocaleString(undefined, { maximumFractionDigits: 2 }) },
                      { label: '중앙값', val: analysisResult.stats.median.toLocaleString() },
                      { label: '최솟값', val: analysisResult.stats.min.toLocaleString() },
                      { label: '최댓값', val: analysisResult.stats.max.toLocaleString() },
                      { label: '표준편차', val: analysisResult.stats.stdDev.toLocaleString(undefined, { maximumFractionDigits: 2 }) },
                    ].map((item, i) => (
                      <div key={i} className="bg-slate-50 dark:bg-slate-700 p-4 rounded-xl border border-slate-100 dark:border-slate-600">
                        <p className="text-sm text-slate-500 dark:text-slate-400 font-bold mb-1">{item.label}</p>
                        <p className="text-2xl font-bold text-indigo-600 dark:text-indigo-400">{item.val}</p>
                      </div>
                    ))}
                  </div>
                )}

                {analysisResult.type === 'bar' && <Bar data={analysisResult.data} options={analysisResult.options} />}
                {analysisResult.type === 'line' && <Line data={analysisResult.data} options={analysisResult.options} />}
                {analysisResult.type === 'scatter' && <Scatter data={analysisResult.data} options={analysisResult.options} />}
              </div>
            </div>
          </div>
        )
      }
    </main >
  );
}

