"use client";

import React, { useCallback, useState, useEffect, useRef, useLayoutEffect } from "react";
import { FileSpreadsheet, Upload, X, Plus, ArrowUp, ArrowDown, ArrowLeft, ArrowRight, Trash, Eraser } from "lucide-react";
import { clsx } from "clsx";
import { twMerge } from "tailwind-merge";

function cn(...inputs: (string | undefined | null | false)[]) {
    return twMerge(clsx(inputs));
}

// --- Phantom Input Component ---
// A single global input that moves around the grid to prevent mounting delays
interface PhantomInputProps {
    activeRect: { top: number; left: number; width: number; height: number } | null;
    initialValue: string; // The value to start with (empty for overwrite, text for F2)
    currentCellValue: string; // The underlying cell value (for visual context if needed, mostly handled by grid)
    onCommit: (val: string, move?: { r: number, c: number }, overrideCoords?: { r: number, c: number }) => void;
    onCancel: () => void;
    onDeleteRange: () => void;
    isDarkMode: boolean;
    r: number;
    c: number;
}

const PhantomInput = ({ activeRect, initialValue, currentCellValue, onCommit, onCancel, onDeleteRange, isDarkMode, r, c }: PhantomInputProps) => {
    const inputRef = useRef<HTMLInputElement>(null);
    const [val, setVal] = useState(initialValue);
    const [isEditing, setIsEditing] = useState(false); // True = Visible (Opacity 1), False = Hidden (Opacity 0)

    // Fix Fly-away: Store the coordinates where editing STARTED.
    // When committing, use these coords instead of current props (which might have changed if user moved fast)
    const editingCoords = useRef<{ r: number, c: number } | null>(null);

    // Sync state when selection/initialValue changes
    useEffect(() => {
        setVal(initialValue);
        // If initialValue is populated (e.g. F2), we are editing immediately
        if (initialValue) {
            setIsEditing(true);
            editingCoords.current = { r, c };
        } else {
            setIsEditing(false);
            editingCoords.current = null;
        }

        // Always focus when activeRect changes (selection moved) so typing works immediately
        if (activeRect && inputRef.current) {
            inputRef.current.focus();
        }
    }, [activeRect, initialValue, r, c]);

    if (!activeRect) return null;

    const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        setVal(e.target.value);
        if (!isEditing) {
            setIsEditing(true); // Make visible on first type
            // Capture coords if we just started typing (overwrite mode)
            if (!editingCoords.current) {
                editingCoords.current = { r, c };
            }
        }
    };

    const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
        if (e.nativeEvent.isComposing) return;

        const move = (dr: number, dc: number) => {
            e.preventDefault();
            // Pass the captured coords to commit
            const commitVal = isEditing ? val : null;
            onCommit(commitVal as any, { r: dr, c: dc }, editingCoords.current || undefined);
        };

        if (e.key === "Enter") {
            if (e.shiftKey) move(-1, 0); else move(1, 0);
        }
        else if (e.key === "Tab") {
            if (e.shiftKey) move(0, -1); else move(0, 1);
        }
        else if (e.key === "ArrowUp") move(-1, 0);
        else if (e.key === "ArrowDown") move(1, 0);
        else if (e.key === "ArrowLeft") {
            if (isEditing && inputRef.current && (inputRef.current.selectionStart || 0) > 0) {
                return;
            }
            move(0, -1);
        }
        else if (e.key === "ArrowRight") {
            if (isEditing && inputRef.current && (inputRef.current.selectionStart || 0) < val.length) {
                return;
            }
            move(0, 1);
        }
        else if (e.key === "Escape") {
            e.preventDefault();
            setIsEditing(false);
            editingCoords.current = null;
            setVal(""); // Reset
            onCancel();
        }
        else if ((e.key === "Delete" || e.key === "Backspace") && !isEditing) {
            // Quick Delete in Ghost Mode
            e.preventDefault();
            // Commit empty string to current coords
            // Override coords might not be needed if we are in ghost mode (selection is synced), 
            // but safer to use props r, c directly or null.
            onCommit("", undefined, { r, c });
        }
    };

    const handleBlur = () => {
        // Commit if editing
        if (isEditing) {
            onCommit(val, undefined, editingCoords.current || undefined);
        }
    };

    return (
        <input
            ref={inputRef}
            type="text"
            value={val}
            onChange={handleChange}
            onKeyDown={handleKeyDown}
            onBlur={handleBlur}
            style={{
                position: 'absolute',
                top: activeRect.top,
                left: activeRect.left,
                width: activeRect.width,
                height: activeRect.height,
                opacity: isEditing ? 1 : 0,
                zIndex: 50,
                pointerEvents: isEditing ? 'auto' : 'none',
            }}
            className={cn(
                "px-2 outline-none ring-2 text-sm font-sans transition-opacity duration-75",
                isDarkMode ? "bg-slate-800 text-white ring-indigo-500" : "bg-white text-gray-900 ring-green-600"
            )}
            autoComplete="off"
        />
    );
};

// Helper: Get Cell Style
const getCellStyle = (cell: any, r: number, c: number, selection: any, isFilling: boolean, isDarkMode: boolean): React.CSSProperties => {
    const isSelected = selection &&
        r >= Math.min(selection.start.r, selection.end.r) &&
        r <= Math.max(selection.start.r, selection.end.r) &&
        c >= Math.min(selection.start.c, selection.end.c) &&
        c <= Math.max(selection.start.c, selection.end.c);

    const isInFillArea = isFilling && selection &&
        r >= Math.min(selection.start.r, selection.end.r) && r <= Math.max(selection.start.r, selection.end.r) &&
        c >= Math.min(selection.start.c, selection.end.c) && c <= Math.max(selection.start.c, selection.end.c);

    let style: React.CSSProperties = {
        minWidth: '80px',
        height: '30px',
        border: '1px solid #e2e8f0',
        padding: '4px 8px',
        outline: 'none',
        whiteSpace: 'nowrap',
        overflow: 'hidden',
    };

    if (cell && typeof cell === 'object' && cell.s) {
        const s = cell.s;
        if (s.font) {
            if (s.font.bold) style.fontWeight = 'bold';
            if (s.font.italic) style.fontStyle = 'italic';
            if (s.font.underline) style.textDecoration = 'underline';
            if (s.font.color && s.font.color.rgb) style.color = '#' + s.font.color.rgb;
        }
        if (s.alignment && s.alignment.horizontal) {
            style.textAlign = s.alignment.horizontal as any;
        }
        if (s.fill && s.fill.fgColor && s.fill.fgColor.rgb) {
            style.backgroundColor = '#' + s.fill.fgColor.rgb;
        }
    }

    // Default styles override
    if (!style.textAlign) {
        // Extract value safely
        let v = cell;
        if (cell && typeof cell === 'object' && 'v' in cell) v = cell.v;
        if (typeof v === 'number') style.textAlign = 'right';
        else style.textAlign = 'left';
    }

    if (isSelected) {
        style.backgroundColor = 'rgba(59, 130, 246, 0.1)';
        style.border = '1px double #3b82f6';
    }

    if (isInFillArea) {
        style.backgroundColor = 'rgba(16, 185, 129, 0.1)';
        style.border = '1px dashed #10b981';
    }

    // Dark mode borders
    if (isDarkMode) {
        style.borderColor = 'rgba(51, 65, 85, 0.5)'; // slate-700/50
    }

    return style;
};


interface SheetPreviewProps {
    excelData: any[][];
    fileName: string | null;
    onFileSelect: (file: File) => void;
    onClear: () => void;
    onUpdateData: (newData: any[][]) => void;
    selection: { start: { r: number; c: number }; end: { r: number; c: number } } | null;
    onSelectionChange: (sel: { start: { r: number; c: number }; end: { r: number; c: number } } | null) => void;
    isDarkMode: boolean;
    onOpen?: () => void;
    sheets?: string[];
    currentSheet?: string;
    onChangeSheet?: (sheetName: string) => void;
    onAddSheet?: () => void;
    onDeleteSheet?: (sheetName: string) => void;
    isFileLoaded: boolean;
    onStartEmpty: () => void;
    onInsertRow?: (rowIndex: number, position: 'above' | 'below') => void;
    onDeleteRow?: (rowIndex: number) => void;
    onInsertCol?: (colIndex: number, position: 'left' | 'right') => void;
    onDeleteCol?: (colIndex: number) => void;
    onClearContent?: (rowIndex: number, colIndex: number) => void;
    onUndo?: () => void;
    onRedo?: () => void;
    isPickingTarget?: boolean;
    onCellClick?: (r: number, c: number) => void;
    onConfirmTarget?: (r: number, c: number) => void;

    onAutofill?: (source: { start: { r: number, c: number }, end: { r: number, c: number } }, target: { start: { r: number, c: number }, end: { r: number, c: number } }) => void;
    onDeleteRange?: () => void;
}

function getColumnLabel(index: number): string {
    let label = "";
    let i = index;
    while (i >= 0) {
        label = String.fromCharCode((i % 26) + 65) + label;
        i = Math.floor(i / 26) - 1;
    }
    return label;
}

export default function SheetPreview({
    excelData,
    fileName,
    onFileSelect,
    onClear,
    onUpdateData,
    selection,
    onSelectionChange,
    isDarkMode,
    onOpen,
    sheets,
    currentSheet,
    onChangeSheet,
    onAddSheet,
    onDeleteSheet,
    isFileLoaded,
    onStartEmpty,
    onInsertRow,
    onDeleteRow,
    onInsertCol,
    onDeleteCol,
    onClearContent,
    onUndo,
    onRedo,
    onDeleteRange,
    isPickingTarget = false,
    onCellClick,
    onConfirmTarget,
    onAutofill,
}: SheetPreviewProps) {
    const [isDraggingFile, setIsDraggingFile] = useState(false);
    const [contextMenu, setContextMenu] = useState<{ visible: boolean; x: number; y: number; r: number; c: number } | null>(null);

    // Grid Interaction State
    const [isSelecting, setIsSelecting] = useState(false);
    const [isFilling, setIsFilling] = useState(false);
    const [fillStartSelection, setFillStartSelection] = useState<{ start: { r: number; c: number }; end: { r: number; c: number } } | null>(null);

    // Phantom Input State
    const [activeRect, setActiveRect] = useState<{ top: number; left: number; width: number; height: number } | null>(null);
    const [initVal, setInitVal] = useState("");

    const isWorkbookLoaded = (sheets && sheets.length > 0) || excelData.length > 0;
    const showGrid = isWorkbookLoaded;

    const dataRows = excelData.length;
    const dataCols = dataRows > 0 ? Math.max(...excelData.map(r => r?.length || 0)) : 0;
    const displayRows = Math.max(dataRows + 20, 100);
    const displayCols = Math.max(dataCols + 10, 26);

    // Theme Configuration
    const theme = {
        bg: isDarkMode ? "bg-slate-800" : "bg-white",
        containerBg: isDarkMode ? "bg-slate-900/50" : "bg-white",
        borderColor: isDarkMode ? "border-slate-700/50" : "border-gray-200",
        headerBg: isDarkMode ? "bg-slate-700" : "bg-gray-50",
        headerText: isDarkMode ? "text-slate-300" : "text-gray-600",
        headerActive: isDarkMode ? "bg-indigo-500/20 text-indigo-300" : "bg-green-100 text-green-800 font-bold",
        cellText: isDarkMode ? "text-slate-300" : "text-gray-900",
        cellBorder: isDarkMode ? "border-slate-700/50" : "border-gray-200",
        selectionOverlay: isDarkMode ? "bg-indigo-500/10" : "bg-green-500/10",
        selectionBorder: isDarkMode ? "bg-indigo-500" : "bg-green-600",
    };

    // Calculate Active Rect when selection changes
    useEffect(() => {
        if (!selection) {
            setActiveRect(null);
            return;
        }
        const { r, c } = selection.start;
        // Use timeout to allow render to settle if needed, but usually immediate is fine
        // Using requestAnimationFrame to ensure DOM is ready
        requestAnimationFrame(() => {
            const el = document.getElementById(`cell-${r}-${c}`);
            if (el) {
                setActiveRect({
                    top: el.offsetTop,
                    left: el.offsetLeft,
                    width: el.offsetWidth,
                    height: el.offsetHeight
                });
                setInitVal(""); // Default to overwrite mode
            }
        });
    }, [selection, excelData]);

    // File Drag & Drop
    const handleDragOver = useCallback((e: React.DragEvent) => {
        e.preventDefault();
        if (e.dataTransfer.types.includes("Files")) setIsDraggingFile(true);
    }, []);

    const handleDragLeave = useCallback((e: React.DragEvent) => {
        e.preventDefault();
        setIsDraggingFile(false);
    }, []);

    const handleDrop = useCallback((e: React.DragEvent) => {
        e.preventDefault();
        setIsDraggingFile(false);
        const file = e.dataTransfer.files[0];
        if (file && (file.name.endsWith(".xlsx") || file.name.endsWith(".xls"))) {
            onFileSelect(file);
        }
    }, [onFileSelect]);


    // Data Updates
    const commitPhantom = (val: any, move?: { r: number, c: number }, overrideCoords?: { r: number, c: number }) => {
        if (!selection && !overrideCoords) return;
        const { r, c } = overrideCoords || selection!.start;

        if (val !== null) {
            const currentRow = excelData[r] || [];
            const currentValueObj = currentRow[c];
            let currentRawValue = currentValueObj;
            if (currentValueObj && typeof currentValueObj === 'object' && 'v' in currentValueObj) {
                currentRawValue = currentValueObj.v;
            }

            let finalValue: any = val;
            const trimmed = String(val).trim();
            if (trimmed !== "" && !isNaN(Number(trimmed))) finalValue = Number(trimmed);

            if (String(currentRawValue ?? "") !== String(finalValue)) {
                const newData = [...excelData];
                if (!newData[r]) {
                    for (let i = excelData.length; i <= r; i++) newData[i] = [];
                }
                newData[r] = [...(newData[r] || [])];

                if (currentValueObj && typeof currentValueObj === 'object') {
                    newData[r][c] = { ...currentValueObj, v: finalValue };
                } else {
                    newData[r][c] = finalValue;
                }
                onUpdateData(newData);
            }
        }

        if (move) {
            const nr = Math.max(0, r + move.r);
            const nc = Math.max(0, c + move.c);
            onSelectionChange({ start: { r: nr, c: nc }, end: { r: nr, c: nc } });
        }
    };

    // Events
    const handleMouseDown = (r: number, c: number, e: React.MouseEvent) => {
        if (e.button !== 0) return;
        if (isPickingTarget) {
            onCellClick?.(r, c);
            return;
        }
        setIsSelecting(true);
        onSelectionChange({ start: { r, c }, end: { r, c } });
    };

    const handleMouseEnter = (r: number, c: number) => {
        if ((isSelecting || isFilling) && selection) {
            onSelectionChange({ ...selection, end: { r, c } });
        }
    };

    const handleMouseUp = () => {
        if (isFilling && selection && fillStartSelection) {
            onAutofill?.(fillStartSelection, selection);
        }
        setIsSelecting(false);
        setIsFilling(false);
        setFillStartSelection(null);
    };

    const handleFillMouseDown = (e: React.MouseEvent) => {
        e.stopPropagation();
        e.preventDefault();
        if (!selection) return;
        setIsFilling(true);
        setFillStartSelection(selection);
        setIsSelecting(true);
    };

    // Global Event Listeners
    useEffect(() => {
        window.addEventListener("mouseup", handleMouseUp);
        return () => window.removeEventListener("mouseup", handleMouseUp);
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [isFilling, selection, fillStartSelection]);

    // F2 Key
    useEffect(() => {
        const handleKeyDown = (e: KeyboardEvent) => {
            if (e.key === "F2" && selection) {
                const { r, c } = selection.start;
                const cell = excelData[r]?.[c];
                let v = "";
                if (cell && typeof cell === 'object' && 'v' in cell) v = String(cell.v);
                else if (cell !== null && cell !== undefined) v = String(cell);

                setInitVal(v);
                // Note: Updating initVal triggers effect in PhantomInput to set isEditing=true
            }
            // Undo/Redo
            if ((e.ctrlKey || e.metaKey) && e.key === "z") {
                e.preventDefault();
                if (e.shiftKey) onRedo?.();
                else onUndo?.();
            }
            if ((e.ctrlKey || e.metaKey) && (e.key === "y" || e.key === "Y")) {
                e.preventDefault();
                onRedo?.();
            }

            // Quick Delete (Fix 1)
            if ((e.key === "Delete" || e.key === "Backspace") && selection) {
                // If input is focused and editing, we shouldn't trigger this (the input handles it).
                // PhantomInput always has focus if active. But if isEditing=false (ghost), we should delete cell.
                // If isEditing=true, we are typing, so let input handle backspace.

                // We can check if the active element is our input AND if we are 'editing' (opacity 1)
                // But simpler: The PhantomInput component swallows the key event if it's focused.
                // However, 'ghost' input propagates keys? NO, ghost input also captures keys.

                // Wait, if PhantomInput captures keys, global listener won't get them if we stopPropagation in Phantom.
                // In PhantomInput handleKeyDown:
                // if (isEditing) -> propagate Backspace to input (return)
                // if (!isEditing) -> ghost mode.

                // Currently PhantomInput handleKeyDown doesn't stop prop for default keys unless overridden.

                // Better approach: Handle "Delete" logic here, but guard against "Input Focused".
                const activeTag = document.activeElement?.tagName.toLowerCase();
                // If we are editing, activeElement is input.
                // BUT, how do we know if it's "Editing" vs "Ghost"?
                // We don't have access to PhantomInput state here.

                // Actually, if we are in Ghost mode, we want Delete to clear cell.
                // If we are in Edit mode, we want Delete to delete a character.

                // Solution: Check if the phantom input value is empty? No.
                // If we are in ghost mode, the phantom input value is usually empty (or initVal).

                // Let's rely on event bubbling prevention. 
                // If PhantomInput handles the key, it should stop propagation if it's actually editing.

                // Actually, just checking `document.activeElement` is tricky if Ghost always has focus.

                // Alternative: Add data attribute to Input to signal state?

                // Let's assume: If user hits Delete, and they are NOT typing (just moving around), cell clears.
                // If they are typing, the input consumes the key naturally (Delete char).

                // If I am typing in input, global listener fires?
                // Yes, unless stopPropagation.
                // PhantomInput doesn't stop prop for everything.

                if (activeTag === 'input' || activeTag === 'textarea') {
                    // Check if it's OUR phantom input
                    // We can't easily distinguish. 

                    // BUT: changing delete logic in PhantomInput handleKeyDown is cleaner!
                    // Let's do it in PhantomInput instead of global here?
                    // "Quick Delete" implies selecting a cell and hitting Delete.
                    // The PhantomInput IS focused. So PhantomInput handleKeyDown will catch "Delete".
                    // I will verify this and add logic there if needed.

                    // Actually, the user asked for "Quick Delete" on selection mode.
                    // In selection mode (ghost), PhantomInput catches keys.
                    // I updated PhantomInput above to allow this? No, I only touched Fly-away logic.
                    // I should add "Delete" handling to PhantomInput handleKeyDown.
                } else {
                    // Fallback for when input loses focus (rare but possible)
                    onUpdateData(excelData.map((row, rowIndex) =>
                        row.map((cell, colIndex) => {
                            if (isSelected(rowIndex, colIndex)) return "";
                            return cell;
                        })
                    ));
                    // Wait, this is inefficient for single cell.
                    // Let's stick to doing it inside PhantomInput or properly here.
                }
            }
        };
        window.addEventListener("keydown", handleKeyDown);
        return () => window.removeEventListener("keydown", handleKeyDown);
    }, [selection, excelData, onUndo, onRedo]);

    const isSelected = (r: number, c: number) => {
        if (!selection) return false;
        const r1 = Math.min(selection.start.r, selection.end.r);
        const r2 = Math.max(selection.start.r, selection.end.r);
        const c1 = Math.min(selection.start.c, selection.end.c);
        const c2 = Math.max(selection.start.c, selection.end.c);
        return r >= r1 && r <= r2 && c >= c1 && c <= c2;
    };

    // Context Menu Logic skipped for brevity? No, user needs functionality.
    const handleContextMenu = useCallback((e: React.MouseEvent, r: number, c: number) => {
        e.preventDefault();
        setContextMenu({ visible: true, x: e.clientX, y: e.clientY, r, c });
    }, []);

    useEffect(() => {
        const handleClick = () => setContextMenu(null);
        window.addEventListener("click", handleClick);
        return () => window.removeEventListener("click", handleClick);
    }, []);


    return (
        <div className={cn("h-full w-full rounded-3xl overflow-hidden flex flex-col shadow-2xl shadow-slate-950/20 border transition-all relative", theme.bg, theme.borderColor)}>
            {/* Header */}
            <div className={cn("flex items-center justify-between p-4 px-6 border-b backdrop-blur-sm z-20", theme.borderColor, isDarkMode ? "bg-slate-800/80" : "bg-white/80")}>
                <div className="flex items-center gap-3">
                    <div className={cn("p-2 rounded-lg", isDarkMode ? "bg-green-500/10 text-green-400" : "bg-green-100 text-green-700")}>
                        <FileSpreadsheet size={20} />
                    </div>
                    <div className="flex flex-col">
                        <span className={cn("text-sm font-bold", isDarkMode ? "text-slate-200" : "text-gray-800")}>{fileName || "새로운 스프레드시트"}</span>
                        <span className="text-xs text-slate-500">{showGrid ? `${dataRows}행 데이터` : "데이터 없음"}</span>
                    </div>
                </div>
                {showGrid && (
                    <button
                        onClick={onClear}
                        className={cn("text-xs transition-colors px-3 py-1.5 rounded-lg", isDarkMode ? "text-slate-500 hover:text-red-400 hover:bg-slate-700/50" : "text-gray-500 hover:text-red-500 hover:bg-gray-100")}
                    >
                        초기화
                    </button>
                )}
            </div>

            {/* Grid */}
            <div
                className={cn("flex-1 overflow-auto relative custom-scrollbar cursor-cell", theme.containerBg)}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
            >
                {/* Phantom Input Overlay */}
                <PhantomInput
                    activeRect={activeRect}
                    initialValue={initVal}
                    currentCellValue=""
                    onCommit={commitPhantom}
                    onCancel={() => { }}
                    onDeleteRange={onDeleteRange || (() => { })}
                    isDarkMode={isDarkMode}
                    r={selection ? selection.start.r : 0}
                    c={selection ? selection.start.c : 0}
                />

                <table className="border-collapse w-max table-fixed">
                    <thead className={cn("sticky top-0 z-30 text-xs shadow-sm", theme.headerBg)}>
                        <tr>
                            <th className={cn("sticky left-0 top-0 z-40 w-12 min-w-[3rem] h-8 border-r border-b", theme.headerBg, theme.borderColor)}></th>
                            {Array.from({ length: displayCols }).map((_, colIndex) => {
                                const isColSelected = selection && colIndex >= Math.min(selection.start.c, selection.end.c) && colIndex <= Math.max(selection.start.c, selection.end.c);
                                return (
                                    <th
                                        key={colIndex}
                                        className={cn(
                                            "px-2 h-8 font-medium text-center border-r border-b min-w-[100px] select-none transition-colors cursor-pointer",
                                            theme.borderColor,
                                            isColSelected ? theme.headerActive : theme.headerText
                                        )}
                                    >
                                        {getColumnLabel(colIndex)}
                                    </th>
                                );
                            })}
                        </tr>
                    </thead>
                    <tbody className="text-sm font-sans">
                        {Array.from({ length: displayRows }).map((_, rowIndex) => {
                            const isRowSelected = selection && rowIndex >= Math.min(selection.start.r, selection.end.r) && rowIndex <= Math.max(selection.start.r, selection.end.r);
                            return (
                                <tr key={rowIndex} className="h-8">
                                    <td className={cn(
                                        "sticky left-0 z-20 w-12 min-w-[3rem] text-center text-xs border-r border-b select-none transition-colors",
                                        isPickingTarget ? "cursor-crosshair" : "cursor-pointer",
                                        theme.borderColor, theme.headerBg,
                                        isRowSelected ? theme.headerActive : theme.headerText
                                    )}>
                                        {rowIndex + 1}
                                    </td>
                                    {Array.from({ length: displayCols }).map((_, colIndex) => {
                                        const cell = excelData[rowIndex]?.[colIndex];
                                        const selected = isSelected(rowIndex, colIndex);

                                        // Edges for selection border
                                        let edges = { top: false, bottom: false, left: false, right: false, isCorner: false };
                                        if (selection) {
                                            const r1 = Math.min(selection.start.r, selection.end.r);
                                            const r2 = Math.max(selection.start.r, selection.end.r);
                                            const c1 = Math.min(selection.start.c, selection.end.c);
                                            const c2 = Math.max(selection.start.c, selection.end.c);
                                            edges = {
                                                top: rowIndex === r1 && colIndex >= c1 && colIndex <= c2,
                                                bottom: rowIndex === r2 && colIndex >= c1 && colIndex <= c2,
                                                left: colIndex === c1 && rowIndex >= r1 && rowIndex <= r2,
                                                right: colIndex === c2 && rowIndex >= r1 && rowIndex <= r2,
                                                isCorner: rowIndex === r2 && colIndex === c2
                                            };
                                        }

                                        const cellStyle = getCellStyle(cell, rowIndex, colIndex, selection, isFilling, isDarkMode);

                                        let cellValue = "";
                                        if (cell && typeof cell === 'object' && 'v' in cell) cellValue = cell.v;
                                        else cellValue = cell;

                                        /* Important: ID for phantom input tracking */
                                        const cellId = `cell-${rowIndex}-${colIndex}`;

                                        return (
                                            <td
                                                key={`${rowIndex}-${colIndex}`}
                                                id={cellId}
                                                className={cn(
                                                    "border-r border-b min-w-[100px] max-w-[300px] relative p-0 overflow-visible",
                                                    isPickingTarget ? "cursor-crosshair" : "cursor-cell",
                                                    theme.cellBorder,
                                                    theme.cellText,
                                                    selected ? theme.selectionOverlay : ""
                                                )}
                                                onMouseDown={(e) => handleMouseDown(rowIndex, colIndex, e)}
                                                onContextMenu={(e) => handleContextMenu(e, rowIndex, colIndex)}
                                                onMouseEnter={() => handleMouseEnter(rowIndex, colIndex)}
                                                onDoubleClick={() => {
                                                    if (isPickingTarget) onConfirmTarget?.(rowIndex, colIndex);
                                                    // Double click could trigger manual edit (F2) here if we wanted
                                                }}
                                                style={cellStyle}
                                            >
                                                {/* Selection Borders */}
                                                {selected && (
                                                    <div className="absolute inset-0 pointer-events-none z-10">
                                                        {edges.top && <div className={cn("absolute top-[-1px] left-[-1px] right-[-1px] h-[2px]", theme.selectionBorder)} />}
                                                        {edges.bottom && <div className={cn("absolute bottom-[-1px] left-[-1px] right-[-1px] h-[2px]", theme.selectionBorder)} />}
                                                        {edges.left && <div className={cn("absolute top-[-1px] left-[-1px] bottom-[-1px] w-[2px]", theme.selectionBorder)} />}
                                                        {edges.right && <div className={cn("absolute top-[-1px] right-[-1px] bottom-[-1px] w-[2px]", theme.selectionBorder)} />}
                                                        {edges.isCorner && !isPickingTarget && (
                                                            <div
                                                                onMouseDown={handleFillMouseDown}
                                                                onDragStart={(e) => e.preventDefault()}
                                                                className={cn("absolute bottom-[-4px] right-[-4px] w-2 h-2 z-20 cursor-crosshair shadow-sm pointer-events-auto", theme.selectionBorder)}
                                                                style={{ border: '1px solid white' }}
                                                            />
                                                        )}
                                                    </div>
                                                )}

                                                <div className="w-full h-full px-2 flex items-center overflow-hidden text-ellipsis whitespace-nowrap text-inherit cursor-cell select-none">
                                                    {cellValue !== undefined && cellValue !== null ? String(cellValue) : ""}
                                                </div>
                                            </td>
                                        );
                                    })}
                                </tr>
                            );
                        })}
                    </tbody>
                </table>
            </div>

            {/* Sheets Tabs (Simple version compared to logic used in previous files) */}
            {sheets && sheets.length > 0 && (
                <div className={cn("flex items-center px-4 bg-gray-100 dark:bg-slate-900 border-t gap-1 overflow-x-auto custom-scrollbar h-10 select-none", theme.borderColor)}>
                    {sheets.map(sheet => (
                        <div
                            key={sheet}
                            onClick={() => onChangeSheet && onChangeSheet(sheet)}
                            className={cn(
                                "relative group flex items-center px-4 h-full text-xs font-medium cursor-pointer transition-colors min-w-[80px] justify-center rounded-t-lg border-t border-l border-r -mb-[1px]",
                                currentSheet === sheet
                                    ? (isDarkMode ? "bg-slate-800 text-white border-slate-700" : "bg-white text-indigo-700 border-gray-300 shadow-sm")
                                    : (isDarkMode ? "text-slate-500 hover:bg-slate-800/50 border-transparent hover:border-slate-700" : "text-gray-500 hover:bg-gray-200 border-transparent hover:border-gray-300")
                            )}
                        >
                            <span>{sheet}</span>
                            {/* Delete Button */}
                            {sheets.length > 1 && currentSheet === sheet && (
                                <button className="ml-2 hover:text-red-500" onClick={(e) => { e.stopPropagation(); onDeleteSheet?.(sheet); }}>
                                    <X size={12} />
                                </button>
                            )}
                            {currentSheet === sheet && <div className="absolute bottom-0 left-0 right-0 h-[3px] bg-indigo-500 z-10" />}
                        </div>
                    ))}
                    <button onClick={onAddSheet} className="ml-2 w-6 h-6 flex items-center justify-center rounded-full hover:bg-gray-200 dark:hover:bg-slate-700">
                        <Plus size={16} />
                    </button>
                </div>
            )}

            {/* Context Menu Render */}
            {contextMenu && contextMenu.visible && (
                <div
                    className="fixed z-50 bg-white dark:bg-slate-800 rounded-lg shadow-xl border border-gray-200 dark:border-slate-700 py-1 min-w-[200px]"
                    style={{ top: contextMenu.y, left: contextMenu.x }}
                    onClick={(e) => e.stopPropagation()}
                >
                    <button className="w-full text-left px-3 py-1.5 hover:bg-gray-100 dark:hover:bg-slate-700 flex items-center gap-2 text-sm text-gray-700 dark:text-slate-300" onClick={() => { onInsertRow?.(contextMenu.r, 'above'); setContextMenu(null); }}>
                        <ArrowUp size={14} /> <span>위에 행 삽입</span>
                    </button>
                    <button className="w-full text-left px-3 py-1.5 hover:bg-gray-100 dark:hover:bg-slate-700 flex items-center gap-2 text-sm text-gray-700 dark:text-slate-300" onClick={() => { onInsertRow?.(contextMenu.r, 'below'); setContextMenu(null); }}>
                        <ArrowDown size={14} /> <span>아래에 행 삽입</span>
                    </button>
                    <button className="w-full text-left px-3 py-1.5 hover:bg-red-50 dark:hover:bg-red-900/20 text-red-600 dark:text-red-400 flex items-center gap-2 text-sm mb-1" onClick={() => { onDeleteRow?.(contextMenu.r); setContextMenu(null); }}>
                        <Trash size={14} /> <span>행 삭제</span>
                    </button>
                    <div className="border-b border-gray-100 dark:border-slate-700 my-1"></div>
                    <button className="w-full text-left px-3 py-1.5 hover:bg-gray-100 dark:hover:bg-slate-700 flex items-center gap-2 text-sm text-gray-700 dark:text-slate-300" onClick={() => { onInsertCol?.(contextMenu.c, 'left'); setContextMenu(null); }}>
                        <ArrowLeft size={14} /> <span>왼쪽에 열 삽입</span>
                    </button>
                    <button className="w-full text-left px-3 py-1.5 hover:bg-gray-100 dark:hover:bg-slate-700 flex items-center gap-2 text-sm text-gray-700 dark:text-slate-300" onClick={() => { onInsertCol?.(contextMenu.c, 'right'); setContextMenu(null); }}>
                        <ArrowRight size={14} /> <span>오른쪽에 열 삽입</span>
                    </button>
                    <button className="w-full text-left px-3 py-1.5 hover:bg-red-50 dark:hover:bg-red-900/20 text-red-600 dark:text-red-400 flex items-center gap-2 text-sm mb-1" onClick={() => { onDeleteCol?.(contextMenu.c); setContextMenu(null); }}>
                        <Trash size={14} /> <span>열 삭제</span>
                    </button>
                    <div className="border-b border-gray-100 dark:border-slate-700 my-1"></div>
                    <button className="w-full text-left px-3 py-1.5 hover:bg-gray-100 dark:hover:bg-slate-700 flex items-center gap-2 text-sm text-gray-700 dark:text-slate-300" onClick={() => { onClearContent?.(contextMenu.r, contextMenu.c); setContextMenu(null); }}>
                        <Eraser size={14} /> <span>내용 지우기</span>
                    </button>
                </div>
            )}
        </div>
    );
}
