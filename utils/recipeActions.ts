
import * as XLSX from "xlsx-js-style";

// Types
type GridData = any[][];
type Selection = { start: { r: number; c: number }; end: { r: number; c: number } };

// Helper: Deep Clone Data
const cloneData = (data: GridData): GridData => {
    return data.map(row => (row ? [...row] : [])); // Ensure row is array
};

// Helper: Ensure Cell Object
const ensureCell = (data: GridData, r: number, c: number) => {
    if (!data[r]) data[r] = [];
    const cell = data[r][c];
    if (cell && typeof cell === 'object') return { ...cell }; // Clone existing
    return { v: cell, t: typeof cell === 'number' ? 'n' : 's' }; // Create new
};

// --- Pure Mover ---
export const applyMove = (selection: Selection | null, direction: 'up' | 'down' | 'left' | 'right'): Selection | null => {
    if (!selection) return null;
    const { start, end } = selection;
    // Move both start and end by 1 step
    // Or just move the "active cell"? Usually user expects the whole selection box to move?
    // Let's assume the whole box moves.

    let dR = 0;
    let dC = 0;
    if (direction === 'up') dR = -1;
    if (direction === 'down') dR = 1;
    if (direction === 'left') dC = -1;
    if (direction === 'right') dC = 1;

    const newStart = { r: Math.max(0, start.r + dR), c: Math.max(0, start.c + dC) };
    const newEnd = { r: Math.max(0, end.r + dR), c: Math.max(0, end.c + dC) };

    return { start: newStart, end: newEnd };
};

// --- Pure Calculation (Row) ---
export const applyRowMath = (
    data: GridData,
    selection: Selection,
    operator: string,
    resultMode: 'overwrite' | 'new' | 'pick_execute' = 'overwrite'
): { newData: GridData, newSelection: Selection } => {
    const newData = cloneData(data);
    const { start, end } = selection;
    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);

    let targetC = c2; // Default Overwrite
    if (resultMode === 'new') targetC = c2 + 1; // Append to right
    // 'pick_execute' handled by caller usually, but if implied, use c2

    for (let r = r1; r <= r2; r++) {
        // Calculate Row Sum/Avg/etc for this row's selection
        let val = 0;
        let count = 0;
        let first = true;

        for (let c = c1; c <= c2; c++) {
            const cell = newData[r]?.[c];
            const raw = (cell && typeof cell === 'object') ? cell.v : cell;
            const num = parseFloat(String(raw));
            if (!isNaN(num)) {
                if (first) { val = num; first = false; }
                else {
                    if (operator === '+') val += num;
                    if (operator === '-') val -= num;
                    if (operator === '*') val *= num;
                    if (operator === '/') val = num !== 0 ? val / num : 0; // Avoid Div0
                    // For Row Math, typically it's Sum/Avg of the range
                }
                count++;
            }
        }

        // Aggregation Logic Override for basic operators if 'range' implied?
        // math_row_row usually means "Operate horizontally"? 
        // Wait, the prompt implies "Average(A1:C1) -> D1".

        // Logic Refinement:
        // If operator is 'avg', we divide by count at end
        if (operator === 'avg' && count > 0) val = val / count;
        if (first) val = 0; // No valid numbers

        // Write Result
        if (!newData[r]) newData[r] = [];

        // Preserve Style if exists or Copy from left?
        const prevCell = newData[r][targetC];
        const newCell = {
            v: val,
            t: 'n',
            s: (prevCell && typeof prevCell === 'object') ? prevCell.s : undefined
        };
        newData[r][targetC] = newCell;
    }

    // Update selection to cover the result?
    const newSelection = {
        start: { r: r1, c: targetC },
        end: { r: r2, c: targetC }
    };

    return { newData, newSelection };
};

// --- Pure Calculation (Column) ---
export const applyColMath = (
    data: GridData,
    selection: Selection,
    operator: string,
    resultMode: 'overwrite' | 'new' | 'pick_execute' = 'overwrite'
): { newData: GridData, newSelection: Selection } => {
    const newData = cloneData(data);
    const { start, end } = selection;
    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);

    let targetR = r2;
    if (resultMode === 'new') targetR = r2 + 1;

    const resultCells = [];

    for (let c = c1; c <= c2; c++) {
        let val = 0;
        let count = 0;
        let first = true;

        for (let r = r1; r <= r2; r++) {
            const cell = newData[r]?.[c];
            const raw = (cell && typeof cell === 'object') ? cell.v : cell;
            const num = parseFloat(String(raw));
            if (!isNaN(num)) {
                if (first) { val = num; first = false; }
                else {
                    if (operator === '+') val += num;
                    if (operator === '-') val -= num;
                    if (operator === '*') val *= num;
                    if (operator === '/') val = num !== 0 ? val / num : 0;
                }
                count++;
            }
        }

        if (operator === 'avg' && count > 0) val = val / count;
        if (first) val = 0;

        // Write Result
        if (!newData[targetR]) {
            // Fill implicit rows
            for (let i = newData.length; i <= targetR; i++) newData[i] = [];
        }
        if (!newData[targetR]) newData[targetR] = [];

        const prevCell = newData[targetR][c];
        const newCell = {
            v: val,
            t: 'n',
            s: (prevCell && typeof prevCell === 'object') ? prevCell.s : undefined
        };
        newData[targetR][c] = newCell;
    }

    const newSelection = {
        start: { r: targetR, c: c1 },
        end: { r: targetR, c: c2 }
    };

    return { newData, newSelection };
};

// --- Pure Style ---
export const applyStyleUtil = (
    data: GridData,
    selection: Selection,
    type: string,
    value?: any
): { newData: GridData, newSelection: Selection } => {
    const newData = cloneData(data);
    const { start, end } = selection;
    // Normalized Bounds
    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);

    for (let r = r1; r <= r2; r++) {
        if (!newData[r]) newData[r] = [];
        for (let c = c1; c <= c2; c++) {
            let cell = newData[r][c];
            // Ensure object
            if (cell === null || cell === undefined || typeof cell !== 'object') {
                cell = { v: cell, t: typeof cell === 'number' ? 'n' : 's' };
            } else {
                cell = { ...cell }; // Clone
            }

            // Init Style
            if (!cell.s) cell.s = {};
            if (!cell.s.font) cell.s.font = {};
            if (!cell.s.alignment) cell.s.alignment = {};
            if (!cell.s.fill) cell.s.fill = {};

            // Apply
            if (type === 'bold') cell.s.font.bold = !cell.s.font.bold;
            if (type === 'italic') cell.s.font.italic = !cell.s.font.italic;
            if (type === 'underline') cell.s.font.underline = !cell.s.font.underline;
            if (type === 'align') cell.s.alignment.horizontal = value;
            if (type === 'color') cell.s.font.color = { rgb: value };
            if (type === 'fill') cell.s.fill.fgColor = { rgb: value };

            newData[r][c] = cell;
        }
    }
    return { newData, newSelection: selection };
};

// --- Pure Calculation (Math Constant) ---
export const applyColConst = (
    data: GridData,
    selection: Selection,
    operator: string,
    value: number,
    resultMode: 'overwrite' | 'new' | 'pick_execute' = 'overwrite'
): { newData: GridData, newSelection: Selection } => {
    const newData = cloneData(data);
    const { start, end } = selection;
    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);

    // Similar to Col Math but second operand is constant
    let targetR = r2;
    if (resultMode === 'new') targetR = r2 + 1;

    for (let c = c1; c <= c2; c++) {
        for (let r = r1; r <= r2; r++) {
            const cell = newData[r]?.[c];
            const raw = (cell && typeof cell === 'object') ? cell.v : cell;
            const num = parseFloat(String(raw));
            let resVal = num;

            if (!isNaN(num)) {
                if (operator === '+') resVal += value;
                if (operator === '-') resVal -= value;
                if (operator === '*') resVal *= value;
                if (operator === '/') resVal = value !== 0 ? resVal / value : 0;
            } else {
                resVal = 0; // Or keep original? Typically math on text is 0 or NaN
            }

            // Write Result
            // If overwrite, write to same cell
            let writeR = r;
            let writeC = c;
            if (resultMode === 'new') {
                // For Col Const, 'new' might mean append row? 
                // Logic follows Col Math pattern: Result goes to targetR (bottom + 1)
                writeR = targetR;
            }

            if (!newData[writeR]) {
                for (let i = newData.length; i <= writeR; i++) newData[i] = [];
            }

            const prevCell = newData[writeR][writeC];
            newData[writeR][writeC] = {
                v: resVal,
                t: 'n',
                s: (prevCell && typeof prevCell === 'object') ? prevCell.s : undefined
            };
        }
    }

    // Select result area if new
    const newSel = resultMode === 'new'
        ? { start: { r: targetR, c: c1 }, end: { r: targetR, c: c2 } }
        : selection;

    return { newData, newSelection: newSel };
};

// --- Pure Logic IF ---
export const applyLogicIf = (
    data: GridData,
    selection: Selection,
    condition: { operator: string, value: string | number },
    trueVal: string,
    falseVal: string
): { newData: GridData, newSelection: Selection } => {
    const newData = cloneData(data);
    const { start, end } = selection;
    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);

    for (let r = r1; r <= r2; r++) {
        if (!newData[r]) newData[r] = [];
        for (let c = c1; c <= c2; c++) {
            const cell = newData[r][c];
            const raw = (cell && typeof cell === 'object') ? cell.v : cell;
            const num = parseFloat(String(raw)); // Try parse number

            let isMatch = false;
            const compVal = parseFloat(String(condition.value));

            if (!isNaN(num) && !isNaN(compVal)) {
                if (condition.operator === '>') isMatch = num > compVal;
                if (condition.operator === '>=') isMatch = num >= compVal;
                if (condition.operator === '<') isMatch = num < compVal;
                if (condition.operator === '<=') isMatch = num <= compVal;
                if (condition.operator === '=') isMatch = num == compVal;
            } else {
                // String comparison
                if (condition.operator === '=') isMatch = String(raw) == String(condition.value);
            }

            if (isMatch && trueVal) {
                const prevCell = newData[r][c];
                newData[r][c] = { v: trueVal, t: 's', s: (prevCell && typeof prevCell === 'object') ? prevCell.s : undefined };
            } else if (!isMatch && falseVal) {
                const prevCell = newData[r][c];
                newData[r][c] = { v: falseVal, t: 's', s: (prevCell && typeof prevCell === 'object') ? prevCell.s : undefined };
            }
        }
    }
    return { newData, newSelection: selection };
};

// --- Pure Calculation (Range Aggregation - Scalar) ---
export const applyRangeMath = (
    data: GridData,
    selection: Selection,
    operator: string
    // resultMode handling can be done by caller or here. Let's do here for consistency.
): { newData: GridData, newSelection: Selection } => {
    const newData = cloneData(data);
    const { start, end } = selection;
    const r1 = Math.min(start.r, end.r);
    const r2 = Math.max(start.r, end.r);
    const c1 = Math.min(start.c, end.c);
    const c2 = Math.max(start.c, end.c);

    // Collect Values
    let values: number[] = [];
    for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
            const cell = newData[r]?.[c];
            const raw = (cell && typeof cell === 'object') ? cell.v : cell;
            const num = parseFloat(String(raw));
            if (!isNaN(num)) values.push(num);
        }
    }

    let val = 0;
    if (operator === 'sum' || operator === '+') val = values.reduce((a, b) => a + b, 0);
    else if (operator === 'avg') val = values.reduce((a, b) => a + b, 0) / (values.length || 1);
    else if (operator === 'count') val = values.length;
    else if (operator === 'max') val = Math.max(...values);
    else if (operator === 'min') val = Math.min(...values);

    // Target Placement: Right of selection
    const targetR = r2;
    const targetC = c2 + 1;

    // Check bounds
    if (!newData[targetR]) {
        for (let i = newData.length; i <= targetR; i++) newData[i] = [];
    }
    if (!newData[targetR]) newData[targetR] = [];

    const prevCell = newData[targetR][targetC];
    newData[targetR][targetC] = {
        v: val,
        t: 'n',
        s: (prevCell && typeof prevCell === 'object') ? prevCell.s : undefined
    };

    const newSelection = {
        start: { r: targetR, c: targetC },
        end: { r: targetR, c: targetC }
    };

    return { newData, newSelection };
};
