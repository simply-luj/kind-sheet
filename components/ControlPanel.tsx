"use client";

import React, { useState, useEffect } from "react";
import { Eraser, Calculator, Paintbrush, Sparkles, Trash2, Rows, Download, AlignLeft, AlignCenter, AlignRight, FolderOpen, Save, Hash, Type, Copy, Sun, Moon, Plus, ChefHat, Play, Star, X, Wrench, Split, Combine, Scissors, Quote, BarChart3, LineChart, ScatterChart, Sigma, Database, ArrowDownAZ, ArrowUpAZ, Filter, Replace as ReplaceIcon, RefreshCw, EyeOff, Divide, Percent, X as MultiplyIcon, Minus, Bold, Italic, Underline, Palette, PaintBucket, Scale, Search, ChevronDown, Zap, LayoutGrid, ArrowLeft, ArrowUp, ArrowRight, ArrowDown, PlusCircle } from "lucide-react";
import { clsx } from "clsx";
import { twMerge } from "tailwind-merge";

function cn(...inputs: (string | undefined | null | false)[]) {
    return twMerge(clsx(inputs));
}

// Reuse KeycapButton
interface KeycapProps {
    shortcut: string;
    label: string;
    onClick: () => void;
    isActive?: boolean;
    disabled?: boolean;
    icon?: React.ElementType;
}

function KeycapButton({ shortcut, label, onClick, isActive, disabled, icon: Icon }: KeycapProps) {
    return (
        <button
            onClick={onClick}
            disabled={disabled}
            className={cn(
                "relative group flex flex-col items-center justify-center w-14 h-14 rounded-xl transition-all duration-100",
                "border-b-4 active:border-b-0 active:translate-y-1",
                isActive
                    ? "bg-indigo-600 border-indigo-800 text-white shadow-lg shadow-indigo-900/50 translate-y-1 border-b-0"
                    : "bg-slate-700 border-slate-900 text-slate-300 hover:bg-slate-600 hover:text-white hover:border-slate-800",
                disabled && "opacity-50 cursor-not-allowed border-b-0 translate-y-1 bg-slate-800 text-slate-600"
            )}
            title={label}
        >
            {Icon ? <Icon size={24} className="mb-1" /> : <span className="text-xs font-bold font-mono uppercase">{shortcut}</span>}
            <span className="absolute -bottom-6 text-[10px] font-sans text-slate-500 opacity-0 group-hover:opacity-100 transition-opacity whitespace-nowrap">
                {label}
            </span>
        </button>
    );
}

interface ControlPanelProps {
    onAction: (action: string | any) => void;
    onExport: () => void;
    onOpen: () => void;
    hasData: boolean;
    alignment: "left" | "center" | "right" | "auto";
    onAlignmentChange: (align: "left" | "center" | "right" | "auto") => void;
    canUndo: boolean;
    canRedo: boolean;
    selection?: { start: { r: number; c: number }; end: { r: number; c: number } } | null;
    isDarkMode: boolean;
    onToggleTheme: () => void;
    className?: string;
    onSave: () => void;
}

// Consolidated Types


type CategoryId = "calc" | "text" | "clean" | "style" | "analyze" | "data" | "logic" | "my_recipe";

// --- Type Definitions ---
type RecipeAction = {
    type: string;
    payload?: any;
    desc: string;
};

// 1. Standard Single-Action Recipe (Legacy)
type StandardRecipe = {
    id: string;
    name: string;
    category: CategoryId;
    type: string; // e.g., 'math_row_row'
    option?: any;
    keepOriginal?: boolean;
    filter?: any;
    format?: boolean;
    queue?: never; // Discriminated union helper
};

// 2. Custom Builder Recipe (New Queue-based)
type BuilderRecipe = {
    id: string;
    name: string;
    category: 'my_recipe'; // Always 'my_recipe'
    queue: RecipeAction[];
    type?: never;
};

// Unified Type
type AnyRecipe = StandardRecipe | BuilderRecipe;

// Alias for legacy code compatibility
type Recipe = StandardRecipe;
// Note: We will cast state to AnyRecipe[]


// Note: We will cast state to AnyRecipe[]


export default function ControlPanel({ onAction, onExport, onOpen, onSave, hasData, alignment, onAlignmentChange, canUndo, canRedo, selection, isDarkMode, onToggleTheme, className }: ControlPanelProps) {
    const [activeCategory, setActiveCategory] = useState<CategoryId>("calc");

    // --- Recipe Builder State ---
    // --- Recipe Builder State ---
    const [isBuildingMode, setIsBuildingMode] = useState(false);
    const [recipeQueue, setRecipeQueue] = useState<RecipeAction[]>([]);

    // Unified Saved Recipes
    const [savedRecipes, setSavedRecipes] = useState<AnyRecipe[]>([]);

    // UI State for Recipe Names
    const [recipeName, setRecipeName] = useState(""); // Legacy? (Keeping for safety if referenced elsewhere)
    const [newRecipeName, setNewRecipeName] = useState(""); // For new Recipe Builder UI


    // Load recipes on mount
    useEffect(() => {
        const saved = localStorage.getItem("kind_sheet_recipes");
        if (saved) {
            try {
                setSavedRecipes(JSON.parse(saved));
            } catch (e) {
                console.error("Failed to load recipes", e);
            }
        }
    }, []);

    // Save recipes on change
    useEffect(() => {
        if (savedRecipes.length > 0) {
            localStorage.setItem("kind_sheet_recipes", JSON.stringify(savedRecipes));
        }
    }, [savedRecipes]);

    // Action Interceptor
    // If Building Mode is ON, we don't fire onAction. We push to queue.
    const handleDispatch = (type: string, payload?: any, desc?: string) => {
        if (isBuildingMode) {
            // Generate Description if missing
            let description = desc || type;
            if (type === 'style') description = `스타일: ${payload.type}`;
            if (type === 'math_row_row') description = `행 계산 (${payload.option?.operator || '+'})`;

            const action: RecipeAction = { type, payload, desc: description };
            setRecipeQueue(prev => [...prev, action]);
            return; // STOP EXECUTION
        }

        // Normal Execution
        if (type === 'run_recipe') {
            // Special payload for running recipe
            onAction({ type: 'run_recipe', recipe: payload });
            return;
        }

        // Standard Action Pass-through
        if (payload) {
            onAction(payload); // If payload is the full action object (legacy pattern)
        } else {
            onAction(type);
        }
    };

    // Helper to wrap legacy handlers
    const dispatchWrapper = (actionObj: any, desc: string) => {
        if (isBuildingMode) {
            // Flatten payload logic:
            // If actionObj is { type: 'x', payload: y }, we want queue item { type: 'x', payload: y }
            // If actionObj is simple string, type=string
            let type = actionObj.type || (typeof actionObj === 'string' ? actionObj : 'unknown');
            let payload = actionObj.payload !== undefined ? actionObj.payload : (actionObj.option || actionObj);

            // Special handling for legacy construct
            if (typeof actionObj === 'object' && !actionObj.payload && !actionObj.option && actionObj.category) {
                // It's likely the full recipePayload from saveRecipe/executeCurrent?
                // e.g. { category: 'calc', type: 'math_row_row', option: ... }
                // We want the queue item to preserve this structure or normalize it.
                // Let's use the object itself as payload if it's complex 
                // BUT handleRunRecipe expects specific payload structure.
                // Let's trust the actionObj IS the payload for 'handleRunRecipe' dispatch?
                // No, handleRunRecipe expects { type, payload }.
                // If actionObj is { type: 'math_row_row', option: ... }
                // Then type='math_row_row', payload= { option: ... }? Or payload = actionObj?

                // Let's stick to: Payload IS the config object.
                payload = actionObj;
            }

            setRecipeQueue(prev => [...prev, { type, payload, desc }]);
        } else {
            onAction(actionObj);
        }
    };

    // [Calc]
    const [calcMain, setCalcMain] = useState<'sum' | 'count' | 'avg' | 'math_col_col' | 'math_col_const' | 'math_row_row'>('sum');
    const [calcLogic, setCalcLogic] = useState({ filter: false, operator: '>', value: '', format: false });
    const [calcResultMode, setCalcResultMode] = useState<'overwrite' | 'new' | 'pick'>('overwrite');

    // [Text]
    const [textMain, setTextMain] = useState<'join' | 'split' | 'extract'>('join');
    const [textOption, setTextOption] = useState({ delimiter: ',', count: 1, mode: 'left' as 'left' | 'right' });
    const [textKeep, setTextKeep] = useState(false);

    // [Clean]
    const [cleanMain, setCleanMain] = useState<'clean_empty' | 'remove_dup' | 'trim'>('clean_empty');
    const [cleanOption, setCleanOption] = useState({ allSpaces: false });

    // [Style]
    const [styleMain, setStyleMain] = useState<'comma' | 'header_style' | 'highlight'>('comma');
    const [styleLogic, setStyleLogic] = useState({ operator: '>', value: '', color: 'yellow' });

    // [Analyze]
    const [analyzeMain, setAnalyzeMain] = useState<'stat_basic' | 'chart_bar' | 'chart_line' | 'chart_scatter'>('stat_basic');
    const [analyzeOption, setAnalyzeOption] = useState({ label: false });

    // [Data]
    // [Data]
    const [dataMain, setDataMain] = useState<'sort_asc' | 'sort_desc' | 'filter' | 'replace'>('sort_asc');
    const [dataOption, setDataOption] = useState({ header: true, condition: '', find: '', replace: '' });


    // [Logic] - New Tab State
    // [Logic] - New Tab State
    const [logicMain, setLogicMain] = useState<'if' | 'vlookup'>('if');
    const [logicIf, setLogicIf] = useState({ operator: '>', value: '', trueVal: '', falseVal: '' });
    const [logicVlookup, setLogicVlookup] = useState({ range: '', colIndex: 2 });

    // --- Legacy State Cleanup (Removed duplicate savedRecipes) ---
    // const [savedRecipes, setSavedRecipes] ... REMOVED
    const [isNamingRecipe, setIsNamingRecipe] = useState(false);
    // const [newRecipeName, setNewRecipeName] ... REMOVED (using recipeName)
    const [hoveredDescription, setHoveredDescription] = useState<string | null>(null);
    const [toastMsg, setToastMsg] = useState<string | null>(null);

    // --- ACCORDION STATE ---
    const [expandedSections, setExpandedSections] = useState({
        builder: false,
        toolkit: true,
        functions: true
    });

    const toggleSection = (section: 'builder' | 'toolkit' | 'functions') => {
        setExpandedSections(prev => ({ ...prev, [section]: !prev[section] }));
    };

    // Auto-Collapse Logic on Builder Mode Change
    useEffect(() => {
        if (isBuildingMode) {
            setExpandedSections({ builder: true, toolkit: false, functions: true });
        } else {
            // Restore default
            setExpandedSections(prev => ({ ...prev, builder: false, toolkit: true }));
        }
    }, [isBuildingMode]);

    const DESCRIPTIONS: Record<string, string> = {
        calc_sum: "선택한 범위의 숫자를 모두 더합니다. (총합)",
        calc_count: "선택한 범위에 값이 몇 개 있는지 셉니다. (인원수/재고)",
        calc_avg: "선택한 범위의 평균값을 계산합니다.",
        math_col_col: "두 열을 선택하면, 왼쪽 값과 오른쪽 값을 계산하여 오른쪽에 덮어씁니다. (A+B)",
        math_col_const: "선택한 모든 셀에 특정 숫자를 더하거나 곱합니다. (A+N)",
        math_row_row: "선택한 범위의 위쪽 행과 아래쪽 행을 계산합니다. (1행+2행)",
        opt_result_overwrite: "결과를 마지막 칸에 덮어씁니다. (기존 데이터 삭제)",
        opt_result_new: "결과를 새로운 칸에 추가합니다. (원본 보존)",
        opt_result_pick: "결과를 출력할 위치를 직접 클릭하여 지정합니다.",
        option_if: "특정 조건(예: 100 이상)에 맞는 값만 골라서 계산합니다.",
        text_join: "여러 칸의 글자를 하나로 합칩니다. (성+이름)",
        text_split: "한 칸의 글자를 여러 칸으로 쪼갭니다. (주소 분리)",
        text_extract: "글자의 앞이나 뒤에서 원하는 만큼만 가져옵니다.",
        option_glue: "합칠 때 사이에 공백이나 쉼표를 넣어줍니다.",
        clean_empty: "데이터가 없는 빈 줄을 찾아 삭제합니다.",
        remove_dup: "완전히 똑같은 중복 데이터를 하나만 남깁니다.",
        trim: "글자 앞뒤의 불필요한 공백을 제거합니다.",
        opt_all_spaces: "체크하면 글자 사이의 모든 띄어쓰기를 없앱니다.",
        comma: "숫자에 천 단위 콤마(,)를 찍어 보기 좋게 만듭니다.",
        header_style: "첫 번째 줄을 제목처럼 진하게 강조합니다.",
        style_highlight: "원하는 조건에 맞는 칸만 자동으로 색칠해 줍니다.",
        save_recipe: "현재 설정한 기능을 '나만의 버튼'으로 저장합니다.",
        stat_basic: "평균, 중앙값, 최솟값, 최댓값 등 기초 통계를 확인합니다.",
        chart_bar: "막대 그래프로 데이터의 크기를 비교합니다.",
        chart_line: "데이터의 추세를 선 그래프로 확인합니다.",
        chart_scatter: "두 데이터 간의 상관관계와 회귀선(추세선)을 분석합니다.",
        opt_label: "첫 번째 열을 이름(라벨)으로 사용하여 그래프를 그립니다.",
        open_file: "컴퓨터에 있는 엑셀 파일을 불러옵니다. (현재 내용 덮어쓰기)",
        download_file: "현재 작업 내용을 엑셀 파일로 다운로드합니다.",
        data_sort_asc: "선택한 열을 기준으로 데이터를 오름차순(가나다순)으로 정렬합니다.",
        data_sort_desc: "선택한 열을 기준으로 데이터를 내림차순(역순)으로 정렬합니다.",
        data_filter: "조건에 맞는 데이터만 남기고 나머지는 제거합니다. (추출)",
        data_replace: "특정 글자를 찾아 다른 글자로 일괄 변경합니다.",
        opt_header: "첫 번째 줄(제목)은 정렬이나 필터에서 제외합니다.",
        opt_condition: "공백 없이 조건 입력 (예: >50, 서울, *김*)",
        opt_find: "찾을 내용을 입력하세요.",
        opt_replace: "바꿀 내용을 입력하세요.",
        logic_if: "조건에 따라 다른 값을 표시합니다. (예: 60점 이상이면 '합격')",
        logic_vlookup: "다른 표에서 원하는 값을 찾아옵니다. (예: 상품명으로 가격 찾기)",
        logic_operator: "비교할 조건(크다, 작다, 같다)을 선택하세요.",
        logic_range: "값을 찾을 범위를 입력하세요. (예: Sheet2!A:B)",
        logic_col: "범위에서 몇 번째 열의 값을 가져올지 입력하세요."
    };

    useEffect(() => {
        const saved = localStorage.getItem('kind_sheet_recipes'); // Unified Key
        if (saved) {
            try {
                setSavedRecipes(JSON.parse(saved));
            } catch (e) { console.error(e); }
        }
    }, []);

    // OLD Legacy Save Logic (Standard/Single Action)
    // This is used by the "Save" button in the *Tabs* (Bottom Action Bar), not the Top Builder.
    const saveRecipe = () => {
        if (!newRecipeName.trim()) return;

        let recipePayload: any = {
            id: Date.now().toString(),
            name: newRecipeName,
            category: activeCategory,
        };

        if (activeCategory === 'calc') {
            if (calcMain.startsWith('math')) {
                recipePayload = { ...recipePayload, type: calcMain, option: { operator: calcLogic.operator, value: calcLogic.value, resultMode: calcResultMode } };
            } else {
                recipePayload = { ...recipePayload, type: calcMain, filter: calcLogic.filter ? { operator: calcLogic.operator, value: calcLogic.value } : undefined, format: calcLogic.format };
            }
        } else if (activeCategory === 'text') {
            recipePayload = {
                ...recipePayload,
                type: textMain,
                option: { delimiter: textOption.delimiter, count: textOption.count, mode: textOption.mode },
                keepOriginal: textKeep
            };
        } else if (activeCategory === 'analyze') {
            recipePayload = { ...recipePayload, type: analyzeMain, option: analyzeOption };
        } else if (activeCategory === 'data') {
            recipePayload = { ...recipePayload, type: dataMain, option: dataOption };
        } else if (activeCategory === 'logic') {
            if (logicMain === 'if') recipePayload = { ...recipePayload, type: 'logic_if', option: logicIf };
            else recipePayload = { ...recipePayload, type: 'logic_vlookup', option: logicVlookup };
        }

        const updated = [...savedRecipes, recipePayload as AnyRecipe];
        setSavedRecipes(updated);
        // Save to SAME storage key
        localStorage.setItem('kind_sheet_recipes', JSON.stringify(updated));
        setIsNamingRecipe(false);
        setNewRecipeName("");

        setToastMsg(`✅ '${newRecipeName}' 저장 완료`);
        setTimeout(() => setToastMsg(null), 2000);
    };

    const deleteRecipe = (id: string, e: React.MouseEvent) => {
        e.stopPropagation();
        const updated = savedRecipes.filter(r => r.id !== id);
        setSavedRecipes(updated);
        localStorage.setItem('kind_sheet_recipes', JSON.stringify(updated));
    };

    const handleExecuteFavorite = (recipe: AnyRecipe) => {
        setToastMsg(`⚡ '${recipe.name}' 실행 중...`);
        // Slight delay to show the "Executing" state
        setTimeout(() => {
            if (recipe.category === 'my_recipe' && recipe.queue) {
                // Builder Recipe
                handleDispatch('run_recipe', recipe);
            } else {
                // Standard Recipe (Legacy)
                // Need to reconstruct the action object? 
                // Currently 'recipe' IS the payload for single actions usually.
                // But handleDispatch handles 'run_recipe' wrapper.
                // Let's just pass it to onAction directly if it's standard.
                // Actually, the original design might have expected specific structure.
                // Let's assume onAction can handle it or we re-map.
                // For safety:
                onAction(recipe);
            }
            setToastMsg(null);
        }, 500);
    };

    // Unified Builder State
    const [formulaState, setFormulaState] = useState({
        start: '',
        end: '',
        connector: '~', // '~' (Range) or ',' (And)
        operation: 'sum',
        resultMode: 'overwrite',
        target: '',
        isTargetLocked: false
    });

    // Helper: Index to A1 (Local)
    const toA1 = (r: number, c: number) => {
        let label = "";
        let i = c;
        while (i >= 0) {
            label = String.fromCharCode((i % 26) + 65) + label;
            i = Math.floor(i / 26) - 1;
        }
        return `${label}${r + 1}`;
    };

    // Sync Selection to Formula Builder
    useEffect(() => {
        if (selection && activeCategory === 'calc') {
            const startA1 = toA1(selection.start.r, selection.start.c);
            const endA1 = toA1(selection.end.r, selection.end.c);

            // If Pick Mode AND Target is NOT Locked, sync to Target
            if (formulaState.resultMode === 'pick' && !formulaState.isTargetLocked) {
                setFormulaState(prev => ({ ...prev, target: startA1 }));
            } else {
                // Normal Mode OR Target Locked -> sync to Start/End (Input Range)
                if (startA1 === endA1) {
                    setFormulaState(prev => ({ ...prev, start: startA1, end: '' }));
                } else {
                    setFormulaState(prev => ({ ...prev, start: startA1, end: endA1, connector: '~' }));
                }
            }
        }
    }, [selection, activeCategory, formulaState.resultMode, formulaState.isTargetLocked]);

    const getCurrentPayload = () => {
        let payload: any = { category: activeCategory };
        if (activeCategory === 'calc') {
            payload = { type: 'unified_calc', payload: formulaState };
        } else if (activeCategory === 'text') {
            payload = {
                type: textMain,
                option: { delimiter: textOption.delimiter, count: textOption.count, mode: textOption.mode },
                keepOriginal: textKeep
            };
        } else if (activeCategory === 'clean') {
            payload = { type: cleanMain, option: cleanOption };
        } else if (activeCategory === 'style') {
            payload = { type: styleMain, option: styleLogic };
        } else if (activeCategory === 'analyze') {
            payload = { type: analyzeMain, option: analyzeOption };
        } else if (activeCategory === 'data') {
            payload = { type: dataMain, option: dataOption };
        } else if (activeCategory === 'logic') {
            if (logicMain === 'if') payload = { type: 'logic_if', payload: logicIf };
            else payload = { type: 'logic_vlookup', payload: logicVlookup };
        }
        return payload;
    };

    const executeCurrent = () => {
        const payload = getCurrentPayload();
        dispatchWrapper(payload, '현재 설정 실행');
    };

    // New Handlers for Split Button
    const handleAddStep = () => {
        const payload = getCurrentPayload();
        // Manually dispatch to queue (Bypassing dispatchWrapper's payload normalization if needed, but dispatchWrapper handles it well)
        // Let's use dispatchWrapper which already handles queueing if isBuildingMode=true
        dispatchWrapper(payload, '현재 설정 추가');
        setToastMsg("✅ 레코드에 추가됨");
        setTimeout(() => setToastMsg(null), 1000);
    };

    const handleTestRun = () => {
        const payload = getCurrentPayload();
        // Force execution by bypassing dispatchWrapper check or calling onAction directly
        onAction(payload);
    };

    const categories: { id: CategoryId, label: string, icon: any }[] = [
        { id: 'calc', label: '계산 (Calc)', icon: Calculator },
        { id: 'text', label: '글자 (Text)', icon: Type },
        { id: 'clean', label: '청소 (Clean)', icon: Eraser },
        { id: 'style', label: '서식 (Style)', icon: Paintbrush },
        { id: 'analyze', label: '분석 (Analyze)', icon: BarChart3 },
        { id: 'data', label: '데이터 (Data)', icon: Database },
        { id: 'logic', label: '논리 (Logic)', icon: Scale },
    ];

    return (
        <div className={cn("h-full w-full bg-slate-800 rounded-3xl overflow-hidden flex flex-col shadow-2xl shadow-slate-950/50 border border-slate-700/30 relative", className)}>

            {/* [A] FIXED ZONE (Header + Builder + Toolkit + Function Tabs) */}
            <div className="flex-none bg-slate-800 z-10 shadow-md">
                {/* 1. Header & File Actions */}
                <div className="p-6 pb-2 border-b border-slate-700/50">
                    <div className="flex justify-between items-center mb-4">
                        <div className="flex items-center gap-2 text-indigo-400">
                            <Sparkles size={20} />
                            <span className="text-sm font-bold uppercase tracking-wider">Kind Recipe</span>
                        </div>
                        <div className="flex gap-4 items-center">
                            <button onClick={onToggleTheme} className={cn("flex items-center justify-center w-10 h-10 rounded-xl transition-all", isDarkMode ? "bg-slate-700 text-yellow-400" : "bg-indigo-100 text-indigo-600")}>
                                {isDarkMode ? <Sun size={20} /> : <Moon size={20} />}
                            </button>
                            <div className="flex gap-2">
                                <button
                                    onClick={onOpen}
                                    onMouseEnter={() => setHoveredDescription(DESCRIPTIONS['open_file'])}
                                    onMouseLeave={() => setHoveredDescription(null)}
                                    className="flex items-center justify-center w-14 h-14 rounded-xl bg-slate-700 border-b-4 border-slate-900 text-slate-300 hover:bg-slate-600 hover:text-white hover:border-slate-800 transition-all active:border-b-0 active:translate-y-1"
                                    title="파일 열기"
                                >
                                    <FolderOpen size={24} />
                                </button>
                                <button
                                    onClick={onExport}
                                    onMouseEnter={() => setHoveredDescription(DESCRIPTIONS['download_file'])}
                                    onMouseLeave={() => setHoveredDescription(null)}
                                    className="flex items-center justify-center w-14 h-14 rounded-xl bg-slate-700 border-b-4 border-slate-900 text-slate-300 hover:bg-slate-600 hover:text-white hover:border-slate-800 transition-all active:border-b-0 active:translate-y-1"
                                    title="파일 저장/다운로드"
                                >
                                    <Download size={24} />
                                </button>
                            </div>
                        </div>
                    </div>

                    {/* Builder Accordion (Top) - HIDDEN ON MOBILE LITE - BETA: HIDDEN GLOBALLY */}
                    {/*
                    <div className={cn("mb-2 rounded-xl border transition-all overflow-hidden hidden lg:block", isBuildingMode ? "bg-indigo-900/40 border-indigo-500 shadow-lg" : "bg-slate-800/50 border-slate-700")}>
                        <div
                            className="p-3 flex items-center justify-between cursor-pointer hover:bg-white/5 transition-colors"
                            onClick={() => toggleSection('builder')}
                        >
                            <div className="flex items-center gap-2">
                                <div className={cn("w-2 h-2 rounded-full", isBuildingMode ? "bg-red-500 animate-pulse" : "bg-slate-600")}></div>
                                <h3 className={cn("text-xs font-bold uppercase tracking-wider", isBuildingMode ? "text-white" : "text-slate-400")}>Recipe Builder</h3>
                            </div>
                            <div className="flex gap-2">
                                {expandedSections.builder && !isBuildingMode && (
                                    <button
                                        onClick={(e) => { e.stopPropagation(); setIsBuildingMode(true); setRecipeQueue([]); }}
                                        className="px-3 py-1 bg-indigo-600 hover:bg-indigo-500 text-white text-[10px] font-bold rounded-lg"
                                    >
                                        START
                                    </button>
                                )}
                                {expandedSections.builder && isBuildingMode && (
                                    <>
                                        <button onClick={(e) => { e.stopPropagation(); setIsBuildingMode(false); setRecipeQueue([]); }} className="px-3 py-1 bg-slate-700 text-slate-300 text-[10px] font-bold rounded-lg">CANCEL</button>
                                        <button onClick={(e) => { e.stopPropagation(); setIsNamingRecipe(true); }} disabled={recipeQueue.length === 0} className="px-3 py-1 bg-green-600 text-white text-[10px] font-bold rounded-lg disabled:opacity-50">SAVE</button>
                                    </>
                                )}
                                <ChevronDown size={14} className={cn("text-slate-500 transition-transform", expandedSections.builder ? "rotate-180" : "")} />
                            </div>
                        </div>
                        {expandedSections.builder && (
                            <div className="px-3 pb-3 animate-in slide-in-from-top-2">
                                {isBuildingMode && (
                                    <div className="min-h-[60px] max-h-[120px] overflow-y-auto bg-slate-900/50 rounded-lg p-2 border border-slate-700/50 custom-scrollbar gap-1 flex flex-col">
                                        {recipeQueue.length === 0 ? (
                                            <p className="text-[10px] text-slate-500 text-center py-2">기능을 추가하세요.</p>
                                        ) : (
                                            recipeQueue.map((step, idx) => (
                                                <div key={idx} className="flex items-center gap-2 px-2 py-1 bg-slate-800 rounded border border-slate-700 text-[10px] text-slate-300">
                                                    <span className="w-4 h-4 rounded-full bg-slate-700 flex items-center justify-center font-bold text-slate-500">{idx + 1}</span>
                                                    <span className="truncate flex-1">{step.desc}</span>
                                                    <button onClick={() => setRecipeQueue(q => q.filter((_, i) => i !== idx))} className="hover:text-red-400 p-0.5"><X size={10} /></button>
                                                </div>
                                            ))
                                        )}
                                    </div>
                                )}
                                {!isBuildingMode && (
                                    <div className="text-center py-2">
                                        <button onClick={() => { setIsBuildingMode(true); setRecipeQueue([]); }} className="text-xs text-indigo-400 hover:text-indigo-300 font-bold underline">레시피 만들기 시작</button>
                                    </div>
                                )}
                            </div>
                        )}
                    </div>
                     */}
                </div>

                {/* Toolkit Button (Compact) - HIDDEN ON MOBILE LITE - BETA: HIDDEN GLOBALLY */}
                {/* 
                <div className="px-6 py-2 border-b border-slate-700 flex items-center justify-between cursor-pointer hover:bg-slate-700/30 hidden lg:flex" onClick={() => toggleSection('toolkit')}>
                    <h3 className="text-xs font-bold text-indigo-400 uppercase tracking-wider flex items-center gap-2"><Zap size={14} /> My Toolkit</h3>
                    <ChevronDown size={14} className={cn("text-slate-500 transition-transform", expandedSections.toolkit ? "rotate-180" : "")} />
                </div>
                {expandedSections.toolkit && (
                    <div className="px-6 py-2 bg-slate-800/50 border-b border-slate-700 overflow-x-auto gap-2 scrollbar-none snap-x h-14 items-center hidden lg:flex">
                        {savedRecipes.length === 0 ? <span className="text-xs text-slate-500">저장된 레시피 없음</span> : savedRecipes.map(r => (
                            <button key={r.id} onClick={() => handleExecuteFavorite(r)} className="snap-start px-3 py-1 bg-indigo-600 text-white rounded text-xs font-bold whitespace-nowrap hover:bg-indigo-500 shadow-sm">{r.name}</button>
                        ))}
                    </div>
                )}
                 */}

                {/* 3. FUNCTION TABS (CRITICAL: FIXED POSITION) */}
                <div className="bg-slate-900 border-b border-slate-700">
                    <div className="px-6 py-2 flex items-center justify-between">
                        <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider flex items-center gap-2"><LayoutGrid size={14} /> Functions</h3>
                    </div>
                    <div className="px-4 pb-0 flex items-center gap-1 overflow-x-auto scrollbar-none">
                        {categories.map(cat => (
                            <button
                                key={cat.id}
                                onClick={() => setActiveCategory(cat.id)}
                                className={cn(
                                    "flex items-center gap-1.5 px-4 py-3 text-xs font-bold transition-all whitespace-nowrap border-b-2",
                                    activeCategory === cat.id
                                        ? "border-indigo-500 text-white bg-indigo-500/10"
                                        : "border-transparent text-slate-500 hover:text-slate-300 hover:bg-white/5"
                                )}
                            >
                                <cat.icon size={14} />
                                {cat.label}
                            </button>
                        ))}
                    </div>
                </div>
            </div>

            {/* [B] SCROLL ZONE (Content Only) */}
            <div className="flex-1 overflow-y-auto p-6 pb-48 min-h-0 custom-scrollbar bg-slate-800/50">


                {/* Builder UI */}
                <div className="space-y-6">

                    {/* [CALC] Unified Builder UI */}
                    {activeCategory === 'calc' && (
                        <div className="space-y-6">
                            {/* 1. Range Input */}
                            <div className="space-y-3">
                                <h3 className="text-white font-bold border-l-4 border-indigo-500 pl-3">1. 범위 선택</h3>
                                <div className="flex items-center gap-2 p-4 bg-slate-800 rounded-xl border border-slate-700">
                                    <input
                                        type="text"
                                        value={formulaState.start}
                                        onChange={(e) => setFormulaState({ ...formulaState, start: e.target.value.toUpperCase() })}
                                        className="w-20 bg-slate-900 border border-slate-600 rounded px-2 py-2 text-white text-center font-bold tracking-wider focus:border-indigo-500 outline-none"
                                        placeholder="A1"
                                    />
                                    <select
                                        value={formulaState.connector}
                                        onChange={(e) => setFormulaState({ ...formulaState, connector: e.target.value })}
                                        className="bg-slate-700 text-slate-300 rounded px-1 py-2 text-sm font-bold border-none outline-none cursor-pointer hover:bg-slate-600"
                                    >
                                        <option value="~">부터 (Range)</option>
                                        <option value=",">와 (And)</option>
                                    </select>
                                    <input
                                        type="text"
                                        value={formulaState.end}
                                        onChange={(e) => setFormulaState({ ...formulaState, end: e.target.value.toUpperCase() })}
                                        className="w-20 bg-slate-900 border border-slate-600 rounded px-2 py-2 text-white text-center font-bold tracking-wider focus:border-indigo-500 outline-none"
                                        placeholder="B10"
                                    />
                                </div>
                                <p className="text-xs text-slate-500 px-1">Tip: 엑셀 표를 드래그하면 자동으로 입력됩니다.</p>
                            </div>

                            {/* 2. Operation Select */}
                            <div className="space-y-3">
                                <h3 className="text-white font-bold border-l-4 border-pink-500 pl-3">2. 계산 방식</h3>
                                <div className="grid grid-cols-4 gap-2">
                                    {[
                                        { id: 'sum', label: '합계', icon: Plus },
                                        { id: 'avg', label: '평균', icon: Divide }, // Using Divide icon for Average metaphor
                                        { id: 'count', label: '개수', icon: Hash },
                                        { id: 'max', label: '최대', icon: ArrowUpAZ },
                                        { id: 'min', label: '최소', icon: ArrowDownAZ },
                                    ].map(op => (
                                        <button
                                            key={op.id}
                                            onClick={() => {
                                                setFormulaState({ ...formulaState, operation: op.id });
                                                if (isBuildingMode) {
                                                    const payload = { type: 'unified_calc', payload: { ...formulaState, operation: op.id } };
                                                    dispatchWrapper(payload, `계산: ${op.label}`);
                                                    setToastMsg(`✅ ${op.label} 추가됨`);
                                                    setTimeout(() => setToastMsg(null), 1000);
                                                }
                                            }}
                                            className={cn(
                                                "p-3 rounded-xl flex flex-col items-center gap-1 border-2 transition-all",
                                                formulaState.operation === op.id
                                                    ? "bg-indigo-600 border-indigo-500 text-white shadow-lg scale-105"
                                                    : "bg-slate-800 border-slate-700 text-slate-400 hover:bg-slate-700"
                                            )}
                                        >
                                            <op.icon size={18} />
                                            <span className="text-xs font-bold">{op.label}</span>
                                        </button>
                                    ))}
                                </div>
                            </div>

                            {/* 3. Result Mode & Target */}
                            <div className="space-y-3">
                                <h3 className="text-white font-bold border-l-4 border-teal-500 pl-3">3. 결과 위치</h3>
                                <div className="space-y-2">
                                    <div className="flex gap-2 p-1 bg-slate-800 rounded-xl border border-slate-700">
                                        {[
                                            { id: 'overwrite', label: '덮어쓰기', icon: Hash },
                                            { id: 'new', label: '새 칸', icon: Plus },
                                            { id: 'pick', label: '직접 선택', icon: Sparkles },
                                        ].map(mode => (
                                            <button
                                                key={mode.id}
                                                onClick={() => setFormulaState({ ...formulaState, resultMode: mode.id })}
                                                className={cn(
                                                    "flex-1 py-2 rounded-lg text-xs font-bold transition-all flex items-center justify-center gap-1",
                                                    formulaState.resultMode === mode.id
                                                        ? "bg-teal-600 text-white shadow-md"
                                                        : "text-slate-500 hover:bg-slate-700 hover:text-slate-300"
                                                )}
                                            >
                                                {mode.label}
                                            </button>
                                        ))}
                                    </div>

                                    {/* Target Input with Lock/Unlock */}
                                    {formulaState.resultMode === 'pick' && (
                                        <div className="flex items-center gap-2 p-2 bg-slate-800/50 rounded-lg border border-teal-500/30 animate-in fade-in slide-in-from-top-1">
                                            <span className="text-teal-500 text-xs font-bold whitespace-nowrap">목표 셀:</span>
                                            <input
                                                type="text"
                                                value={formulaState.target}
                                                // Allow manual edit only if unlocked? Or always allow? 
                                                // If unlocked, selection overrides it. If locked, selection ignores it.
                                                onChange={(e) => setFormulaState({ ...formulaState, target: e.target.value.toUpperCase() })}
                                                disabled={formulaState.isTargetLocked} // Disable input when locked
                                                className={cn("w-full bg-transparent text-sm font-bold outline-none transition-colors", formulaState.isTargetLocked ? "text-slate-400 cursor-not-allowed" : "text-white")}
                                                placeholder="선택하세요"
                                            />

                                            {/* Lock/Unlock Buttons */}
                                            {formulaState.isTargetLocked ? (
                                                <button
                                                    onClick={() => setFormulaState({ ...formulaState, isTargetLocked: false })}
                                                    className="p-1 rounded bg-slate-700 hover:bg-slate-600 text-slate-300 transition-colors"
                                                    title="다시 선택 (잠금 해제)"
                                                >
                                                    <RefreshCw size={14} />
                                                </button>
                                            ) : (
                                                <button
                                                    onClick={() => {
                                                        if (!formulaState.target) return;
                                                        setFormulaState({ ...formulaState, isTargetLocked: true })
                                                    }}
                                                    className="p-1 rounded bg-teal-600 hover:bg-teal-500 text-white transition-colors"
                                                    title="확인 (고정)"
                                                >
                                                    <div className="flex items-center gap-1 px-1">
                                                        <span className="text-[10px] font-bold">확인</span>
                                                    </div>
                                                </button>
                                            )}
                                        </div>
                                    )}
                                </div>
                            </div>

                            {/* 4. Execute Button (Shown only in Normal Mode for Calc, because One-Click is active in Builder) */}
                            {/* 4. Execute Button (Context Aware) */}
                            {isBuildingMode ? (
                                <button
                                    onClick={() => {
                                        // Reuse dispatchWrapper logic but with specific description
                                        const payload = getCurrentPayload();
                                        dispatchWrapper(payload, '계산 단계 추가');
                                        setToastMsg("✅ 계산 단계가 추가되었습니다.");
                                        setTimeout(() => setToastMsg(null), 1000);
                                    }}
                                    className="w-full py-4 bg-emerald-600 hover:bg-emerald-500 rounded-xl font-bold text-white shadow-lg hover:shadow-emerald-500/50 hover:scale-[1.02] active:scale-95 transition-all text-lg flex items-center justify-center gap-2 mt-4"
                                >
                                    <PlusCircle size={20} />
                                    <span>레시피에 계산 추가 (Add Calc)</span>
                                </button>
                            ) : (
                                <button
                                    onClick={executeCurrent}
                                    className="w-full py-4 bg-gradient-to-r from-indigo-600 to-pink-600 rounded-xl font-bold text-white shadow-lg hover:shadow-indigo-500/50 hover:scale-[1.02] active:scale-95 transition-all text-lg flex items-center justify-center gap-2"
                                >
                                    <Sparkles size={20} className="animate-pulse" />
                                    <span>계산 실행</span>
                                </button>
                            )}
                        </div>
                    )}

                    {/* [TEXT] UI */}
                    {activeCategory === 'text' && (
                        <>
                            <div className="space-y-3">
                                <h3 className="text-white font-bold border-l-4 border-emerald-500 pl-3">어떤 작업을 할까요?</h3>
                                <div className="grid grid-cols-3 gap-2">
                                    {[
                                        { id: 'join', label: '합치기', icon: Combine, desc: 'text_join' },
                                        { id: 'split', label: '나누기', icon: Split, desc: 'text_split' },
                                        { id: 'extract', label: '추출', icon: Scissors, desc: 'text_extract' },
                                    ].map((opt) => (
                                        <button
                                            key={opt.id}
                                            onClick={() => setTextMain(opt.id as any)}
                                            onMouseEnter={() => setHoveredDescription(DESCRIPTIONS[opt.desc])}
                                            onMouseLeave={() => setHoveredDescription(null)}
                                            className={cn("p-4 rounded-xl flex flex-col items-center gap-2 border-2 transition-all", textMain === opt.id ? "bg-emerald-600 border-emerald-500 text-white" : "bg-slate-800 border-slate-700 text-slate-400")}
                                        >
                                            <opt.icon size={24} />
                                            <span className="font-bold text-sm">{opt.label}</span>
                                        </button>
                                    ))}
                                </div>
                            </div>

                            <div className="space-y-3">
                                <h3 className="text-white font-bold border-l-4 border-orange-500 pl-3">세부 설정</h3>

                                {/* Dynamic Options based on Main */}
                                <div
                                    onMouseEnter={() => setHoveredDescription(DESCRIPTIONS['option_glue'])}
                                    onMouseLeave={() => setHoveredDescription(null)}
                                    className="p-4 bg-slate-800 rounded-xl border border-slate-700"
                                >
                                    {textMain === 'join' && (
                                        <div className="flex items-center gap-4">
                                            <span className="text-slate-400 text-sm font-bold">구분기호</span>
                                            <div className="flex gap-2">
                                                {[' ', ',', '-', '', '\n'].map(char => (
                                                    <button key={char} onClick={() => setTextOption({ ...textOption, delimiter: char })} className={cn("w-8 h-8 rounded bg-slate-900 border flex items-center justify-center text-sm", textOption.delimiter === char ? "border-orange-500 text-orange-500" : "border-slate-700 text-slate-500")}>
                                                        {char === ' ' ? '␣' : (char === '' ? '🚫' : (char === '\n' ? '↵' : char))}
                                                    </button>
                                                ))}
                                            </div>
                                        </div>
                                    )}
                                    {textMain === 'split' && (
                                        <div className="flex items-center gap-4">
                                            <span className="text-slate-400 text-sm font-bold">무엇으로 나눌까요?</span>
                                            <input type="text" value={textOption.delimiter} onChange={(e) => setTextOption({ ...textOption, delimiter: e.target.value })} className="w-12 bg-slate-900 border border-slate-600 rounded px-2 py-1 text-white text-center" placeholder="," />
                                        </div>
                                    )}
                                    {textMain === 'extract' && (
                                        <div className="flex items-center gap-4">
                                            <select value={textOption.mode} onChange={(e) => setTextOption({ ...textOption, mode: e.target.value as any })} className="bg-slate-900 border-slate-600 rounded p-1 text-white text-sm">
                                                <option value="left">앞에서</option>
                                                <option value="right">뒤에서</option>
                                            </select>
                                            <input type="number" value={textOption.count} onChange={(e) => setTextOption({ ...textOption, count: Number(e.target.value) })} className="w-12 bg-slate-900 border border-slate-600 rounded px-2 py-1 text-white text-center" />
                                            <span className="text-slate-400 text-sm">글자</span>
                                        </div>
                                    )}
                                </div>

                                <div className={cn("p-4 rounded-xl border transition-all cursor-pointer", textKeep ? "bg-slate-800 border-orange-500" : "bg-slate-800/50 border-slate-700")}>
                                    <label className="flex items-center gap-3 cursor-pointer">
                                        <input type="checkbox" checked={textKeep} onChange={(e) => setTextKeep(e.target.checked)} className="w-5 h-5 rounded bg-slate-700 border-slate-600 text-orange-500 focus:ring-0" />
                                        <div className="flex flex-col">
                                            <span className={cn("font-bold", textKeep ? "text-white" : "text-slate-400")}>원본 유지하기</span>
                                            <span className="text-xs text-slate-500">체크하면 결과가 오른쪽 칸에 생성됩니다</span>
                                        </div>
                                    </label>
                                </div>
                            </div>

                            <div className="pt-2">
                                {isBuildingMode && (
                                    <button onClick={handleAddStep} className="w-full py-3 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl font-bold flex items-center justify-center gap-2 shadow-lg">
                                        <Plus size={18} /> 레코드에 추가 (Text)
                                    </button>
                                )}
                            </div>
                        </>
                    )}

                    {/* [CLEAN] UI */}
                    {activeCategory === 'clean' && (
                        <>
                            <div className="space-y-3">
                                <h3 className="text-white font-bold border-l-4 border-slate-500 pl-3">무엇을 청소할까요?</h3>
                                <div className="grid grid-cols-3 gap-2">
                                    {[
                                        { id: 'clean_empty', label: '빈 줄 삭제', icon: Eraser, desc: 'clean_empty' },
                                        { id: 'remove_dup', label: '중복 제거', icon: Copy, desc: 'remove_dup' },
                                        { id: 'trim', label: '공백 정리', icon: Scissors, desc: 'trim' },
                                    ].map(opt => (
                                        <button
                                            key={opt.id}
                                            onClick={() => setCleanMain(opt.id as any)}
                                            onMouseEnter={() => setHoveredDescription(DESCRIPTIONS[opt.desc])}
                                            onMouseLeave={() => setHoveredDescription(null)}
                                            className={cn("p-4 rounded-xl flex flex-col items-center gap-2 border-2 transition-all", cleanMain === opt.id ? "bg-slate-600 border-slate-400 text-white" : "bg-slate-800 border-slate-700 text-slate-400")}
                                        >
                                            <opt.icon size={24} />
                                            <span className="font-bold text-sm">{opt.label}</span>
                                        </button>
                                    ))}
                                </div>
                            </div>
                            <div className="pt-2">
                                {isBuildingMode && (
                                    <button onClick={handleAddStep} className="w-full py-3 bg-slate-600 hover:bg-slate-500 text-white rounded-xl font-bold flex items-center justify-center gap-2 shadow-lg">
                                        <Plus size={18} /> 레코드에 추가 (Clean)
                                    </button>
                                )}
                            </div>
                        </>
                    )}

                    {/* [LOGIC] UI */}
                    {activeCategory === 'logic' && (
                        <>
                            <div className="space-y-3">
                                <h3 className="text-white font-bold border-l-4 border-cyan-500 pl-3">논리 함수</h3>
                                <div className="grid grid-cols-2 gap-2">
                                    {[
                                        { id: 'if', label: 'IF (조건 판단)', icon: Scale, desc: 'logic_if' },
                                        { id: 'vlookup', label: 'VLOOKUP (값 찾기)', icon: Search, desc: 'logic_vlookup' },
                                    ].map(opt => (
                                        <button
                                            key={opt.id}
                                            onClick={() => setLogicMain(opt.id as any)}
                                            onMouseEnter={() => setHoveredDescription(DESCRIPTIONS[opt.desc])}
                                            onMouseLeave={() => setHoveredDescription(null)}
                                            className={cn("p-4 rounded-xl flex flex-col items-center gap-2 border-2 transition-all", logicMain === opt.id ? "bg-cyan-600 border-cyan-500 text-white" : "bg-slate-800 border-slate-700 text-slate-400")}
                                        >
                                            <opt.icon size={24} />
                                            <span className="font-bold text-sm">{opt.label}</span>
                                        </button>
                                    ))}
                                </div>
                            </div>

                            {/* Logic Options */}
                            <div className="space-y-3">
                                <h3 className="text-white font-bold border-l-4 border-yellow-500 pl-3">설정</h3>
                                <div className="p-4 bg-slate-800 rounded-xl border border-slate-700 space-y-4">

                                    {logicMain === 'if' && (
                                        <div className="space-y-3 animate-in fade-in slide-in-from-top-2">
                                            {/* IF Condition */}
                                            <div className="bg-slate-900/50 p-3 rounded-lg border border-slate-600 space-y-2">
                                                <div className="flex items-center justify-between text-xs text-slate-400 font-bold mb-1">
                                                    <span>조건 (Condition)</span>
                                                    <span className="text-cyan-400">현재 셀 기준</span>
                                                </div>
                                                <div className="flex items-center gap-2">
                                                    <select
                                                        value={logicIf.operator}
                                                        onChange={(e) => setLogicIf({ ...logicIf, operator: e.target.value })}
                                                        className="bg-slate-800 border-slate-600 rounded p-2 text-white text-sm font-bold flex-1"
                                                    >
                                                        <option value=">">&gt; (크다)</option>
                                                        <option value=">=">&ge; (크거나 같다)</option>
                                                        <option value="<">&lt; (작다)</option>
                                                        <option value="<=">&le; (작거나 같다)</option>
                                                        <option value="=">= (같다)</option>
                                                        <option value="!=">&ne; (다르다)</option>
                                                    </select>
                                                    <input
                                                        type="text"
                                                        value={logicIf.value}
                                                        onChange={(e) => setLogicIf({ ...logicIf, value: e.target.value })}
                                                        placeholder="비교값 (예: 60)"
                                                        className="flex-[2] bg-slate-800 border-slate-600 rounded px-3 py-2 text-white text-sm font-bold"
                                                    />
                                                </div>
                                            </div>

                                            {/* True/False Values */}
                                            <div className="grid grid-cols-2 gap-2">
                                                <div className="space-y-1">
                                                    <label className="text-xs font-bold text-green-400 pl-1">참(True)일 때</label>
                                                    <input
                                                        type="text"
                                                        value={logicIf.trueVal}
                                                        onChange={(e) => setLogicIf({ ...logicIf, trueVal: e.target.value })}
                                                        placeholder="예: 합격"
                                                        className="w-full bg-slate-900 border border-green-500/30 rounded px-3 py-2 text-white text-sm focus:border-green-500 transition-colors"
                                                    />
                                                </div>
                                                <div className="space-y-1">
                                                    <label className="text-xs font-bold text-rose-400 pl-1">거짓(False)일 때</label>
                                                    <input
                                                        type="text"
                                                        value={logicIf.falseVal}
                                                        onChange={(e) => setLogicIf({ ...logicIf, falseVal: e.target.value })}
                                                        placeholder="예: 불합격"
                                                        className="w-full bg-slate-900 border border-rose-500/30 rounded px-3 py-2 text-white text-sm focus:border-rose-500 transition-colors"
                                                    />
                                                </div>
                                            </div>
                                            <p className="text-xs text-slate-500 pt-2 border-t border-slate-700">
                                                * 결과는 선택한 셀의 <span className="text-cyan-400 font-bold">오른쪽 칸</span>에 표시됩니다.
                                            </p>
                                        </div>
                                    )}

                                    {logicMain === 'vlookup' && (
                                        <div className="space-y-4 animate-in fade-in slide-in-from-top-2">
                                            {/* Lookup Visual */}
                                            <div className="bg-slate-900/50 p-3 rounded-lg border border-slate-600 flex items-center gap-3">
                                                <Search size={18} className="text-cyan-400" />
                                                <div className="flex flex-col">
                                                    <span className="text-xs text-slate-400 font-bold">찾을 값 (Lookup Value)</span>
                                                    <span className="text-sm text-white font-mono">현재 선택된 셀의 값</span>
                                                </div>
                                            </div>

                                            {/* Range Input */}
                                            <div className="space-y-1">
                                                <label className="text-xs font-bold text-slate-300 pl-1">찾을 범위 (Table Array)</label>
                                                <input
                                                    type="text"
                                                    value={logicVlookup.range}
                                                    onChange={(e) => setLogicVlookup({ ...logicVlookup, range: e.target.value.toUpperCase() })}
                                                    placeholder="예: A1:C10"
                                                    className="w-full bg-slate-900 border border-slate-600 rounded px-3 py-2 text-white text-sm font-bold tracking-wider"
                                                />
                                            </div>

                                            {/* Col Index Input */}
                                            <div className="space-y-1">
                                                <label className="text-xs font-bold text-slate-300 pl-1">가져올 열 번호 (Col Index)</label>
                                                <div className="flex items-center gap-2">
                                                    <input
                                                        type="number"
                                                        min="1"
                                                        value={logicVlookup.colIndex}
                                                        onChange={(e) => setLogicVlookup({ ...logicVlookup, colIndex: Number(e.target.value) })}
                                                        className="w-20 bg-slate-900 border border-slate-600 rounded px-3 py-2 text-white text-sm font-bold text-center"
                                                    />
                                                    <span className="text-xs text-slate-500">번째 열의 값을 가져옵니다.</span>
                                                </div>
                                            </div>

                                            <p className="text-xs text-slate-500 pt-2 border-t border-slate-700">
                                                * 정확히 일치하는 값만 찾습니다 (False/0). <br />
                                                * 결과는 <span className="text-cyan-400 font-bold">오른쪽 칸</span>에 표시됩니다.
                                            </p>
                                        </div>
                                    )}

                                    <button
                                        onClick={executeCurrent}
                                        className="w-full py-3 bg-gradient-to-r from-cyan-600 to-blue-600 rounded-lg font-bold text-white shadow-lg hover:shadow-cyan-500/50 transition-all active:scale-95 flex items-center justify-center gap-2"
                                    >
                                        <Sparkles size={18} />
                                        <span>{isBuildingMode ? "레코드에 추가 (Add Step)" : "함수 실행 (Insert Formula)"}</span>
                                    </button>
                                </div>
                            </div>
                        </>
                    )}

                    {activeCategory === 'clean' && (
                        <div className="space-y-3">
                            <h3 className="text-white font-bold border-l-4 border-slate-500 pl-3">데이터 정리 (Clean)</h3>
                            <div className="p-4 bg-slate-800 rounded-xl border border-slate-700">
                                {cleanMain === 'trim' ? (
                                    <label
                                        onMouseEnter={() => setHoveredDescription(DESCRIPTIONS['opt_all_spaces'])}
                                        onMouseLeave={() => setHoveredDescription(null)}
                                        className="flex items-center gap-3 cursor-pointer"
                                    >
                                        <input type="checkbox" checked={cleanOption.allSpaces} onChange={(e) => setCleanOption({ ...cleanOption, allSpaces: e.target.checked })} className="w-5 h-5 rounded bg-slate-700 border-slate-600 text-slate-500 focus:ring-0" />
                                        <span className={cn("font-bold", cleanOption.allSpaces ? "text-white" : "text-slate-400")}>모든 공백 제거 (띄어쓰기 포함)</span>
                                    </label>
                                ) : (
                                    <p className="text-slate-500 text-sm">별도의 옵션이 없습니다.</p>
                                )}
                            </div>
                        </div>
                    )}

                    {/* [STYLE] UI */}
                    {
                        activeCategory === 'style' && (
                            <>
                                {/* Manual Toolbar */}
                                <div className="space-y-4 mb-6">
                                    <h3 className="text-white font-bold border-l-4 border-indigo-500 pl-3">기본 서식</h3>

                                    {/* Row 1: Font & Align */}
                                    <div className="p-4 bg-slate-800 rounded-xl border border-slate-700 flex flex-wrap gap-4 items-center">
                                        {/* Font Styles */}
                                        <div className="flex bg-slate-900 rounded-lg p-1 border border-slate-600">
                                            <button onClick={() => dispatchWrapper({ type: 'style', payload: { type: 'bold' } }, '굵게')} className="p-2 hover:bg-slate-700 rounded text-slate-300 hover:text-white transition-colors" title="굵게"><Bold size={18} /></button>
                                            <button onClick={() => dispatchWrapper({ type: 'style', payload: { type: 'italic' } }, '기울임')} className="p-2 hover:bg-slate-700 rounded text-slate-300 hover:text-white transition-colors" title="기울임"><Italic size={18} /></button>
                                            <button onClick={() => dispatchWrapper({ type: 'style', payload: { type: 'underline' } }, '밑줄')} className="p-2 hover:bg-slate-700 rounded text-slate-300 hover:text-white transition-colors" title="밑줄"><Underline size={18} /></button>
                                        </div>

                                        <div className="w-px h-8 bg-slate-600 mx-2"></div>

                                        {/* Alignment */}
                                        <div className="flex bg-slate-900 rounded-lg p-1 border border-slate-600">
                                            <button onClick={() => dispatchWrapper({ type: 'style', payload: { type: 'align', value: 'left' } }, '왼쪽 정렬')} className="p-2 hover:bg-slate-700 rounded text-slate-300 hover:text-white transition-colors" title="왼쪽 정렬"><AlignLeft size={18} /></button>
                                            <button onClick={() => dispatchWrapper({ type: 'style', payload: { type: 'align', value: 'center' } }, '가운데 정렬')} className="p-2 hover:bg-slate-700 rounded text-slate-300 hover:text-white transition-colors" title="가운데 정렬"><AlignCenter size={18} /></button>
                                            <button onClick={() => dispatchWrapper({ type: 'style', payload: { type: 'align', value: 'right' } }, '오른쪽 정렬')} className="p-2 hover:bg-slate-700 rounded text-slate-300 hover:text-white transition-colors" title="오른쪽 정렬"><AlignRight size={18} /></button>
                                        </div>
                                    </div>

                                    {/* Row 2: Colors */}
                                    <div className="grid grid-cols-2 gap-2">
                                        {/* Text Color */}
                                        <div className="p-3 bg-slate-800 rounded-xl border border-slate-700 flex flex-col gap-2">
                                            <div className="flex items-center gap-2 text-slate-400 text-xs font-bold">
                                                <Palette size={14} /> 글자색
                                            </div>
                                            <div className="flex gap-1 justify-between">
                                                {[
                                                    { c: '000000', n: '검정' }, { c: 'EF4444', n: '빨강' }, { c: '3B82F6', n: '파랑' }, { c: '10B981', n: '초록' }, { c: 'FFFFFF', n: '흰색' }
                                                ].map(color => (
                                                    <button key={color.c} onClick={() => dispatchWrapper({ type: 'style', payload: { type: 'color', value: color.c } }, `글자색: ${color.n}`)} className="w-6 h-6 rounded-full border border-slate-600 hover:scale-110" style={{ backgroundColor: `#${color.c}` }} title={color.n} />
                                                ))}
                                            </div>
                                        </div>
                                        {/* Fill Color */}
                                        <div className="p-3 bg-slate-800 rounded-xl border border-slate-700 flex flex-col gap-2">
                                            <div className="flex items-center gap-2 text-slate-400 text-xs font-bold">
                                                <PaintBucket size={14} /> 배경색
                                            </div>
                                            <div className="flex gap-1 justify-between">
                                                {[
                                                    { c: 'FFFFFF', n: '없음' }, { c: 'FEF08A', n: '노랑' }, { c: 'FECACA', n: '빨강(연)' }, { c: 'BFDBFE', n: '파랑(연)' }, { c: 'E5E7EB', n: '회색' }
                                                ].map(color => (
                                                    <button key={color.c} onClick={() => dispatchWrapper({ type: 'style', payload: { type: 'fill', value: color.c } }, `배경색: ${color.n}`)} className="w-6 h-6 rounded-full border border-slate-600 hover:scale-110" style={{ backgroundColor: `#${color.c}` }} title={color.n} />
                                                ))}
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                {/* Existing Smart Features */}

                                <div className="space-y-3">
                                    <h3 className="text-white font-bold border-l-4 border-pink-500 pl-3">어떻게 꾸밀까요?</h3>
                                    <div className="grid grid-cols-3 gap-2">
                                        {[
                                            { id: 'comma', label: '천단위 콤마', icon: Hash, desc: 'comma' },
                                            { id: 'header_style', label: '헤더 강조', icon: Rows, desc: 'header_style' },
                                            { id: 'highlight', label: '조건부 강조', icon: Paintbrush, desc: 'style_highlight' },
                                        ].map(opt => (
                                            <button
                                                key={opt.id}
                                                onClick={() => setStyleMain(opt.id as any)}
                                                onMouseEnter={() => setHoveredDescription(DESCRIPTIONS[opt.desc])}
                                                onMouseLeave={() => setHoveredDescription(null)}
                                                className={cn("p-4 rounded-xl flex flex-col items-center gap-2 border-2 transition-all", styleMain === opt.id ? "bg-pink-600 border-pink-500 text-white" : "bg-slate-800 border-slate-700 text-slate-400")}
                                            >
                                                <opt.icon size={24} />
                                                <span className="font-bold text-sm">{opt.label}</span>
                                            </button>
                                        ))}
                                    </div>
                                </div>

                                {/* Style Options */}
                                <div className="space-y-3">
                                    <h3 className="text-white font-bold border-l-4 border-purple-500 pl-3">세부 설정</h3>
                                    <div className="p-4 bg-slate-800 rounded-xl border border-slate-700 space-y-3">
                                        {styleMain === 'highlight' ? (
                                            <>
                                                <div className="flex items-center gap-2">
                                                    <span className="text-slate-400 font-bold text-sm w-12">조건</span>
                                                    <select value={styleLogic.operator} onChange={(e) => setStyleLogic({ ...styleLogic, operator: e.target.value })} className="bg-slate-900 border-slate-600 rounded p-1 text-white text-sm"><option value=">">&gt;</option><option value="<">&lt;</option><option value="=">=</option><option value="contains">포함</option></select>
                                                    <input type="text" value={styleLogic.value} onChange={(e) => setStyleLogic({ ...styleLogic, value: e.target.value })} placeholder="값" className="flex-1 bg-slate-900 border-slate-600 rounded px-2 py-1 text-white text-sm" />
                                                </div>
                                                <div className="flex items-center gap-2">
                                                    <span className="text-slate-400 font-bold text-sm w-12">색상</span>
                                                    <div className="flex gap-2">
                                                        {[
                                                            { id: 'yellow', bg: 'bg-yellow-400/20', border: 'border-yellow-400' },
                                                            { id: 'red', bg: 'bg-red-400/20', border: 'border-red-400' },
                                                            { id: 'green', bg: 'bg-green-400/20', border: 'border-green-400' },
                                                            { id: 'blue', bg: 'bg-blue-400/20', border: 'border-blue-400' },
                                                        ].map(c => (
                                                            <button
                                                                key={c.id}
                                                                onClick={() => setStyleLogic({ ...styleLogic, color: c.id })}
                                                                className={cn("w-8 h-8 rounded-full border-2 transition-all", c.bg, styleLogic.color === c.id ? c.border + " scale-110 shadow-lg" : "border-transparent opacity-50 hover:opacity-100")}
                                                            />
                                                        ))}
                                                    </div>
                                                </div>
                                            </>
                                        ) : (
                                            <p className="text-slate-500 text-sm">기본 설정으로 적용됩니다.</p>
                                        )}
                                    </div>
                                </div>
                            </>
                        )
                    }

                    {/* [ANALYZE] UI */}
                    {
                        activeCategory === 'analyze' && (
                            <>
                                <div className="space-y-3">
                                    <h3 className="text-white font-bold border-l-4 border-blue-500 pl-3">무엇을 분석할까요?</h3>
                                    <div className="grid grid-cols-2 gap-2">
                                        {[
                                            { id: 'stat_basic', label: '기초 통계', icon: Sigma, desc: 'stat_basic' },
                                            { id: 'chart_bar', label: '막대 그래프', icon: BarChart3, desc: 'chart_bar' },
                                            { id: 'chart_line', label: '선 그래프', icon: LineChart, desc: 'chart_line' },
                                            { id: 'chart_scatter', label: '산점도/회귀', icon: ScatterChart, desc: 'chart_scatter' },
                                        ].map(opt => (
                                            <button
                                                key={opt.id}
                                                onClick={() => setAnalyzeMain(opt.id as any)}
                                                onMouseEnter={() => setHoveredDescription(DESCRIPTIONS[opt.desc])}
                                                onMouseLeave={() => setHoveredDescription(null)}
                                                className={cn("p-4 rounded-xl flex flex-col items-center gap-2 border-2 transition-all", analyzeMain === opt.id ? "bg-blue-600 border-blue-500 text-white" : "bg-slate-800 border-slate-700 text-slate-400")}
                                            >
                                                <opt.icon size={24} />
                                                <span className="font-bold text-sm">{opt.label}</span>
                                            </button>
                                        ))}
                                    </div>
                                </div>

                                {/* Analyze Options */}
                                <div className="space-y-3">
                                    <h3 className="text-white font-bold border-l-4 border-cyan-500 pl-3">세부 설정</h3>
                                    <div className="p-4 bg-slate-800 rounded-xl border border-slate-700">
                                        {analyzeMain.startsWith('chart') ? (
                                            <label
                                                onMouseEnter={() => setHoveredDescription(DESCRIPTIONS['opt_label'])}
                                                onMouseLeave={() => setHoveredDescription(null)}
                                                className="flex items-center gap-3 cursor-pointer"
                                            >
                                                <input type="checkbox" checked={analyzeOption.label} onChange={(e) => setAnalyzeOption({ ...analyzeOption, label: e.target.checked })} className="w-5 h-5 rounded bg-slate-700 border-slate-600 text-cyan-500 focus:ring-0" />
                                                <span className={cn("font-bold", analyzeOption.label ? "text-white" : "text-slate-400")}>첫 번째 열을 라벨(이름)로 사용</span>
                                            </label>
                                        ) : (
                                            <p className="text-slate-500 text-sm">별도의 옵션이 없습니다.</p>
                                        )}
                                    </div>
                                    {isBuildingMode && (
                                        <button onClick={handleAddStep} className="w-full mt-4 py-3 bg-blue-600 hover:bg-blue-500 text-white rounded-xl font-bold flex items-center justify-center gap-2 shadow-lg">
                                            <Plus size={18} /> 레코드에 추가 (Analyze)
                                        </button>
                                    )}
                                </div>
                            </>
                        )
                    }


                    {/* [DATA] UI */}
                    {
                        activeCategory === 'data' && (
                            <>
                                <div className="space-y-3">
                                    <h3 className="text-white font-bold border-l-4 border-emerald-500 pl-3">데이터 관리</h3>
                                    <div className="grid grid-cols-2 gap-2">
                                        {[
                                            { id: 'sort_asc', label: '오름차순 정렬', icon: ArrowDownAZ, desc: 'data_sort_asc' },
                                            { id: 'sort_desc', label: '내림차순 정렬', icon: ArrowUpAZ, desc: 'data_sort_desc' },
                                            { id: 'filter', label: '조건 추출(Filter)', icon: Filter, desc: 'data_filter' },
                                            { id: 'replace', label: '찾아 바꾸기', icon: ReplaceIcon, desc: 'data_replace' },
                                        ].map(opt => (
                                            <button
                                                key={opt.id}
                                                onClick={() => setDataMain(opt.id as any)}
                                                onMouseEnter={() => setHoveredDescription(DESCRIPTIONS[opt.desc])}
                                                onMouseLeave={() => setHoveredDescription(null)}
                                                className={cn("p-4 rounded-xl flex flex-col items-center gap-2 border-2 transition-all", dataMain === opt.id ? "bg-emerald-600 border-emerald-500 text-white" : "bg-slate-800 border-slate-700 text-slate-400")}
                                            >
                                                <opt.icon size={24} />
                                                <span className="font-bold text-sm">{opt.label}</span>
                                            </button>
                                        ))}
                                    </div>
                                </div>

                                {/* Data Options */}
                                <div className="space-y-3">
                                    <h3 className="text-white font-bold border-l-4 border-teal-500 pl-3">세부 설정</h3>
                                    <div className="p-4 bg-slate-800 rounded-xl border border-slate-700 space-y-3">
                                        {/* Header Option (Common) */}
                                        <label
                                            onMouseEnter={() => setHoveredDescription(DESCRIPTIONS['opt_header'])}
                                            onMouseLeave={() => setHoveredDescription(null)}
                                            className="flex items-center gap-3 cursor-pointer mb-2 border-b border-slate-700 pb-2"
                                        >
                                            <input type="checkbox" checked={dataOption.header} onChange={(e) => setDataOption({ ...dataOption, header: e.target.checked })} className="w-5 h-5 rounded bg-slate-700 border-slate-600 text-teal-500 focus:ring-0" />
                                            <span className={cn("font-bold text-sm", dataOption.header ? "text-white" : "text-slate-400")}>첫 번째 행(제목) 제외</span>
                                        </label>

                                        {dataMain.startsWith('sort') && (
                                            <p className="text-slate-500 text-xs text-center pt-1">
                                                현재 선택된 <strong>열(Column)</strong>을 기준으로 정렬합니다.
                                            </p>
                                        )}

                                        {dataMain === 'filter' && (
                                            <div className="flex flex-col gap-2">
                                                <label className="text-slate-400 text-xs font-bold">조건 (예: {'>'}50, 서울)</label>
                                                <input
                                                    type="text"
                                                    value={dataOption.condition}
                                                    onChange={(e) => setDataOption({ ...dataOption, condition: e.target.value })}
                                                    onMouseEnter={() => setHoveredDescription(DESCRIPTIONS['opt_condition'])}
                                                    onMouseLeave={() => setHoveredDescription(null)}
                                                    placeholder="조건 입력"
                                                    className="w-full bg-slate-900 border border-slate-600 rounded px-3 py-2 text-white text-sm focus:border-teal-500 outline-none"
                                                />
                                            </div>
                                        )}

                                        {dataMain === 'replace' && (
                                            <div className="flex flex-col gap-2">
                                                <div className="grid grid-cols-2 gap-2">
                                                    <div>
                                                        <label className="text-slate-400 text-xs font-bold">찾을 내용</label>
                                                        <input
                                                            type="text"
                                                            value={dataOption.find}
                                                            onChange={(e) => setDataOption({ ...dataOption, find: e.target.value })}
                                                            onMouseEnter={() => setHoveredDescription(DESCRIPTIONS['opt_find'])}
                                                            onMouseLeave={() => setHoveredDescription(null)}
                                                            className="w-full bg-slate-900 border border-slate-600 rounded px-3 py-2 text-white text-sm focus:border-teal-500 outline-none"
                                                        />
                                                    </div>
                                                    <div>
                                                        <label className="text-slate-400 text-xs font-bold">바꿀 내용</label>
                                                        <input
                                                            type="text"
                                                            value={dataOption.replace}
                                                            onChange={(e) => setDataOption({ ...dataOption, replace: e.target.value })}
                                                            onMouseEnter={() => setHoveredDescription(DESCRIPTIONS['opt_replace'])}
                                                            onMouseLeave={() => setHoveredDescription(null)}
                                                            className="w-full bg-slate-900 border border-slate-600 rounded px-3 py-2 text-white text-sm focus:border-teal-500 outline-none"
                                                        />
                                                    </div>
                                                </div>
                                            </div>
                                        )}
                                    </div>
                                    {isBuildingMode && (
                                        <button onClick={handleAddStep} className="w-full mt-4 py-3 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl font-bold flex items-center justify-center gap-2 shadow-lg">
                                            <Plus size={18} /> 레코드에 추가 (Data)
                                        </button>
                                    )}
                                </div>
                            </>
                        )
                    }

                    {/* My Recipes Section */}
                    {activeCategory === 'my_recipe' && (
                        <div className="flex-1 overflow-y-auto p-5">
                            <div className="space-y-4 animate-in fade-in zoom-in-95 duration-200">
                                <h3 className="text-slate-400 text-xs font-bold uppercase tracking-wider mb-2">저장된 버튼 ({savedRecipes.length})</h3>

                                {savedRecipes.length === 0 ? (
                                    <div className="text-center p-8 border border-dashed border-slate-700 rounded-xl">
                                        <ChefHat className="mx-auto text-slate-600 mb-2" />
                                        <p className="text-slate-500 text-sm">저장된 버튼이 없습니다.</p>
                                    </div>
                                ) : (
                                    <div className="grid grid-cols-2 gap-3">
                                        {savedRecipes.map(recipe => (
                                            <div key={recipe.id} className="relative group">
                                                <button
                                                    onClick={() => handleExecuteFavorite(recipe)}
                                                    className="w-full text-left bg-slate-800 hover:bg-indigo-900/30 border border-slate-700 hover:border-indigo-500/50 p-3 rounded-xl transition-all shadow-sm group-hover:shadow-md"
                                                >
                                                    <div className="flex items-center gap-3 mb-1">
                                                        <div className="w-8 h-8 rounded-lg bg-indigo-600/20 text-indigo-400 flex items-center justify-center">
                                                            <Play size={14} fill="currentColor" />
                                                        </div>
                                                        <span className="font-bold text-slate-200 text-sm truncate">{recipe.name}</span>
                                                    </div>
                                                    <div className="text-[10px] text-slate-500 pl-11">
                                                        {recipe.category === 'my_recipe' ? `Step: ${recipe.queue?.length || 0}` : '단일 기능'}
                                                    </div>
                                                </button>
                                                <button
                                                    onClick={(e) => deleteRecipe(recipe.id, e)}
                                                    className="absolute top-2 right-2 p-1.5 text-slate-500 hover:text-red-400 hover:bg-slate-700/50 rounded-lg opacity-0 group-hover:opacity-100 transition-all"
                                                >
                                                    <Trash2 size={14} />
                                                </button>
                                            </div>
                                        ))}
                                    </div>
                                )}
                            </div>
                        </div>
                    )}
                </div>
            </div>

            {/* Info Deck */}
            <div className="mt-auto p-5 bg-slate-900/80 border-t border-slate-700/50 backdrop-blur-sm">
                <div className="flex items-start gap-3">
                    <Quote size={20} className={cn("mt-1 transition-colors", hoveredDescription ? "text-indigo-400" : "text-slate-600")} />
                    <p className={cn("text-sm font-medium leading-relaxed transition-colors duration-300", hoveredDescription ? "text-slate-200" : "text-slate-500")}>
                        {hoveredDescription || "원하는 기능을 마우스로 가리켜보세요."}
                    </p>
                </div>
            </div>

            {/* Footer Action */}
            <div className="sticky bottom-0 bg-slate-900/95 backdrop-blur border-t border-slate-700/50 p-4 flex gap-3 z-10">
                {isBuildingMode ? (
                    /* Footer Hidden in Building Mode */
                    null
                ) : (
                    <button onClick={executeCurrent} className="flex-1 bg-indigo-600 hover:bg-indigo-500 text-white py-4 rounded-xl font-bold text-lg shadow-lg flex items-center justify-center gap-2 transition-all">
                        <Play fill="currentColor" size={20} /> 실행
                    </button>
                )}


                {/* Recipe Save (Visible only in Normal Mode? Or Always?) 
                                    Builder Save is in Header. Standard Save (favorites) is here.
                                */}
                {!isBuildingMode && (
                    <button
                        onClick={() => setIsNamingRecipe(true)}
                        onMouseEnter={() => setHoveredDescription(DESCRIPTIONS['save_recipe'])}
                        onMouseLeave={() => setHoveredDescription(null)}
                        className="px-4 bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-700 rounded-xl transition-all"
                    >
                        <Save size={24} />
                    </button>
                )}
            </div>

            {/* Name Modal */}
            {
                isNamingRecipe && (
                    <div className="absolute inset-0 z-50 bg-slate-900/90 backdrop-blur-sm flex items-center justify-center p-6 animate-in fade-in duration-200">
                        <div className="w-full max-w-sm bg-slate-800 p-6 rounded-2xl border border-slate-700 shadow-2xl">
                            <h3 className="text-xl font-bold text-white mb-4">나만의 버튼 이름 짓기</h3>
                            <input
                                autoFocus
                                type="text"
                                className="w-full bg-slate-900 border border-slate-600 rounded-xl px-4 py-3 text-white mb-4 focus:ring-2 focus:ring-indigo-500 outline-none"
                                placeholder="예: 이름 합치기"
                                value={newRecipeName}
                                onChange={(e) => setNewRecipeName(e.target.value)}
                                onKeyDown={(e) => e.key === 'Enter' && saveRecipe()}
                            />
                            <div className="flex gap-2">
                                <button onClick={saveRecipe} className="flex-1 bg-indigo-600 hover:bg-indigo-500 text-white py-3 rounded-xl font-bold transition-colors">저장</button>
                                <button onClick={() => setIsNamingRecipe(false)} className="px-6 bg-slate-700 hover:bg-slate-600 text-slate-300 py-3 rounded-xl font-bold transition-colors">취소</button>
                            </div>
                        </div>
                    </div>
                )
            }
        </div>
    );
}
