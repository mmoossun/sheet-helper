"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useSession, signIn, signOut } from "next-auth/react";
import { AgGridReact } from "ag-grid-react";
import {
  AllCommunityModule,
  ModuleRegistry,
  type ColDef,
  type GridApi,
  type GridReadyEvent,
  type CellValueChangedEvent,
  type CellEditingStartedEvent,
  type CellEditingStoppedEvent,
  type RowSelectionOptions,
} from "ag-grid-community";
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  LineChart,
  Line,
  PieChart,
  Pie,
  Cell,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
} from "recharts";

const PIE_COLORS = ["#6366f1", "#22c55e", "#f59e0b", "#ef4444", "#06b6d4", "#a855f7", "#ec4899", "#84cc16"];

type Agg = "sum" | "avg" | "count" | "max" | "min";
const AGG_OPTIONS: { key: Agg; label: string }[] = [
  { key: "sum", label: "합계" },
  { key: "avg", label: "평균" },
  { key: "count", label: "개수" },
  { key: "max", label: "최댓값" },
  { key: "min", label: "최솟값" },
];

// AG Grid v33+ 는 모듈 등록이 필요하다.
ModuleRegistry.registerModules([AllCommunityModule]);

const POLL_INTERVAL_MS = 5000;

/** 0 → A, 1 → B, ... 26 → AA 형태의 스프레드시트 열 이름 */
function columnLetter(index: number): string {
  let s = "";
  let i = index + 1;
  while (i > 0) {
    const m = (i - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
}

/** "A" → 0, "B" → 1, ... "AA" → 26. columnLetter 의 역함수. */
function letterToIndex(letter: string): number {
  let n = 0;
  for (let k = 0; k < letter.length; k++) {
    n = n * 26 + (letter.charCodeAt(k) - 64);
  }
  return n - 1;
}

/** 전체 URL을 붙여넣어도 시트 ID만 뽑아낸다. */
function extractSpreadsheetId(input: string): string {
  const trimmed = input.trim();
  const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : trimmed;
}

type ColValidation = { type: "list" | "checkbox"; values?: string[] };

/** 데이터 열 개수에 맞춰 컬럼 정의를 만든다 (맨 앞에 행번호 열). 검증 있는 열은 드롭다운 편집기. */
function buildColumnDefs(
  maxCols: number,
  validations: Record<string, ColValidation> = {},
): ColDef[] {
  return [
    {
      headerName: "",
      field: "__row",
      width: 56,
      pinned: "left",
      sortable: false,
      filter: false,
      editable: false,
      cellClass: "ag-row-number",
    },
    ...Array.from({ length: Math.max(maxCols, 1) }, (_, i) => {
      const field = columnLetter(i);
      const v = validations[field];
      const def: ColDef = { headerName: field, field, flex: 1 };
      if (v) {
        const list = v.type === "checkbox" ? ["TRUE", "FALSE"] : v.values ?? [];
        def.cellEditor = "agSelectCellEditor";
        def.cellEditorParams = { values: list };
        def.headerName = `${field} ▾`;
      }
      return def;
    }),
  ];
}

type RowShape = Record<string, string | number>;
type SaveStatus = "idle" | "saving" | "saved" | "error";
type SheetTab = { sheetId: number; title: string; index: number; hidden: boolean };

// Level 1 함수 블록 — 클릭하면 포커스한 셀에 수식을 자동으로 넣어준다.
type FormulaBlock = {
  key: string;
  label: string;
  desc: string;
  make: (col: string, row: number) => string;
};
const FORMULA_BLOCKS: FormulaBlock[] = [
  { key: "SUM", label: "합계", desc: "위 칸들을 모두 더해요", make: (c, r) => `=SUM(${c}1:${c}${Math.max(r - 1, 1)})` },
  { key: "AVERAGE", label: "평균", desc: "위 칸들의 평균", make: (c, r) => `=AVERAGE(${c}1:${c}${Math.max(r - 1, 1)})` },
  { key: "COUNT", label: "개수", desc: "값이 있는 칸의 개수", make: (c, r) => `=COUNTA(${c}1:${c}${Math.max(r - 1, 1)})` },
  { key: "MAX", label: "최댓값", desc: "가장 큰 값", make: (c, r) => `=MAX(${c}1:${c}${Math.max(r - 1, 1)})` },
  { key: "MIN", label: "최솟값", desc: "가장 작은 값", make: (c, r) => `=MIN(${c}1:${c}${Math.max(r - 1, 1)})` },
  { key: "TODAY", label: "오늘 날짜", desc: "오늘 날짜를 자동으로", make: () => `=TODAY()` },
  { key: "IF", label: "조건(IF)", desc: "조건에 따라 다른 값", make: (c, r) => `=IF(${c}${Math.max(r - 1, 1)}>0, "양수", "음수")` },
  { key: "VLOOKUP", label: "찾기(VLOOKUP)", desc: "다른 표에서 값 찾기", make: () => `=VLOOKUP("찾을값", A:B, 2, FALSE)` },
];

export default function Home() {
  const { data: session, status } = useSession();
  const [sheetInput, setSheetInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [rowData, setRowData] = useState<RowShape[]>([]);
  const [columnDefs, setColumnDefs] = useState<ColDef[]>([]);
  const [loadedId, setLoadedId] = useState<string | null>(null);
  const [tabs, setTabs] = useState<SheetTab[]>([]);
  const [activeTab, setActiveTab] = useState<string | null>(null);
  const [saveStatus, setSaveStatus] = useState<SaveStatus>("idle");
  const [aiOpen, setAiOpen] = useState(false);
  const [aiPrompt, setAiPrompt] = useState("");
  const [aiLoading, setAiLoading] = useState(false);
  const [blockOpen, setBlockOpen] = useState(false);
  const [chartOpen, setChartOpen] = useState(false);
  const [ruleOpen, setRuleOpen] = useState(false);
  const [dropdownOptions, setDropdownOptions] = useState("");

  const gridApiRef = useRef<GridApi | null>(null);
  const colCountRef = useRef(0); // 현재 데이터 열 개수(행번호 열 제외)
  const validationsRef = useRef<Record<string, ColValidation>>({}); // 열별 입력 규칙

  // dirty 보호용 자료구조
  const pendingRef = useRef<Map<string, string>>(new Map()); // 저장 대기(아직 전송 안 함)
  const lastWrittenRef = useRef<Map<string, string>>(new Map()); // 저장 완료, 폴링 확인 대기
  const editingRef = useRef<Set<string>>(new Set()); // 현재 편집 중인 셀
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const defaultColDef = useMemo<ColDef>(
    () => ({ resizable: true, sortable: true, filter: true, minWidth: 90, editable: true }),
    [],
  );

  // 행 선택(체크박스, 다중) — 행 삭제에 사용
  const rowSelection = useMemo<RowSelectionOptions>(
    () => ({ mode: "multiRow", enableClickSelection: false }),
    [],
  );

  // 세션이 만료되면(리프레시 실패) 재로그인 유도
  useEffect(() => {
    if (session?.error === "RefreshAccessTokenError") {
      signIn("google");
    }
  }, [session?.error]);

  const onGridReady = useCallback((e: GridReadyEvent) => {
    gridApiRef.current = e.api;
  }, []);

  // 특정 탭의 값을 읽어 그리드에 채운다.
  const loadTab = useCallback(async (id: string, title: string) => {
    setError(null);
    try {
      const res = await fetch(
        `/api/sheet?spreadsheetId=${encodeURIComponent(id)}&sheet=${encodeURIComponent(title)}&validations=1`,
      );
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "시트를 불러오지 못했습니다.");

      const values: string[][] = data.values ?? [];
      const maxCols = values.reduce((m, r) => Math.max(m, r.length), 0);
      const rows: RowShape[] = values.map((r, ri) => {
        const obj: RowShape = { __row: ri + 1 };
        for (let i = 0; i < maxCols; i++) obj[columnLetter(i)] = r[i] ?? "";
        return obj;
      });

      // 탭을 바꾸면 이전 탭의 dirty 상태 초기화
      pendingRef.current.clear();
      lastWrittenRef.current.clear();
      editingRef.current.clear();
      colCountRef.current = maxCols;
      validationsRef.current = data.validations ?? {};

      setColumnDefs(buildColumnDefs(maxCols, validationsRef.current));
      setRowData(rows);
      setActiveTab(title);
      setSaveStatus("idle");
    } catch (e) {
      setError(e instanceof Error ? e.message : "알 수 없는 오류가 발생했습니다.");
    }
  }, []);

  // 시트 링크 입력 → 탭 목록 조회 후 첫 탭 로드
  const loadSheet = useCallback(async () => {
    setError(null);
    const id = extractSpreadsheetId(sheetInput);
    if (!id) {
      setError("시트 링크 또는 ID를 입력해 주세요.");
      return;
    }
    setLoading(true);
    try {
      const res = await fetch(`/api/sheet/meta?spreadsheetId=${encodeURIComponent(id)}`);
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "시트 정보를 불러오지 못했습니다.");

      const all: SheetTab[] = data.sheets ?? [];
      const visible = all.filter((s) => !s.hidden);
      if (visible.length === 0) throw new Error("열 수 있는 시트 탭이 없습니다.");

      setTabs(visible);
      setLoadedId(id);
      await loadTab(id, visible[0].title);
    } catch (e) {
      setError(e instanceof Error ? e.message : "알 수 없는 오류가 발생했습니다.");
      setLoadedId(null);
      setTabs([]);
      setActiveTab(null);
      setRowData([]);
      setColumnDefs([]);
    } finally {
      setLoading(false);
    }
  }, [sheetInput, loadTab]);

  // ---- 저장 (웹 → 시트) ----
  const flushSaves = useCallback(async () => {
    if (!loadedId || !activeTab) return;
    const entries = Array.from(pendingRef.current.entries());
    if (entries.length === 0) return;
    pendingRef.current.clear();

    setSaveStatus("saving");
    try {
      const res = await fetch("/api/sheet", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          spreadsheetId: loadedId,
          sheet: activeTab,
          updates: entries.map(([range, value]) => ({ range, value })),
        }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "저장에 실패했습니다.");
      // 수식(=...)이 포함됐으면: 쓴 값(수식) ≠ 읽는 값(계산결과)이라 보호 로직이
      // 막아버리므로, 탭을 다시 읽어 계산 결과로 반영한다.
      const hasFormula = entries.some(([, v]) => v.trimStart().startsWith("="));
      if (hasFormula) {
        await loadTab(loadedId, activeTab);
      } else {
        // 일반 값 → 폴링이 이 값을 확인할 때까지 보호
        for (const [range, value] of entries) lastWrittenRef.current.set(range, value);
      }
      setSaveStatus("saved");
    } catch (e) {
      setSaveStatus("error");
      setError(e instanceof Error ? e.message : "저장 중 오류가 발생했습니다.");
    }
  }, [loadedId, activeTab, loadTab]);

  const onCellValueChanged = useCallback(
    (e: CellValueChangedEvent) => {
      const field = e.colDef.field;
      if (!field || field === "__row") return;
      const row = (e.data as RowShape).__row;
      const value = e.newValue == null ? "" : String(e.newValue);
      pendingRef.current.set(`${field}${row}`, value);

      setSaveStatus("saving");
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
      saveTimerRef.current = setTimeout(() => void flushSaves(), 700);
    },
    [flushSaves],
  );

  const onCellEditingStarted = useCallback((e: CellEditingStartedEvent) => {
    const field = e.colDef.field;
    if (field && field !== "__row") {
      editingRef.current.add(`${field}${(e.data as RowShape).__row}`);
    }
  }, []);

  const onCellEditingStopped = useCallback((e: CellEditingStoppedEvent) => {
    const field = e.colDef.field;
    if (field && field !== "__row") {
      editingRef.current.delete(`${field}${(e.data as RowShape).__row}`);
    }
  }, []);

  // 탭 전환: 떠나는 탭의 미저장분을 먼저 저장하고 새 탭 로드
  const switchTab = useCallback(
    async (title: string) => {
      if (!loadedId || title === activeTab) return;
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
      setLoading(true);
      await flushSaves();
      await loadTab(loadedId, title);
      setLoading(false);
    },
    [loadedId, activeTab, flushSaves, loadTab],
  );

  // ---- 행/열 추가 (값 입력 시 시트에 생성됨) ----
  const addRow = useCallback(() => {
    const api = gridApiRef.current;
    if (!api) return;
    let maxRow = 0;
    api.forEachNode((n) => {
      const v = (n.data as RowShape).__row as number;
      if (v > maxRow) maxRow = v;
    });
    const obj: RowShape = { __row: maxRow + 1 };
    for (let i = 0; i < colCountRef.current; i++) obj[columnLetter(i)] = "";
    api.applyTransaction({ add: [obj] });
  }, []);

  const addColumn = useCallback(() => {
    colCountRef.current += 1;
    setColumnDefs(buildColumnDefs(colCountRef.current, validationsRef.current));
  }, []);

  // 공통: 행/열 삭제 요청 후 해당 탭을 다시 로드(번호가 밀리므로 재동기화)
  const deleteDimension = useCallback(
    async (dimension: "ROWS" | "COLUMNS", indices: number[]) => {
      if (!loadedId || !activeTab) return;
      const tab = tabs.find((t) => t.title === activeTab);
      if (!tab) return;

      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
      await flushSaves();
      setSaveStatus("saving");
      try {
        const res = await fetch("/api/sheet/dimension", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            spreadsheetId: loadedId,
            sheetId: tab.sheetId,
            dimension,
            indices,
          }),
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || "삭제에 실패했습니다.");
        await loadTab(loadedId, activeTab);
        setSaveStatus("saved");
      } catch (e) {
        setSaveStatus("error");
        setError(e instanceof Error ? e.message : "삭제 중 오류가 발생했습니다.");
      }
    },
    [loadedId, activeTab, tabs, flushSaves, loadTab],
  );

  const deleteRows = useCallback(() => {
    const api = gridApiRef.current;
    if (!api) return;
    const nodes = api.getSelectedNodes();
    if (nodes.length === 0) {
      setError("삭제할 행의 왼쪽 체크박스를 먼저 선택해 주세요.");
      return;
    }
    const indices = nodes.map((n) => ((n.data as RowShape).__row as number) - 1);
    void deleteDimension("ROWS", indices);
  }, [deleteDimension]);

  const deleteColumn = useCallback(() => {
    const api = gridApiRef.current;
    if (!api) return;
    const cell = api.getFocusedCell();
    if (!cell) {
      setError("삭제할 열의 셀을 먼저 클릭해 주세요.");
      return;
    }
    const colId = cell.column.getColId();
    if (!/^[A-Z]+$/.test(colId)) {
      setError("삭제할 열의 데이터 셀을 클릭해 주세요.");
      return;
    }
    void deleteDimension("COLUMNS", [letterToIndex(colId)]);
  }, [deleteDimension]);

  // ---- AI 생성 ----
  const refreshTabs = useCallback(async (id: string): Promise<SheetTab[]> => {
    const res = await fetch(`/api/sheet/meta?spreadsheetId=${encodeURIComponent(id)}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "탭 목록을 불러오지 못했습니다.");
    const visible: SheetTab[] = (data.sheets ?? []).filter((s: SheetTab) => !s.hidden);
    setTabs(visible);
    return visible;
  }, []);

  const generateWithAI = useCallback(async () => {
    if (!loadedId || !aiPrompt.trim()) return;
    setAiLoading(true);
    setError(null);
    try {
      const res = await fetch("/api/sheet/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ spreadsheetId: loadedId, prompt: aiPrompt }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "AI 생성에 실패했습니다.");
      await refreshTabs(loadedId);
      const created: string[] = data.createdTabs ?? [];
      if (created.length > 0) await loadTab(loadedId, created[0]);
      setAiOpen(false);
      setAiPrompt("");
    } catch (e) {
      setError(e instanceof Error ? e.message : "AI 생성 중 오류가 발생했습니다.");
    } finally {
      setAiLoading(false);
    }
  }, [loadedId, aiPrompt, refreshTabs, loadTab]);

  // ---- 함수 블록: 포커스한 셀에 수식 삽입 ----
  const insertBlock = useCallback(
    (make: (col: string, row: number) => string) => {
      const api = gridApiRef.current;
      if (!api) return;
      const cell = api.getFocusedCell();
      if (!cell) {
        setError("먼저 수식을 넣을 셀을 클릭하세요.");
        return;
      }
      const colId = cell.column.getColId();
      if (!/^[A-Z]+$/.test(colId)) {
        setError("수식을 넣을 데이터 셀을 클릭하세요.");
        return;
      }
      const node = api.getDisplayedRowAtIndex(cell.rowIndex);
      if (!node) return;
      const sheetRow = (node.data as RowShape).__row as number;
      const formula = make(colId, sheetRow);
      const range = `${colId}${sheetRow}`;

      node.setDataValue(colId, formula); // 낙관적 표시(잠깐 수식 텍스트)
      pendingRef.current.set(range, formula);
      setError(null);
      setSaveStatus("saving");
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
      // 저장 후 flushSaves가 계산 결과로 다시 읽어온다(수식 경로).
      saveTimerRef.current = setTimeout(() => void flushSaves(), 300);
    },
    [flushSaves],
  );

  // ---- 입력 규칙(드롭다운/체크박스): 포커스 셀의 열에 적용 ----
  const applyValidation = useCallback(
    async (type: "list" | "checkbox" | "clear", values?: string[]) => {
      const api = gridApiRef.current;
      if (!api || !loadedId || !activeTab) return;
      const cell = api.getFocusedCell();
      if (!cell) {
        setError("규칙을 적용할 열의 셀을 먼저 클릭하세요.");
        return;
      }
      const colId = cell.column.getColId();
      if (!/^[A-Z]+$/.test(colId)) {
        setError("데이터 열의 셀을 클릭하세요.");
        return;
      }
      const tab = tabs.find((t) => t.title === activeTab);
      if (!tab) return;
      setError(null);
      setSaveStatus("saving");
      try {
        const res = await fetch("/api/sheet/validation", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            spreadsheetId: loadedId,
            sheetId: tab.sheetId,
            columnIndex: letterToIndex(colId),
            type,
            values,
          }),
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || "규칙 적용에 실패했습니다.");
        await loadTab(loadedId, activeTab); // 규칙을 다시 읽어 그리드에 반영
        setSaveStatus("saved");
      } catch (e) {
        setSaveStatus("error");
        setError(e instanceof Error ? e.message : "규칙 적용 중 오류가 발생했습니다.");
      }
    },
    [loadedId, activeTab, tabs, loadTab],
  );

  // ---- 시트 → 웹: 폴링 결과를 그리드에 반영 ----
  const applyRemoteValues = useCallback((values: string[][]) => {
    const api = gridApiRef.current;
    if (!api) return;

    const maxCols = values.reduce((m, r) => Math.max(m, r.length), 0);
    if (maxCols > colCountRef.current) {
      colCountRef.current = maxCols;
      setColumnDefs(buildColumnDefs(maxCols, validationsRef.current));
    }

    const addRows: RowShape[] = [];
    values.forEach((r, ri) => {
      const sheetRow = ri + 1;
      const node = api.getRowNode(String(sheetRow));

      if (!node) {
        const obj: RowShape = { __row: sheetRow };
        for (let i = 0; i < maxCols; i++) obj[columnLetter(i)] = r[i] ?? "";
        addRows.push(obj);
        return;
      }

      for (let i = 0; i < maxCols; i++) {
        const field = columnLetter(i);
        const range = `${field}${sheetRow}`;
        const incoming = r[i] ?? "";

        // 사용자가 편집 중이거나 저장 대기 중인 셀은 건드리지 않는다.
        if (pendingRef.current.has(range) || editingRef.current.has(range)) continue;

        // 방금 우리가 쓴 값: 폴링에 반영됐으면 보호 해제, 아니면 stale 이므로 건너뜀
        if (lastWrittenRef.current.has(range)) {
          if (lastWrittenRef.current.get(range) === incoming) {
            lastWrittenRef.current.delete(range);
          } else {
            continue;
          }
        }

        const data = node.data as RowShape;
        if (data[field] !== incoming) node.setDataValue(field, incoming);
      }
    });

    if (addRows.length > 0) api.applyTransaction({ add: addRows });
  }, []);

  // 폴링 루프 (활성 탭 기준, 백그라운드 탭이면 스킵)
  const applyRef = useRef(applyRemoteValues);
  useEffect(() => {
    applyRef.current = applyRemoteValues;
  }, [applyRemoteValues]);

  useEffect(() => {
    if (!loadedId || !activeTab) return;
    const tick = async () => {
      if (typeof document !== "undefined" && document.hidden) return;
      try {
        const res = await fetch(
          `/api/sheet?spreadsheetId=${encodeURIComponent(loadedId)}&sheet=${encodeURIComponent(activeTab)}`,
        );
        if (!res.ok) return;
        const data = await res.json();
        applyRef.current((data.values ?? []) as string[][]);
      } catch {
        /* 네트워크 일시 오류 무시 */
      }
    };
    const id = setInterval(tick, POLL_INTERVAL_MS);
    return () => clearInterval(id);
  }, [loadedId, activeTab]);

  // ---- 세션 로딩 중 ----
  if (status === "loading") {
    return (
      <main className="flex flex-1 items-center justify-center">
        <p className="text-gray-400">불러오는 중…</p>
      </main>
    );
  }

  // ---- 비로그인 화면 ----
  if (!session) {
    return (
      <main className="flex flex-1 items-center justify-center px-4">
        <div className="w-full max-w-sm rounded-2xl bg-white p-8 text-center shadow-sm ring-1 ring-gray-100">
          <h1 className="text-2xl font-bold tracking-tight">Sheet Helper</h1>
          <p className="mt-2 text-sm text-gray-500">구글 시트를 웹에서 쉽고 간편하게.</p>
          <button
            onClick={() => signIn("google")}
            className="mt-6 w-full rounded-xl bg-gray-900 px-4 py-3 text-sm font-semibold text-white transition hover:bg-gray-700"
          >
            Google 계정으로 시작하기
          </button>
        </div>
      </main>
    );
  }

  // ---- 로그인 후 화면 ----
  return (
    <div className="flex flex-1 flex-col">
      <header className="flex items-center justify-between border-b border-gray-200 bg-white px-6 py-3">
        <div className="flex items-center gap-3">
          <h1 className="text-lg font-bold tracking-tight">Sheet Helper</h1>
          <SaveBadge status={saveStatus} />
          {loadedId && (
            <span className="flex items-center gap-1.5 text-xs text-gray-400">
              <span className="h-1.5 w-1.5 animate-pulse rounded-full bg-green-500" />
              실시간 동기화 중
            </span>
          )}
        </div>
        <div className="flex items-center gap-3 text-sm">
          <span className="text-gray-500">{session.user?.email}</span>
          <button
            onClick={() => signOut()}
            className="rounded-lg px-3 py-1.5 text-gray-600 ring-1 ring-gray-200 transition hover:bg-gray-50"
          >
            로그아웃
          </button>
        </div>
      </header>

      <section className="border-b border-gray-200 bg-white px-6 py-4">
        <label className="mb-1.5 block text-xs font-medium text-gray-500">
          구글 시트 링크 또는 ID
        </label>
        <div className="flex gap-2">
          <input
            value={sheetInput}
            onChange={(e) => setSheetInput(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && loadSheet()}
            placeholder="https://docs.google.com/spreadsheets/d/..."
            className="flex-1 rounded-xl border border-gray-300 px-4 py-2.5 text-sm outline-none transition focus:border-gray-900 focus:ring-2 focus:ring-gray-900/10"
          />
          <button
            onClick={loadSheet}
            disabled={loading}
            className="rounded-xl bg-gray-900 px-5 py-2.5 text-sm font-semibold text-white transition hover:bg-gray-700 disabled:opacity-50"
          >
            {loading ? "불러오는 중…" : "불러오기"}
          </button>
        </div>
        {error && <p className="mt-2 text-sm text-red-600">{error}</p>}
      </section>

      <main className="flex flex-1 flex-col p-6">
        {loadedId ? (
          <>
            {/* 탭 바 + 도구 */}
            <div className="mb-2 flex items-end justify-between gap-3">
              <div className="flex gap-1 overflow-x-auto">
                {tabs.map((t) => (
                  <button
                    key={t.sheetId}
                    onClick={() => switchTab(t.title)}
                    className={`whitespace-nowrap rounded-t-lg px-3 py-1.5 text-sm transition ${
                      activeTab === t.title
                        ? "bg-white font-semibold text-gray-900 ring-1 ring-gray-200"
                        : "text-gray-500 hover:bg-gray-100"
                    }`}
                  >
                    {t.title}
                  </button>
                ))}
              </div>
              <div className="flex shrink-0 gap-2">
                <button
                  onClick={() => setAiOpen((v) => !v)}
                  className="rounded-lg bg-indigo-600 px-3 py-1.5 text-sm font-semibold text-white transition hover:bg-indigo-500"
                >
                  ✨ AI로 만들기
                </button>
                <button
                  onClick={() => setBlockOpen((v) => !v)}
                  className="rounded-lg bg-emerald-600 px-3 py-1.5 text-sm font-semibold text-white transition hover:bg-emerald-500"
                >
                  🧱 함수 블록
                </button>
                <button
                  onClick={() => setChartOpen(true)}
                  className="rounded-lg bg-sky-600 px-3 py-1.5 text-sm font-semibold text-white transition hover:bg-sky-500"
                >
                  📊 차트
                </button>
                <button
                  onClick={() => setRuleOpen((v) => !v)}
                  className="rounded-lg bg-amber-500 px-3 py-1.5 text-sm font-semibold text-white transition hover:bg-amber-400"
                >
                  ✓ 입력 규칙
                </button>
                <button
                  onClick={addRow}
                  className="rounded-lg px-3 py-1.5 text-sm text-gray-700 ring-1 ring-gray-200 transition hover:bg-gray-50"
                >
                  + 행
                </button>
                <button
                  onClick={addColumn}
                  className="rounded-lg px-3 py-1.5 text-sm text-gray-700 ring-1 ring-gray-200 transition hover:bg-gray-50"
                >
                  + 열
                </button>
                <button
                  onClick={deleteRows}
                  className="rounded-lg px-3 py-1.5 text-sm text-red-600 ring-1 ring-red-200 transition hover:bg-red-50"
                >
                  행 삭제
                </button>
                <button
                  onClick={deleteColumn}
                  className="rounded-lg px-3 py-1.5 text-sm text-red-600 ring-1 ring-red-200 transition hover:bg-red-50"
                >
                  열 삭제
                </button>
              </div>
            </div>

            {blockOpen && (
              <div className="mb-3 rounded-xl border border-emerald-100 bg-emerald-50/50 p-4">
                <p className="mb-2 text-xs font-medium text-emerald-900">
                  넣을 셀을 클릭한 뒤 블록을 누르면, 함수를 몰라도 수식이 자동으로 들어가요.
                </p>
                <div className="flex flex-wrap gap-1.5">
                  {FORMULA_BLOCKS.map((b) => (
                    <button
                      key={b.key}
                      onClick={() => insertBlock(b.make)}
                      title={b.desc}
                      className="rounded-lg bg-white px-3 py-1.5 text-sm text-emerald-800 ring-1 ring-emerald-200 transition hover:bg-emerald-100"
                    >
                      {b.label}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {aiOpen && (
              <div className="mb-3 rounded-xl border border-indigo-100 bg-indigo-50/50 p-4">
                <label className="mb-1.5 block text-xs font-medium text-indigo-900">
                  무엇을 만들까요? (자연어로 입력하면 AI가 시트를 새 탭으로 만들어줘요)
                </label>
                <textarea
                  value={aiPrompt}
                  onChange={(e) => setAiPrompt(e.target.value)}
                  rows={2}
                  placeholder="예: 쇼핑몰 매출 관리 시스템 만들어줘"
                  className="w-full resize-none rounded-lg border border-indigo-200 px-3 py-2 text-sm outline-none transition focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20"
                />
                <div className="mt-2 flex flex-wrap gap-1.5">
                  {["쇼핑몰 매출 관리 시스템", "가계부 (수입·지출·예산)", "재고 관리", "영업 CRM"].map(
                    (ex) => (
                      <button
                        key={ex}
                        onClick={() => setAiPrompt(ex)}
                        className="rounded-full bg-white px-2.5 py-1 text-xs text-indigo-700 ring-1 ring-indigo-200 transition hover:bg-indigo-100"
                      >
                        {ex}
                      </button>
                    ),
                  )}
                </div>
                <div className="mt-3 flex items-center gap-2">
                  <button
                    onClick={generateWithAI}
                    disabled={aiLoading || !aiPrompt.trim()}
                    className="rounded-lg bg-indigo-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-indigo-500 disabled:opacity-50"
                  >
                    {aiLoading ? "AI가 만들고 있어요…" : "생성"}
                  </button>
                  <span className="text-xs text-indigo-700/70">
                    새 탭으로 추가됩니다. 수십 초 걸릴 수 있어요.
                  </span>
                </div>
              </div>
            )}

            {ruleOpen && (
              <div className="mb-3 rounded-xl border border-amber-100 bg-amber-50/50 p-4">
                <p className="mb-2 text-xs font-medium text-amber-900">
                  규칙을 적용할 열의 셀을 클릭한 뒤 선택하세요. 구글 시트에도 그대로 적용돼요.
                </p>
                <div className="flex flex-wrap items-center gap-2">
                  <input
                    value={dropdownOptions}
                    onChange={(e) => setDropdownOptions(e.target.value)}
                    placeholder="드롭다운 항목: 예) 진행중, 완료, 보류"
                    className="min-w-[200px] flex-1 rounded-lg border border-amber-200 px-3 py-1.5 text-sm outline-none transition focus:border-amber-500"
                  />
                  <button
                    onClick={() => {
                      const vals = dropdownOptions
                        .split(",")
                        .map((s) => s.trim())
                        .filter(Boolean);
                      if (vals.length === 0) {
                        setError("드롭다운 항목을 콤마(,)로 구분해 입력하세요.");
                        return;
                      }
                      void applyValidation("list", vals);
                    }}
                    className="rounded-lg bg-amber-500 px-3 py-1.5 text-sm font-semibold text-white transition hover:bg-amber-400"
                  >
                    드롭다운 적용
                  </button>
                  <button
                    onClick={() => void applyValidation("checkbox")}
                    className="rounded-lg px-3 py-1.5 text-sm text-amber-800 ring-1 ring-amber-200 transition hover:bg-amber-100"
                  >
                    체크박스
                  </button>
                  <button
                    onClick={() => void applyValidation("clear")}
                    className="rounded-lg px-3 py-1.5 text-sm text-gray-600 ring-1 ring-gray-200 transition hover:bg-gray-50"
                  >
                    규칙 해제
                  </button>
                </div>
              </div>
            )}

            <p className="mb-2 text-xs text-gray-400">
              셀을 더블클릭해 수정하면 자동 저장돼요. 헤더로 정렬·필터할 수 있고, 시트에서 바뀐 내용도 5초 안에 반영됩니다.
            </p>
            <div className="h-[calc(100vh-290px)] w-full overflow-hidden rounded-xl ring-1 ring-gray-200">
              <AgGridReact
                rowData={rowData}
                columnDefs={columnDefs}
                defaultColDef={defaultColDef}
                rowSelection={rowSelection}
                getRowId={(p) => String((p.data as RowShape).__row)}
                onGridReady={onGridReady}
                onCellValueChanged={onCellValueChanged}
                onCellEditingStarted={onCellEditingStarted}
                onCellEditingStopped={onCellEditingStopped}
                stopEditingWhenCellsLoseFocus
              />
            </div>
          </>
        ) : (
          <div className="flex flex-1 items-center justify-center text-center text-gray-400">
            <div>
              <p className="text-sm">
                위에 구글 시트 링크를 붙여넣고 <b>불러오기</b>를 눌러보세요.
              </p>
              <p className="mt-1 text-xs">본인 계정으로 접근 가능한 시트만 열 수 있어요.</p>
            </div>
          </div>
        )}
      </main>

      {chartOpen && loadedId && (
        <ChartModal
          rowData={rowData}
          columnDefs={columnDefs}
          onClose={() => setChartOpen(false)}
        />
      )}
    </div>
  );
}

/** 저장 상태 배지 */
function SaveBadge({ status }: { status: SaveStatus }) {
  if (status === "idle") return null;
  const config = {
    saving: { text: "저장 중…", className: "bg-gray-100 text-gray-500" },
    saved: { text: "저장됨 ✓", className: "bg-green-50 text-green-600" },
    error: { text: "저장 실패", className: "bg-red-50 text-red-600" },
  }[status];
  return (
    <span className={`rounded-full px-2.5 py-0.5 text-xs font-medium ${config.className}`}>
      {config.text}
    </span>
  );
}

/** 차트 / 대시보드 모달 — 현재 탭 데이터를 읽어 KPI + 차트로 시각화 */
function ChartModal({
  rowData,
  columnDefs,
  onClose,
}: {
  rowData: RowShape[];
  columnDefs: ColDef[];
  onClose: () => void;
}) {
  const dataCols = useMemo(
    () =>
      columnDefs
        .map((c) => c.field)
        .filter((f): f is string => !!f && f !== "__row"),
    [columnDefs],
  );
  const headerRow = useMemo(() => rowData.find((r) => r.__row === 1), [rowData]);
  const headerLabel = (col: string) => {
    const h = headerRow?.[col];
    return h !== undefined && String(h).trim() !== "" ? `${col} · ${h}` : col;
  };

  const [chartType, setChartType] = useState<"bar" | "line" | "pie">("bar");
  const [groupCol, setGroupCol] = useState(dataCols[0] ?? "A");
  const [valueCol, setValueCol] = useState(dataCols[1] ?? dataCols[0] ?? "B");
  const [agg, setAgg] = useState<Agg>("sum");

  const aggLabel = AGG_OPTIONS.find((o) => o.key === agg)?.label ?? "";
  const fmt = (n: number) => n.toLocaleString("ko-KR", { maximumFractionDigits: 1 });

  // 그룹 기준 열로 묶어서 값 열을 집계 → 의미 있는 요약 차트
  const chartData = useMemo(() => {
    const parseNum = (raw: string | number | undefined) =>
      typeof raw === "number" ? raw : parseFloat(String(raw ?? "").replace(/,/g, ""));
    const map = new Map<string, { sum: number; count: number; max: number; min: number }>();
    for (const r of rowData) {
      if (r.__row === 1) continue; // 1행은 머리글
      const key = String(r[groupCol] ?? "").trim() || "(빈값)";
      let acc = map.get(key);
      if (!acc) {
        acc = { sum: 0, count: 0, max: -Infinity, min: Infinity };
        map.set(key, acc);
      }
      if (agg === "count") {
        acc.count += 1;
      } else {
        const v = parseNum(r[valueCol]);
        if (!Number.isFinite(v)) continue;
        acc.sum += v;
        acc.count += 1;
        acc.max = Math.max(acc.max, v);
        acc.min = Math.min(acc.min, v);
      }
    }
    const arr = [...map.entries()]
      .map(([name, a]) => {
        let value: number;
        if (agg === "count") value = a.count;
        else if (agg === "sum") value = a.sum;
        else if (agg === "avg") value = a.count ? a.sum / a.count : 0;
        else if (agg === "max") value = a.count ? a.max : 0;
        else value = a.count ? a.min : 0;
        return { name, value };
      })
      .filter((d) => (agg === "count" ? d.value > 0 : Number.isFinite(d.value)));
    if (chartType !== "line") arr.sort((x, y) => y.value - x.value); // 값 큰 순(랭킹)
    return arr.slice(0, 30);
  }, [rowData, groupCol, valueCol, agg, chartType]);

  const kpis = useMemo(() => {
    if (chartData.length === 0) return null;
    const vals = chartData.map((d) => d.value);
    const sum = vals.reduce((a, b) => a + b, 0);
    const count = vals.length;
    const sorted = [...chartData].sort((a, b) => b.value - a.value);
    return { count, sum, avg: sum / count, top: sorted[0], bottom: sorted[sorted.length - 1] };
  }, [chartData]);

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      onClick={onClose}
    >
      <div
        className="max-h-[90vh] w-full max-w-3xl overflow-auto rounded-2xl bg-white p-6 shadow-xl"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="mb-4 flex items-center justify-between">
          <h2 className="text-lg font-bold tracking-tight">차트 / 대시보드</h2>
          <button
            onClick={onClose}
            className="rounded-lg px-3 py-1.5 text-sm text-gray-600 ring-1 ring-gray-200 transition hover:bg-gray-50"
          >
            닫기
          </button>
        </div>

        {/* 설정 */}
        <div className="mb-4 flex flex-wrap items-end gap-3">
          <div className="flex gap-1">
            {(["bar", "line", "pie"] as const).map((t) => (
              <button
                key={t}
                onClick={() => setChartType(t)}
                className={`rounded-lg px-3 py-1.5 text-sm transition ${
                  chartType === t
                    ? "bg-gray-900 font-semibold text-white"
                    : "text-gray-600 ring-1 ring-gray-200 hover:bg-gray-50"
                }`}
              >
                {t === "bar" ? "막대" : t === "line" ? "선" : "원형"}
              </button>
            ))}
          </div>
          <label className="text-xs text-gray-500">
            그룹 기준
            <select
              value={groupCol}
              onChange={(e) => setGroupCol(e.target.value)}
              className="ml-1 rounded-lg border border-gray-300 px-2 py-1 text-sm text-gray-900"
            >
              {dataCols.map((c) => (
                <option key={c} value={c}>
                  {headerLabel(c)}
                </option>
              ))}
            </select>
          </label>
          <label className="text-xs text-gray-500">
            집계
            <select
              value={agg}
              onChange={(e) => setAgg(e.target.value as Agg)}
              className="ml-1 rounded-lg border border-gray-300 px-2 py-1 text-sm text-gray-900"
            >
              {AGG_OPTIONS.map((o) => (
                <option key={o.key} value={o.key}>
                  {o.label}
                </option>
              ))}
            </select>
          </label>
          <label className={`text-xs ${agg === "count" ? "text-gray-300" : "text-gray-500"}`}>
            값
            <select
              value={valueCol}
              onChange={(e) => setValueCol(e.target.value)}
              disabled={agg === "count"}
              className="ml-1 rounded-lg border border-gray-300 px-2 py-1 text-sm text-gray-900 disabled:bg-gray-100 disabled:text-gray-400"
            >
              {dataCols.map((c) => (
                <option key={c} value={c}>
                  {headerLabel(c)}
                </option>
              ))}
            </select>
          </label>
        </div>

        <p className="mb-3 text-xs text-gray-500">
          「{headerLabel(groupCol)}」별{" "}
          {agg === "count" ? "개수" : `「${headerLabel(valueCol)}」 ${aggLabel}`}
        </p>

        {chartData.length === 0 ? (
          <p className="py-12 text-center text-sm text-gray-400">
            표시할 데이터가 없어요. 그룹 기준·집계·값을 확인해 주세요. (1행은 머리글, 2행부터
            데이터로 읽어요.)
          </p>
        ) : (
          <>
            {/* KPI 카드 (그룹 집계 기준) */}
            {kpis && (
              <div className="mb-4 grid grid-cols-2 gap-2 sm:grid-cols-5">
                {[
                  { label: "그룹 수", value: fmt(kpis.count), sub: "" },
                  { label: "합계", value: fmt(kpis.sum), sub: "" },
                  { label: "평균", value: fmt(kpis.avg), sub: "" },
                  { label: "최대", value: fmt(kpis.top.value), sub: kpis.top.name },
                  { label: "최소", value: fmt(kpis.bottom.value), sub: kpis.bottom.name },
                ].map((k) => (
                  <div key={k.label} className="rounded-xl bg-gray-50 p-3 text-center">
                    <div className="text-xs text-gray-500">{k.label}</div>
                    <div className="mt-0.5 text-lg font-bold text-gray-900">{k.value}</div>
                    {k.sub && <div className="truncate text-[11px] text-gray-400">{k.sub}</div>}
                  </div>
                ))}
              </div>
            )}

            {/* 차트 */}
            <div className="h-80 w-full">
              <ResponsiveContainer width="100%" height="100%">
                {chartType === "bar" ? (
                  <BarChart data={chartData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                    <XAxis dataKey="name" tick={{ fontSize: 12 }} />
                    <YAxis tick={{ fontSize: 12 }} />
                    <Tooltip />
                    <Bar dataKey="value" fill="#6366f1" radius={[4, 4, 0, 0]} />
                  </BarChart>
                ) : chartType === "line" ? (
                  <LineChart data={chartData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#eee" />
                    <XAxis dataKey="name" tick={{ fontSize: 12 }} />
                    <YAxis tick={{ fontSize: 12 }} />
                    <Tooltip />
                    <Line type="monotone" dataKey="value" stroke="#6366f1" strokeWidth={2} />
                  </LineChart>
                ) : (
                  <PieChart>
                    <Tooltip />
                    <Legend />
                    <Pie data={chartData} dataKey="value" nameKey="name" outerRadius={110} label>
                      {chartData.map((_, i) => (
                        <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />
                      ))}
                    </Pie>
                  </PieChart>
                )}
              </ResponsiveContainer>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
