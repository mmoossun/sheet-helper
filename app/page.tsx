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

/** 데이터 열 개수에 맞춰 컬럼 정의를 만든다 (맨 앞에 행번호 열). */
function buildColumnDefs(maxCols: number): ColDef[] {
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
    ...Array.from({ length: Math.max(maxCols, 1) }, (_, i) => ({
      headerName: columnLetter(i),
      field: columnLetter(i),
      flex: 1,
    })),
  ];
}

type RowShape = Record<string, string | number>;
type SaveStatus = "idle" | "saving" | "saved" | "error";
type SheetTab = { sheetId: number; title: string; index: number; hidden: boolean };

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

  const gridApiRef = useRef<GridApi | null>(null);
  const colCountRef = useRef(0); // 현재 데이터 열 개수(행번호 열 제외)

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
        `/api/sheet?spreadsheetId=${encodeURIComponent(id)}&sheet=${encodeURIComponent(title)}`,
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

      setColumnDefs(buildColumnDefs(maxCols));
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
      // 저장 성공 → 폴링이 이 값을 확인할 때까지 보호
      for (const [range, value] of entries) lastWrittenRef.current.set(range, value);
      setSaveStatus("saved");
    } catch (e) {
      setSaveStatus("error");
      setError(e instanceof Error ? e.message : "저장 중 오류가 발생했습니다.");
    }
  }, [loadedId, activeTab]);

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
    setColumnDefs(buildColumnDefs(colCountRef.current));
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

  // ---- 시트 → 웹: 폴링 결과를 그리드에 반영 ----
  const applyRemoteValues = useCallback((values: string[][]) => {
    const api = gridApiRef.current;
    if (!api) return;

    const maxCols = values.reduce((m, r) => Math.max(m, r.length), 0);
    if (maxCols > colCountRef.current) {
      colCountRef.current = maxCols;
      setColumnDefs(buildColumnDefs(maxCols));
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
