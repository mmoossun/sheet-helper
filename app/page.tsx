"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { useSession, signIn, signOut } from "next-auth/react";
import { AgGridReact } from "ag-grid-react";
import {
  AllCommunityModule,
  ModuleRegistry,
  type ColDef,
  type CellValueChangedEvent,
} from "ag-grid-community";

// AG Grid v33+ 는 모듈 등록이 필요하다.
ModuleRegistry.registerModules([AllCommunityModule]);

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

/** 전체 URL을 붙여넣어도 시트 ID만 뽑아낸다. */
function extractSpreadsheetId(input: string): string {
  const trimmed = input.trim();
  const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : trimmed;
}

type RowShape = Record<string, string | number>;
type SaveStatus = "idle" | "saving" | "saved" | "error";

export default function Home() {
  const { data: session, status } = useSession();
  const [sheetInput, setSheetInput] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [rowData, setRowData] = useState<RowShape[]>([]);
  const [columnDefs, setColumnDefs] = useState<ColDef[]>([]);
  const [loadedId, setLoadedId] = useState<string | null>(null);
  const [saveStatus, setSaveStatus] = useState<SaveStatus>("idle");

  // 저장 대기열(셀 범위 → 값)과 디바운스 타이머
  const pendingRef = useRef<Map<string, string>>(new Map());
  const timerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const defaultColDef = useMemo<ColDef>(
    () => ({ resizable: true, sortable: false, minWidth: 90, editable: true }),
    [],
  );

  // 세션이 만료되면(리프레시 실패) 재로그인 유도
  useEffect(() => {
    if (session?.error === "RefreshAccessTokenError") {
      signIn("google");
    }
  }, [session?.error]);

  const loadSheet = useCallback(async () => {
    setError(null);
    const id = extractSpreadsheetId(sheetInput);
    if (!id) {
      setError("시트 링크 또는 ID를 입력해 주세요.");
      return;
    }
    setLoading(true);
    try {
      const res = await fetch(
        `/api/sheet?spreadsheetId=${encodeURIComponent(id)}`,
      );
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "시트를 불러오지 못했습니다.");

      const values: string[][] = data.values ?? [];
      const maxCols = values.reduce((m, r) => Math.max(m, r.length), 0);

      const cols: ColDef[] = [
        {
          headerName: "",
          field: "__row",
          width: 56,
          pinned: "left",
          sortable: false,
          editable: false,
          cellClass: "ag-row-number",
        },
        ...Array.from({ length: Math.max(maxCols, 1) }, (_, i) => ({
          headerName: columnLetter(i),
          field: columnLetter(i),
          flex: 1,
        })),
      ];

      const rows: RowShape[] = values.map((r, ri) => {
        const obj: RowShape = { __row: ri + 1 };
        for (let i = 0; i < maxCols; i++) obj[columnLetter(i)] = r[i] ?? "";
        return obj;
      });

      setColumnDefs(cols);
      setRowData(rows);
      setLoadedId(id);
      setSaveStatus("idle");
    } catch (e) {
      setError(e instanceof Error ? e.message : "알 수 없는 오류가 발생했습니다.");
      setRowData([]);
      setColumnDefs([]);
      setLoadedId(null);
    } finally {
      setLoading(false);
    }
  }, [sheetInput]);

  // 대기열에 쌓인 변경분을 시트에 저장
  const flushSaves = useCallback(async () => {
    if (!loadedId) return;
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
          updates: entries.map(([range, value]) => ({ range, value })),
        }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "저장에 실패했습니다.");
      setSaveStatus("saved");
    } catch (e) {
      setSaveStatus("error");
      setError(e instanceof Error ? e.message : "저장 중 오류가 발생했습니다.");
    }
  }, [loadedId]);

  // 셀 편집 → 낙관적 반영(그리드가 즉시 표시) + 디바운스 저장
  const onCellValueChanged = useCallback(
    (e: CellValueChangedEvent) => {
      const field = e.colDef.field;
      if (!field || field === "__row") return;
      const row = (e.data as RowShape).__row;
      const value = e.newValue == null ? "" : String(e.newValue);
      pendingRef.current.set(`${field}${row}`, value);

      setSaveStatus("saving");
      if (timerRef.current) clearTimeout(timerRef.current);
      timerRef.current = setTimeout(() => {
        void flushSaves();
      }, 700);
    },
    [flushSaves],
  );

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
          <p className="mt-2 text-sm text-gray-500">
            구글 시트를 웹에서 쉽고 간편하게.
          </p>
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
      {/* 상단 바 */}
      <header className="flex items-center justify-between border-b border-gray-200 bg-white px-6 py-3">
        <div className="flex items-center gap-3">
          <h1 className="text-lg font-bold tracking-tight">Sheet Helper</h1>
          <SaveBadge status={saveStatus} />
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

      {/* 시트 불러오기 입력 */}
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

      {/* 그리드 */}
      <main className="flex-1 p-6">
        {loadedId ? (
          <>
            <p className="mb-2 text-xs text-gray-400">
              셀을 더블클릭해 수정하면 구글 시트에 자동 저장돼요.
            </p>
            <div className="h-[calc(100vh-250px)] w-full overflow-hidden rounded-xl ring-1 ring-gray-200">
              <AgGridReact
                rowData={rowData}
                columnDefs={columnDefs}
                defaultColDef={defaultColDef}
                onCellValueChanged={onCellValueChanged}
                stopEditingWhenCellsLoseFocus
              />
            </div>
          </>
        ) : (
          <div className="flex h-full items-center justify-center text-center text-gray-400">
            <div>
              <p className="text-sm">
                위에 구글 시트 링크를 붙여넣고 <b>불러오기</b>를 눌러보세요.
              </p>
              <p className="mt-1 text-xs">
                본인 계정으로 접근 가능한 시트만 열 수 있어요.
              </p>
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
