import { NextResponse, type NextRequest } from "next/server";
import type { sheets_v4 } from "googleapis";
import { getSheetsClient, toFriendlyError, buildRange } from "@/lib/sheets";

// 세션(쿠키)에 의존하므로 항상 동적으로 처리하고 Node 런타임을 사용한다.
export const dynamic = "force-dynamic";
export const runtime = "nodejs";

type ColValidation = { type: "list" | "checkbox"; values?: string[] };

/** 0 → A, 1 → B, ... 26 → AA */
function colLetter(index: number): string {
  let s = "";
  let i = index + 1;
  while (i > 0) {
    const m = (i - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    i = Math.floor((i - 1) / 26);
  }
  return s;
}

/** 데이터 검증을 읽어 "열 문자 → 규칙" 맵으로 요약한다 (2행 셀을 대표로 사용). */
async function readValidations(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  range: string,
): Promise<Record<string, ColValidation>> {
  const out: Record<string, ColValidation> = {};
  try {
    const meta = await sheets.spreadsheets.get({
      spreadsheetId,
      ranges: [range],
      fields: "sheets(data(rowData(values(dataValidation))))",
    });
    const rows = meta.data.sheets?.[0]?.data?.[0]?.rowData ?? [];
    const sample = rows[1]?.values ?? []; // 2행(첫 데이터 행) 기준
    sample.forEach((cell, idx) => {
      const cond = cell.dataValidation?.condition;
      if (!cond?.type) return;
      const letter = colLetter(idx);
      if (cond.type === "ONE_OF_LIST") {
        out[letter] = {
          type: "list",
          values: (cond.values ?? [])
            .map((v) => v.userEnteredValue ?? "")
            .filter((v) => v !== ""),
        };
      } else if (cond.type === "BOOLEAN") {
        out[letter] = { type: "checkbox" };
      }
    });
  } catch {
    // 검증 읽기 실패는 조용히 무시
  }
  return out;
}

/** 시트 값 읽기 (특정 탭) */
export async function GET(req: NextRequest) {
  const client = await getSheetsClient();
  if (client.error) return client.error;

  const { searchParams } = new URL(req.url);
  const spreadsheetId = searchParams.get("spreadsheetId");
  const sheet = searchParams.get("sheet"); // 탭 제목 (옵션, 없으면 첫 탭)
  const a1 = searchParams.get("range") || "A1:Z1000";

  if (!spreadsheetId) {
    return NextResponse.json({ error: "spreadsheetId가 필요합니다." }, { status: 400 });
  }

  try {
    const res = await client.sheets.spreadsheets.values.get({
      spreadsheetId,
      range: buildRange(sheet, a1),
    });

    // validations=1 일 때만 데이터 검증을 함께 읽는다(폴링 때는 생략해 가볍게).
    let validations: Record<string, ColValidation> | undefined;
    if (searchParams.get("validations")) {
      validations = await readValidations(client.sheets, spreadsheetId, buildRange(sheet, a1));
    }

    return NextResponse.json({
      values: res.data.values ?? [],
      range: res.data.range,
      validations,
    });
  } catch (err) {
    const { message, status } = toFriendlyError(err);
    return NextResponse.json({ error: message }, { status });
  }
}

type CellUpdate = { range: string; value: string };

/** 셀 값 저장 (특정 탭, 여러 셀을 한 번에 batchUpdate) */
export async function POST(req: NextRequest) {
  const client = await getSheetsClient();
  if (client.error) return client.error;

  let body: { spreadsheetId?: string; sheet?: string; updates?: CellUpdate[] };
  try {
    body = await req.json();
  } catch {
    return NextResponse.json({ error: "잘못된 요청 형식입니다." }, { status: 400 });
  }

  const { spreadsheetId, sheet, updates } = body;
  if (!spreadsheetId) {
    return NextResponse.json({ error: "spreadsheetId가 필요합니다." }, { status: 400 });
  }
  if (!Array.isArray(updates) || updates.length === 0) {
    return NextResponse.json({ error: "저장할 변경 내용이 없습니다." }, { status: 400 });
  }

  try {
    await client.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        // 사용자가 직접 입력한 것처럼 숫자/수식/날짜를 해석한다.
        valueInputOption: "USER_ENTERED",
        data: updates.map((u) => ({ range: buildRange(sheet, u.range), values: [[u.value]] })),
      },
    });
    return NextResponse.json({ ok: true, saved: updates.length });
  } catch (err) {
    const { message, status } = toFriendlyError(err);
    return NextResponse.json({ error: message }, { status });
  }
}
