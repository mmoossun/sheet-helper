import { NextResponse, type NextRequest } from "next/server";
import { getSheetsClient, toFriendlyError, buildRange } from "@/lib/sheets";

// 세션(쿠키)에 의존하므로 항상 동적으로 처리하고 Node 런타임을 사용한다.
export const dynamic = "force-dynamic";
export const runtime = "nodejs";

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
    return NextResponse.json({ values: res.data.values ?? [], range: res.data.range });
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
