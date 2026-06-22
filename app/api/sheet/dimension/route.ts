import { NextResponse, type NextRequest } from "next/server";
import { getSheetsClient, toFriendlyError } from "@/lib/sheets";

export const dynamic = "force-dynamic";
export const runtime = "nodejs";

/** 행/열 삭제 (deleteDimension). 인덱스는 0-based. */
export async function POST(req: NextRequest) {
  const client = await getSheetsClient();
  if (client.error) return client.error;

  let body: {
    spreadsheetId?: string;
    sheetId?: number;
    dimension?: "ROWS" | "COLUMNS";
    indices?: number[];
  };
  try {
    body = await req.json();
  } catch {
    return NextResponse.json({ error: "잘못된 요청 형식입니다." }, { status: 400 });
  }

  const { spreadsheetId, sheetId, dimension, indices } = body;
  if (!spreadsheetId) {
    return NextResponse.json({ error: "spreadsheetId가 필요합니다." }, { status: 400 });
  }
  if (sheetId == null) {
    return NextResponse.json({ error: "sheetId가 필요합니다." }, { status: 400 });
  }
  if (dimension !== "ROWS" && dimension !== "COLUMNS") {
    return NextResponse.json({ error: "dimension이 올바르지 않습니다." }, { status: 400 });
  }
  if (!Array.isArray(indices) || indices.length === 0) {
    return NextResponse.json({ error: "삭제할 대상이 없습니다." }, { status: 400 });
  }

  // 큰 인덱스부터 삭제해야 남은 인덱스가 밀리지 않는다.
  const sorted = [...new Set(indices)].sort((a, b) => b - a);
  const requests = sorted.map((i) => ({
    deleteDimension: {
      range: { sheetId, dimension, startIndex: i, endIndex: i + 1 },
    },
  }));

  try {
    await client.sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests },
    });
    return NextResponse.json({ ok: true, deleted: sorted.length });
  } catch (err) {
    const { message, status } = toFriendlyError(err);
    return NextResponse.json({ error: message }, { status });
  }
}
