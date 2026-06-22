import { NextResponse, type NextRequest } from "next/server";
import { getSheetsClient, toFriendlyError } from "@/lib/sheets";

export const dynamic = "force-dynamic";
export const runtime = "nodejs";

// 열(columnIndex, 0-based)에 데이터 검증을 적용/해제한다. 2행~1000행 대상(1행은 머리글).
export async function POST(req: NextRequest) {
  const client = await getSheetsClient();
  if (client.error) return client.error;

  let body: {
    spreadsheetId?: string;
    sheetId?: number;
    columnIndex?: number;
    type?: "list" | "checkbox" | "clear";
    values?: string[];
  };
  try {
    body = await req.json();
  } catch {
    return NextResponse.json({ error: "잘못된 요청 형식입니다." }, { status: 400 });
  }

  const { spreadsheetId, sheetId, columnIndex, type, values } = body;
  if (!spreadsheetId) {
    return NextResponse.json({ error: "spreadsheetId가 필요합니다." }, { status: 400 });
  }
  if (sheetId == null || columnIndex == null) {
    return NextResponse.json({ error: "대상 열 정보가 필요합니다." }, { status: 400 });
  }
  if (type !== "list" && type !== "checkbox" && type !== "clear") {
    return NextResponse.json({ error: "규칙 종류가 올바르지 않습니다." }, { status: 400 });
  }
  if (type === "list" && (!Array.isArray(values) || values.length === 0)) {
    return NextResponse.json({ error: "드롭다운 항목을 입력해 주세요." }, { status: 400 });
  }

  const range = {
    sheetId,
    startRowIndex: 1, // 2행부터 (1행 머리글 제외)
    endRowIndex: 1000,
    startColumnIndex: columnIndex,
    endColumnIndex: columnIndex + 1,
  };

  let request;
  if (type === "clear") {
    request = { setDataValidation: { range } }; // rule 생략 → 규칙 해제
  } else if (type === "checkbox") {
    request = {
      setDataValidation: { range, rule: { condition: { type: "BOOLEAN" }, showCustomUi: true } },
    };
  } else {
    request = {
      setDataValidation: {
        range,
        rule: {
          condition: {
            type: "ONE_OF_LIST",
            values: values!.map((v) => ({ userEnteredValue: v })),
          },
          showCustomUi: true,
          strict: false, // 목록 외 값도 입력 가능(경고만)
        },
      },
    };
  }

  try {
    await client.sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [request] },
    });
    return NextResponse.json({ ok: true });
  } catch (err) {
    const { message, status } = toFriendlyError(err);
    return NextResponse.json({ error: message }, { status });
  }
}
