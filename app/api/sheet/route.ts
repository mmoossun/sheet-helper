import { NextResponse, type NextRequest } from "next/server";
import { getServerSession } from "next-auth/next";
import { google, type sheets_v4 } from "googleapis";
import { authOptions } from "@/lib/auth";

// 세션(쿠키)에 의존하므로 항상 동적으로 처리하고 Node 런타임을 사용한다.
export const dynamic = "force-dynamic";
export const runtime = "nodejs";

/** 세션을 확인하고 인증된 Sheets 클라이언트를 돌려준다. 실패 시 에러 응답을 반환. */
async function getSheetsClient(): Promise<
  | { sheets: sheets_v4.Sheets; error?: undefined }
  | { sheets?: undefined; error: NextResponse }
> {
  const session = await getServerSession(authOptions);

  if (!session?.accessToken) {
    return { error: NextResponse.json({ error: "로그인이 필요합니다." }, { status: 401 }) };
  }
  if (session.error === "RefreshAccessTokenError") {
    return {
      error: NextResponse.json(
        { error: "세션이 만료되었습니다. 다시 로그인해 주세요." },
        { status: 401 },
      ),
    };
  }

  const auth = new google.auth.OAuth2();
  auth.setCredentials({ access_token: session.accessToken });
  return { sheets: google.sheets({ version: "v4", auth }) };
}

/** googleapis 에러를 사용자 친화적인 메시지/상태코드로 변환 */
function toFriendlyError(err: unknown): { message: string; status: number } {
  const e = err as { code?: number; message?: string; errors?: { message?: string }[] };
  const code = e.code ?? 500;
  const raw = e.errors?.[0]?.message || e.message || "요청을 처리하지 못했습니다.";

  if (code === 403)
    return { message: "이 시트에 접근 권한이 없습니다. 본인 계정의 시트인지 확인해 주세요.", status: 403 };
  if (code === 404)
    return { message: "시트를 찾을 수 없습니다. 링크나 ID를 확인해 주세요.", status: 404 };
  return { message: raw, status: 500 };
}

/** 시트 값 읽기 */
export async function GET(req: NextRequest) {
  const client = await getSheetsClient();
  if (client.error) return client.error;

  const { searchParams } = new URL(req.url);
  const spreadsheetId = searchParams.get("spreadsheetId");
  const range = searchParams.get("range") || "A1:Z1000";

  if (!spreadsheetId) {
    return NextResponse.json({ error: "spreadsheetId가 필요합니다." }, { status: 400 });
  }

  try {
    const res = await client.sheets.spreadsheets.values.get({ spreadsheetId, range });
    return NextResponse.json({ values: res.data.values ?? [], range: res.data.range });
  } catch (err) {
    const { message, status } = toFriendlyError(err);
    return NextResponse.json({ error: message }, { status });
  }
}

type CellUpdate = { range: string; value: string };

/** 셀 값 저장 (여러 셀을 한 번에 batchUpdate) */
export async function POST(req: NextRequest) {
  const client = await getSheetsClient();
  if (client.error) return client.error;

  let body: { spreadsheetId?: string; updates?: CellUpdate[] };
  try {
    body = await req.json();
  } catch {
    return NextResponse.json({ error: "잘못된 요청 형식입니다." }, { status: 400 });
  }

  const { spreadsheetId, updates } = body;
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
        data: updates.map((u) => ({ range: u.range, values: [[u.value]] })),
      },
    });
    return NextResponse.json({ ok: true, saved: updates.length });
  } catch (err) {
    const { message, status } = toFriendlyError(err);
    return NextResponse.json({ error: message }, { status });
  }
}
