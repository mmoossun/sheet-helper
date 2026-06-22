import { NextResponse } from "next/server";
import { getServerSession } from "next-auth/next";
import { google, type sheets_v4 } from "googleapis";
import { authOptions } from "@/lib/auth";

/** 세션을 확인하고 인증된 Sheets 클라이언트를 돌려준다. 실패 시 에러 응답을 반환. */
export async function getSheetsClient(): Promise<
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
export function toFriendlyError(err: unknown): { message: string; status: number } {
  const e = err as { code?: number; message?: string; errors?: { message?: string }[] };
  const code = e.code ?? 500;
  const raw = e.errors?.[0]?.message || e.message || "요청을 처리하지 못했습니다.";

  if (code === 403)
    return { message: "이 시트에 접근 권한이 없습니다. 본인 계정의 시트인지 확인해 주세요.", status: 403 };
  if (code === 404)
    return { message: "시트를 찾을 수 없습니다. 링크나 ID를 확인해 주세요.", status: 404 };
  return { message: raw, status: 500 };
}

/** 시트 탭 제목을 A1 표기에 안전하게 넣기 위해 따옴표 처리한다. (제목 내 ' 는 '' 로 이스케이프) */
function quoteSheetTitle(title: string): string {
  return `'${title.replace(/'/g, "''")}'`;
}

/** 탭 제목이 있으면 `'제목'!A1` 형태로, 없으면 그대로 A1 범위를 만든다. */
export function buildRange(sheetTitle: string | null | undefined, a1: string): string {
  if (!sheetTitle) return a1;
  return `${quoteSheetTitle(sheetTitle)}!${a1}`;
}
