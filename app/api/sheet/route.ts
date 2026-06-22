import { NextResponse, type NextRequest } from "next/server";
import { getServerSession } from "next-auth/next";
import { google } from "googleapis";
import { authOptions } from "@/lib/auth";

// 세션(쿠키)에 의존하므로 항상 동적으로 처리하고 Node 런타임을 사용한다.
export const dynamic = "force-dynamic";
export const runtime = "nodejs";

export async function GET(req: NextRequest) {
  const session = await getServerSession(authOptions);

  if (!session?.accessToken) {
    return NextResponse.json({ error: "로그인이 필요합니다." }, { status: 401 });
  }
  if (session.error === "RefreshAccessTokenError") {
    return NextResponse.json(
      { error: "세션이 만료되었습니다. 다시 로그인해 주세요." },
      { status: 401 },
    );
  }

  const { searchParams } = new URL(req.url);
  const spreadsheetId = searchParams.get("spreadsheetId");
  const range = searchParams.get("range") || "A1:Z1000";

  if (!spreadsheetId) {
    return NextResponse.json(
      { error: "spreadsheetId가 필요합니다." },
      { status: 400 },
    );
  }

  try {
    const auth = new google.auth.OAuth2();
    auth.setCredentials({ access_token: session.accessToken });
    const sheets = google.sheets({ version: "v4", auth });

    const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });

    return NextResponse.json({
      values: res.data.values ?? [],
      range: res.data.range,
    });
  } catch (err: unknown) {
    // googleapis 에러를 사용자 친화적 메시지로 변환
    const e = err as { code?: number; message?: string; errors?: { message?: string }[] };
    const code = e.code ?? 500;
    const message =
      e.errors?.[0]?.message ||
      e.message ||
      "시트를 읽지 못했습니다.";

    const friendly =
      code === 403
        ? "이 시트에 접근 권한이 없습니다. 본인 계정의 시트인지 확인해 주세요."
        : code === 404
          ? "시트를 찾을 수 없습니다. 링크나 ID를 확인해 주세요."
          : message;

    return NextResponse.json({ error: friendly }, { status: code === 403 || code === 404 ? code : 500 });
  }
}
