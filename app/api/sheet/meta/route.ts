import { NextResponse, type NextRequest } from "next/server";
import { getSheetsClient, toFriendlyError } from "@/lib/sheets";

export const dynamic = "force-dynamic";
export const runtime = "nodejs";

/** 스프레드시트의 탭(시트) 목록과 제목을 반환 */
export async function GET(req: NextRequest) {
  const client = await getSheetsClient();
  if (client.error) return client.error;

  const { searchParams } = new URL(req.url);
  const spreadsheetId = searchParams.get("spreadsheetId");
  if (!spreadsheetId) {
    return NextResponse.json({ error: "spreadsheetId가 필요합니다." }, { status: 400 });
  }

  try {
    const res = await client.sheets.spreadsheets.get({
      spreadsheetId,
      fields: "properties.title,sheets(properties(sheetId,title,index,hidden))",
    });

    const sheets = (res.data.sheets ?? []).map((s) => ({
      sheetId: s.properties?.sheetId ?? 0,
      title: s.properties?.title ?? "",
      index: s.properties?.index ?? 0,
      hidden: s.properties?.hidden ?? false,
    }));

    return NextResponse.json({ title: res.data.properties?.title ?? "", sheets });
  } catch (err) {
    const { message, status } = toFriendlyError(err);
    return NextResponse.json({ error: message }, { status });
  }
}
