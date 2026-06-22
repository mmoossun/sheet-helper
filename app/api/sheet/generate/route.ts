import { NextResponse, type NextRequest } from "next/server";
import Anthropic from "@anthropic-ai/sdk";
import { getSheetsClient, toFriendlyError } from "@/lib/sheets";

export const dynamic = "force-dynamic";
export const runtime = "nodejs";

// Claude가 돌려줄 시트 시스템 설계의 JSON 스키마 (구조화 출력으로 강제)
const SHEET_SPEC_SCHEMA = {
  type: "object",
  additionalProperties: false,
  properties: {
    sheets: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        properties: {
          title: { type: "string", description: "탭(시트) 이름" },
          columns: {
            type: "array",
            items: { type: "string" },
            description: "첫 행에 들어갈 열 머리글",
          },
          rows: {
            type: "array",
            items: { type: "array", items: { type: "string" } },
            description:
              "샘플 데이터 행들. 계산이 필요한 칸은 구글 시트 수식(=SUM(...) 등)을 그대로 문자열로 넣는다.",
          },
        },
        required: ["title", "columns", "rows"],
      },
    },
  },
  required: ["sheets"],
} as const;

const SYSTEM_PROMPT = `당신은 구글 시트로 실무 업무 시스템을 설계하는 전문가입니다.
사용자의 요청을 받아, 바로 쓸 수 있는 한국어 스프레드시트 시스템을 설계하세요.

원칙:
- 여러 개의 탭으로 구성하세요 (예: 데이터 탭들 + 요약/대시보드 탭).
- 각 탭은 명확한 한국어 열 머리글로 시작하고, 현실적인 샘플 데이터 5~12행을 채우세요.
- 계산이 필요한 칸에는 구글 시트 수식을 문자열로 그대로 넣으세요 (예: "=C2*D2", "=SUM(매출!E2:E100)", "=TODAY()").
- 대시보드/요약 탭은 다른 탭을 참조하는 수식으로 핵심 지표를 보여주세요.
- 수식의 시트 참조는 당신이 정한 탭 이름과 정확히 일치해야 합니다.
- 탭 개수는 2~6개로 실용적으로 유지하세요.`;

type SheetSpec = { title: string; columns: string[]; rows: string[][] };

/** 시트 제목이 기존/생성목록과 겹치지 않게 유니크하게 만든다. */
function uniqueTitle(base: string, used: Set<string>): string {
  let title = base.trim() || "시트";
  let n = 2;
  while (used.has(title)) title = `${base} (${n++})`;
  used.add(title);
  return title;
}

function quoteTitle(title: string): string {
  return `'${title.replace(/'/g, "''")}'`;
}

export async function POST(req: NextRequest) {
  if (!process.env.ANTHROPIC_API_KEY) {
    return NextResponse.json(
      { error: "AI 기능을 쓰려면 서버에 ANTHROPIC_API_KEY를 설정해야 합니다." },
      { status: 400 },
    );
  }

  const client = await getSheetsClient();
  if (client.error) return client.error;

  let body: { spreadsheetId?: string; prompt?: string };
  try {
    body = await req.json();
  } catch {
    return NextResponse.json({ error: "잘못된 요청 형식입니다." }, { status: 400 });
  }

  const { spreadsheetId, prompt } = body;
  if (!spreadsheetId) {
    return NextResponse.json({ error: "spreadsheetId가 필요합니다." }, { status: 400 });
  }
  if (!prompt || !prompt.trim()) {
    return NextResponse.json({ error: "무엇을 만들지 입력해 주세요." }, { status: 400 });
  }

  // 1) Claude로 시트 시스템 설계(JSON) 생성
  let spec: { sheets: SheetSpec[] };
  try {
    const anthropic = new Anthropic();
    const stream = anthropic.messages.stream({
      model: "claude-opus-4-8",
      max_tokens: 16000,
      thinking: { type: "adaptive" },
      output_config: {
        effort: "medium",
        format: { type: "json_schema", schema: SHEET_SPEC_SCHEMA },
      },
      system: SYSTEM_PROMPT,
      messages: [{ role: "user", content: prompt }],
    });
    const message = await stream.finalMessage();

    if (message.stop_reason === "refusal") {
      return NextResponse.json(
        { error: "이 요청은 처리할 수 없어요. 다른 내용으로 시도해 주세요." },
        { status: 400 },
      );
    }

    const textBlock = message.content.find((b) => b.type === "text");
    if (!textBlock || textBlock.type !== "text") {
      throw new Error("AI 응답을 해석하지 못했습니다.");
    }
    spec = JSON.parse(textBlock.text);
  } catch (err) {
    const msg = err instanceof Error ? err.message : "AI 생성에 실패했습니다.";
    return NextResponse.json({ error: msg }, { status: 500 });
  }

  if (!spec?.sheets?.length) {
    return NextResponse.json({ error: "AI가 만들 시트를 찾지 못했습니다." }, { status: 500 });
  }

  // 2) 설계를 구글 시트에 실제로 반영
  try {
    const meta = await client.sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets(properties(title))",
    });
    const used = new Set<string>(
      (meta.data.sheets ?? []).map((s) => s.properties?.title ?? ""),
    );

    // 탭 생성 (유니크 제목)
    const createdTitles: string[] = [];
    const addRequests = spec.sheets.map((s) => {
      const title = uniqueTitle(s.title, used);
      createdTitles.push(title);
      return { addSheet: { properties: { title } } };
    });
    await client.sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: addRequests },
    });

    // 각 탭에 머리글 + 데이터 채우기 (USER_ENTERED → 수식/숫자 해석)
    const data = spec.sheets.map((s, i) => ({
      range: `${quoteTitle(createdTitles[i])}!A1`,
      values: [s.columns, ...s.rows],
    }));
    await client.sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: "USER_ENTERED", data },
    });

    return NextResponse.json({ ok: true, createdTabs: createdTitles });
  } catch (err) {
    const { message, status } = toFriendlyError(err);
    return NextResponse.json({ error: message }, { status });
  }
}
