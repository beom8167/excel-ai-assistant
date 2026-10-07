exports.handler = async function(event) {
  if (event.httpMethod !== "POST") {
    return json(405, { error: "POST 요청만 지원합니다." });
  }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    return json(500, {
      error: "서버에 OPENAI_API_KEY가 설정되지 않았습니다. Netlify 환경변수에 API 키를 등록해주세요."
    });
  }

  let payload;
  try {
    payload = JSON.parse(event.body || "{}");
  } catch {
    return json(400, { error: "요청 데이터 형식이 올바르지 않습니다." });
  }

  const userRequest = String(payload.request || "").trim();
  const workbook = payload.workbook;

  if (!userRequest || !workbook || !Array.isArray(workbook.sheets)) {
    return json(400, { error: "request와 workbook 정보가 필요합니다." });
  }

  // Keep the request deliberately small for an MVP.
  const compactWorkbook = {
    active_sheet: workbook.active_sheet,
    sheets: workbook.sheets.slice(0, 8).map(s => ({
      name: String(s.name || "").slice(0, 100),
      used_range: s.used_range,
      row_count: s.row_count,
      column_count: s.column_count,
      sample_rows: Array.isArray(s.sample_rows)
        ? s.sample_rows.slice(0, 12).map(r => Array.isArray(r) ? r.slice(0, 18) : [])
        : []
    }))
  };

  const instructions = `
You are an Excel editing planner for a Korean-language web application.
Interpret the user's request using the workbook structure and sample rows.
Return ONLY one valid JSON object. No markdown, no code fences, no commentary outside JSON.

The app supports ONLY these action types:
1) fill_formula
{
  "type":"fill_formula",
  "sheet":"exact sheet name",
  "target_column":"K",
  "start_row":2,
  "end_row":"last",
  "formula_template":"=IFERROR(H{row}/J{row},0)",
  "header":"달성률",
  "header_row":1
}
Use {row} wherever each data row number must be substituted.

2) summary_formula
{
  "type":"summary_formula",
  "sheet":"exact sheet name",
  "function_name":"SUM",
  "range":"J2:J10",
  "target_cell":"J11"
}
Allowed function_name: SUM, AVERAGE, COUNT, COUNTA, MAX, MIN.

3) set_formula_cell
{
  "type":"set_formula_cell",
  "sheet":"exact sheet name",
  "cell":"K2",
  "formula":"=H2/J2"
}

4) set_value_cell
{
  "type":"set_value_cell",
  "sheet":"exact sheet name",
  "cell":"K1",
  "value":"달성률"
}

5) insert_blank_rows_between
{
  "type":"insert_blank_rows_between",
  "sheet":"exact sheet name",
  "start_row":1,
  "end_row":10
}

Output object:
{
  "summary":"short Korean explanation of what will be changed",
  "note":"optional short caution or assumption in Korean",
  "actions":[ ... ]
}

Rules:
- Never invent a sheet name. Use exact workbook sheet names.
- Prefer header names and sample data to infer source columns.
- For a new calculated column, normally choose the first reasonable empty column to the right of existing data, add a clear Korean header, and fill from the first data row through "last".
- If row 1 looks like headers, use start_row 2. If not, infer cautiously.
- Use IFERROR when division may divide by zero.
- Do not delete data.
- Do not create macros, VBA, external links, or unsupported formatting.
- If the request cannot be safely represented with supported actions, return:
  {"summary":"현재 MVP에서 지원하지 않는 요청입니다.","note":"지원 가능한 범위로 요청을 바꿔주세요.","actions":[]}
- Keep the plan minimal. Usually 1 to 3 actions.
`;

  const inputObject = {
    user_request: userRequest,
    workbook: compactWorkbook
  };

  try {
    const response = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        model: process.env.OPENAI_MODEL || "gpt-6-luna",
        instructions,
        input: JSON.stringify(inputObject),
        max_output_tokens: 1400,
        store: false
      })
    });

    const data = await response.json();

    if (!response.ok) {
      const message = data?.error?.message || "OpenAI API 요청에 실패했습니다.";
      return json(response.status, { error: message });
    }

    const outputText = extractOutputText(data);
    if (!outputText) {
      return json(502, { error: "AI 응답에서 작업 계획을 읽지 못했습니다." });
    }

    let plan;
    try {
      plan = JSON.parse(stripCodeFence(outputText));
    } catch {
      return json(502, {
        error: "AI가 올바른 JSON 작업 계획을 반환하지 않았습니다.",
        debug: outputText.slice(0, 500)
      });
    }

    const validationError = validatePlan(plan, compactWorkbook.sheets.map(s => s.name));
    if (validationError) {
      return json(422, { error: validationError });
    }

    return json(200, { plan });
  } catch (error) {
    return json(500, { error: error?.message || "AI 서버 처리 중 오류가 발생했습니다." });
  }
};

function extractOutputText(data) {
  const parts = [];
  for (const item of data?.output || []) {
    for (const content of item?.content || []) {
      if (content?.type === "output_text" && typeof content.text === "string") {
        parts.push(content.text);
      }
    }
  }
  return parts.join("\n").trim();
}

function stripCodeFence(text) {
  return String(text)
    .trim()
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/i, "")
    .trim();
}

function validatePlan(plan, sheetNames) {
  if (!plan || typeof plan !== "object") return "작업 계획 형식이 올바르지 않습니다.";
  if (!Array.isArray(plan.actions)) return "actions 배열이 없습니다.";
  if (plan.actions.length > 5) return "한 번에 최대 5개의 작업만 허용합니다.";

  const allowed = new Set([
    "fill_formula",
    "summary_formula",
    "set_formula_cell",
    "set_value_cell",
    "insert_blank_rows_between"
  ]);

  for (const action of plan.actions) {
    if (!action || !allowed.has(action.type)) return `지원하지 않는 작업 유형입니다: ${action?.type}`;
    if (!sheetNames.includes(action.sheet)) return `존재하지 않는 시트를 요청했습니다: ${action.sheet}`;

    if (action.type === "fill_formula") {
      if (!/^[A-Z]+$/i.test(String(action.target_column || ""))) return "잘못된 대상 열입니다.";
      if (typeof action.formula_template !== "string" || !action.formula_template.includes("{row}")) {
        return "fill_formula의 formula_template에는 {row}가 필요합니다.";
      }
    }
    if (action.type === "summary_formula") {
      if (!["SUM","AVERAGE","COUNT","COUNTA","MAX","MIN"].includes(String(action.function_name || "").toUpperCase())) {
        return "지원하지 않는 요약 함수입니다.";
      }
      if (!/^[A-Z]+\d+:[A-Z]+\d+$/i.test(String(action.range || ""))) return "잘못된 함수 범위입니다.";
      if (!/^[A-Z]+\d+$/i.test(String(action.target_cell || ""))) return "잘못된 결과 셀입니다.";
    }
    if (action.type === "set_formula_cell" || action.type === "set_value_cell") {
      if (!/^[A-Z]+\d+$/i.test(String(action.cell || ""))) return "잘못된 셀 주소입니다.";
    }
    if (action.type === "insert_blank_rows_between") {
      const a = Number(action.start_row), b = Number(action.end_row);
      if (!Number.isInteger(a) || !Number.isInteger(b) || a < 1 || b <= a) return "빈 행 삽입 범위가 올바르지 않습니다.";
    }
  }
  return null;
}

function json(statusCode, body) {
  return {
    statusCode,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store"
    },
    body: JSON.stringify(body)
  };
}
