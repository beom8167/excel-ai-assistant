exports.handler = async function (event) {
  function respond(statusCode, body) {
    return {
      statusCode,
      headers: {
        "Content-Type": "application/json; charset=utf-8",
        "Cache-Control": "no-store"
      },
      body: JSON.stringify(body)
    };
  }

  // Test whether the Netlify function itself is working.
  if (event.httpMethod === "GET") {
    return respond(200, {
      ok: true,
      message: "AI function is deployed and handler is working."
    });
  }

  if (event.httpMethod !== "POST") {
    return respond(405, {
      error: "POST only"
    });
  }

  const apiKey = process.env.OPENAI_API_KEY;

  if (!apiKey) {
    return respond(500, {
      error: "OPENAI_API_KEY is missing in Netlify environment variables."
    });
  }

  // Use OPENAI_MODEL from Netlify if set.
  // Otherwise use gpt-6-astra.
  const model = process.env.OPENAI_MODEL || "gpt-6-astra";

  let payload;

  try {
    payload = JSON.parse(event.body || "{}");
  } catch (err) {
    return respond(400, {
      error: "Invalid JSON request body.",
      detail: String(err && err.message ? err.message : err)
    });
  }

  const userRequest = String(payload.request || "").trim();
  const workbook = payload.workbook;

  if (!userRequest) {
    return respond(400, {
      error: "User request is empty."
    });
  }

  if (!workbook || !Array.isArray(workbook.sheets)) {
    return respond(400, {
      error: "Workbook structure is missing."
    });
  }

  const compactWorkbook = {
    active_sheet: workbook.active_sheet,
    sheets: workbook.sheets.slice(0, 8).map(function (sheet) {
      return {
        name: String(sheet.name || "").slice(0, 100),
        used_range: sheet.used_range || null,
        row_count: Number(sheet.row_count || 0),
        column_count: Number(sheet.column_count || 0),
        sample_rows: Array.isArray(sheet.sample_rows)
          ? sheet.sample_rows.slice(0, 12).map(function (row) {
              return Array.isArray(row)
                ? row.slice(0, 18)
                : [];
            })
          : []
      };
    })
  };

  const instructions = [
    "You are an Excel editing planner for a Korean-language web application.",
    "",
    "The user will describe an Excel editing task in Korean.",
    "Use the workbook sheet names, headers, ranges, and sample data to understand the request.",
    "",
    "Return exactly one valid JSON object.",
    "Do not use markdown.",
    "Do not use code fences.",
    "Do not include commentary outside the JSON object.",
    "",
    "SUPPORTED ACTIONS",
    "",
    "1. fill_formula",
    JSON.stringify({
      type: "fill_formula",
      sheet: "Sheet1",
      target_column: "D",
      start_row: 2,
      end_row: "last",
      formula_template: "=IFERROR(C{row}/B{row},0)",
      header: "달성률",
      header_row: 1
    }),
    "",
    "Use {row} for the row number.",
    "",
    "2. summary_formula",
    JSON.stringify({
      type: "summary_formula",
      sheet: "Sheet1",
      function_name: "SUM",
      range: "C2:C10",
      target_cell: "C11"
    }),
    "",
    "Allowed summary functions:",
    "SUM, AVERAGE, COUNT, COUNTA, MAX, MIN",
    "",
    "3. set_formula_cell",
    JSON.stringify({
      type: "set_formula_cell",
      sheet: "Sheet1",
      cell: "D2",
      formula: "=C2/B2"
    }),
    "",
    "4. set_value_cell",
    JSON.stringify({
      type: "set_value_cell",
      sheet: "Sheet1",
      cell: "D1",
      value: "달성률"
    }),
    "",
    "5. insert_blank_rows_between",
    JSON.stringify({
      type: "insert_blank_rows_between",
      sheet: "Sheet1",
      start_row: 1,
      end_row: 10
    }),
    "",
    "OUTPUT FORMAT",
    JSON.stringify({
      summary: "작업 내용을 한국어로 짧게 설명",
      note: "필요한 경우 주의사항이나 가정",
      actions: []
    }),
    "",
    "RULES",
    "- Always use an exact sheet name that exists in the workbook.",
    "- Infer columns from Korean headers when possible.",
    "- If row 1 contains headers, normally start formulas from row 2.",
    "- When creating a new calculated column, normally use the first empty column to the right of the existing data.",
    "- Give a clear Korean header to a newly created column.",
    "- Use {row} when the same formula needs to be applied to multiple rows.",
    "- For division, use IFERROR when division by zero could occur.",
    "- Do not delete existing data.",
    "- Do not create VBA or macros.",
    "- Do not create external links.",
    "- Do not create charts or pivots.",
    "- Keep the plan minimal, normally 1 to 3 actions.",
    "- If the request cannot be performed using the supported actions, return an empty actions array.",
    "",
    "Example:",
    "If the workbook has:",
    "A = 부서",
    "B = 목표금액",
    "C = 실적금액",
    "",
    "and the user asks:",
    "목표금액과 실적금액을 비교해서 오른쪽에 달성률 열을 만들어줘",
    "",
    "a suitable result is:",
    JSON.stringify({
      summary: "목표금액과 실적금액을 비교해 D열에 달성률을 추가합니다.",
      note: "0으로 나누는 오류는 IFERROR로 처리합니다.",
      actions: [
        {
          type: "fill_formula",
          sheet: "Sheet1",
          target_column: "D",
          start_row: 2,
          end_row: "last",
          formula_template: "=IFERROR(C{row}/B{row},0)",
          header: "달성률",
          header_row: 1
        }
      ]
    })
  ].join("\n");

  const input = JSON.stringify({
    user_request: userRequest,
    workbook: compactWorkbook
  });

  let response;
  let data;

  try {
    response = await fetch(
      "https://api.openai.com/v1/responses",
      {
        method: "POST",
        headers: {
          "Authorization": "Bearer " + apiKey,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          model: model,
          instructions: instructions,
          input: input,
          max_output_tokens: 3000,
          store: false
        })
      }
    );

    data = await response.json();
  } catch (err) {
    console.error("OPENAI_NETWORK_ERROR", err);

    return respond(500, {
      error: "OpenAI network error.",
      detail: String(
        err && err.message
          ? err.message
          : err
      )
    });
  }

  if (!response.ok) {
    let message = "OpenAI HTTP " + response.status;

    if (
      data &&
      data.error &&
      data.error.message
    ) {
      message = data.error.message;
    }

    console.error(
      "OPENAI_API_ERROR",
      response.status,
      message
    );

    return respond(response.status, {
      error: "OpenAI API error: " + message,
      openai_status: response.status,
      model: model
    });
  }

  const outputText = extractOutputText(data);

  if (!outputText) {
    return respond(500, {
      error: "OpenAI response did not contain output text.",
      response_status:
        data && data.status
          ? data.status
          : null
    });
  }

  let plan;

  try {
    plan = parseJsonObject(outputText);
  } catch (err) {
    console.error(
      "PLAN_JSON_PARSE_ERROR",
      outputText
    );

    return respond(422, {
      error: "Could not parse AI output as JSON.",
      preview: outputText.slice(0, 500)
    });
  }

  const validationError = validatePlan(
    plan,
    compactWorkbook.sheets.map(
      function (sheet) {
        return sheet.name;
      }
    )
  );

  if (validationError) {
    return respond(422, {
      error: validationError,
      plan: plan
    });
  }

  return respond(200, {
    plan: plan,
    model: model
  });
};


function extractOutputText(data) {
  const chunks = [];

  const output =
    data && Array.isArray(data.output)
      ? data.output
      : [];

  output.forEach(function (item) {
    const content =
      item && Array.isArray(item.content)
        ? item.content
        : [];

    content.forEach(function (part) {
      if (
        part &&
        part.type === "output_text" &&
        typeof part.text === "string"
      ) {
        chunks.push(part.text);
      }
    });
  });

  return chunks.join("\n").trim();
}


function parseJsonObject(text) {
  let cleaned =
    String(text || "").trim();

  cleaned = cleaned
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/i, "")
    .trim();

  try {
    return JSON.parse(cleaned);
  } catch (err) {
    // Continue and try to find a JSON object.
  }

  const first =
    cleaned.indexOf("{");

  const last =
    cleaned.lastIndexOf("}");

  if (
    first !== -1 &&
    last > first
  ) {
    return JSON.parse(
      cleaned.slice(
        first,
        last + 1
      )
    );
  }

  throw new Error(
    "JSON object not found"
  );
}


function validatePlan(
  plan,
  sheetNames
) {
  if (
    !plan ||
    typeof plan !== "object"
  ) {
    return "Invalid plan object.";
  }

  if (
    !Array.isArray(plan.actions)
  ) {
    return "Plan actions array is missing.";
  }

  if (
    plan.actions.length > 5
  ) {
    return "Maximum 5 actions per request.";
  }

  const allowed = [
    "fill_formula",
    "summary_formula",
    "set_formula_cell",
    "set_value_cell",
    "insert_blank_rows_between"
  ];

  for (
    let i = 0;
    i < plan.actions.length;
    i++
  ) {
    const action =
      plan.actions[i];

    if (
      !action ||
      allowed.indexOf(
        action.type
      ) === -1
    ) {
      return (
        "Unsupported action type: " +
        (
          action
            ? action.type
            : "unknown"
        )
      );
    }

    if (
      sheetNames.indexOf(
        action.sheet
      ) === -1
    ) {
      return (
        "Unknown sheet: " +
        action.sheet
      );
    }

    if (
      action.type ===
      "fill_formula"
    ) {
      if (
        !/^[A-Z]+$/i.test(
          String(
            action.target_column ||
            ""
          )
        )
      ) {
        return "Invalid target column.";
      }

      if (
        typeof
          action.formula_template !==
          "string" ||
        action.formula_template.indexOf(
          "{row}"
        ) === -1
      ) {
        return (
          "fill_formula requires {row}."
        );
      }
    }

    if (
      action.type ===
      "summary_formula"
    ) {
      const functionName =
        String(
          action.function_name ||
          ""
        ).toUpperCase();

      const functions = [
        "SUM",
        "AVERAGE",
        "COUNT",
        "COUNTA",
        "MAX",
        "MIN"
      ];

      if (
        functions.indexOf(
          functionName
        ) === -1
      ) {
        return (
          "Unsupported summary function."
        );
      }

      if (
        !/^[A-Z]+\d+:[A-Z]+\d+$/i.test(
          String(
            action.range ||
            ""
          )
        )
      ) {
        return "Invalid summary range.";
      }

      if (
        !/^[A-Z]+\d+$/i.test(
          String(
            action.target_cell ||
            ""
          )
        )
      ) {
        return (
          "Invalid summary target cell."
        );
      }
    }

    if (
      action.type ===
        "set_formula_cell" ||
      action.type ===
        "set_value_cell"
    ) {
      if (
        !/^[A-Z]+\d+$/i.test(
          String(
            action.cell ||
            ""
          )
        )
      ) {
        return "Invalid cell address.";
      }
    }

    if (
      action.type ===
      "insert_blank_rows_between"
    ) {
      const start =
        Number(
          action.start_row
        );

      const end =
        Number(
          action.end_row
        );

      if (
        !Number.isInteger(start) ||
        !Number.isInteger(end) ||
        start < 1 ||
        end <= start
      ) {
        return (
          "Invalid row insertion range."
        );
      }
    }
  }

  return null;
}
