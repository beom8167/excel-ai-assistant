exports.handler = async function(event) {
  이벤트의 httpMethod가 "POST"가 아닌 경우 {
    return json(405, { error: "POST 요청만 지원합니다." });
  }

  const apiKey = process.env.OPENAI_API_KEY;
  만약 (!apiKey)라면 {
    json(500, {를 반환합니다.
      오류: "서버에 OPENAI_API_KEY가 설정되어 있지 않습니다. Netlify 환경 변수에 API 키를 등록해 주세요."
    });
  }

  페이로드를 저장하세요;
  노력하다 {
    payload = JSON.parse(event.body || "{}");
  } 잡다 {
    return json(400, { error: "요청 데이터 형식이 올바르지 않습니다." });
  }

  const userRequest = String(payload.request || "").trim();
  const workbook = payload.workbook;

  만약 userRequest가 아니고 workbook이 아니고 workbook.sheets가 Array인 경우라면
    return json(400, { error: "요청과 통합 문서 정보가 필요합니다." });
  }

  // MVP를 위해 요청 규모를 의도적으로 작게 유지합니다.
  const compactWorkbook = {
    활성 시트: 워크북.활성 시트,
    시트: workbook.sheets.slice(0, 8).map(s => ({
      이름: 문자열(s.name || "").slice(0, 100),
      used_range: s.used_range,
      행 수: s.row_count,
      column_count: s.column_count,
      sample_rows: Array.isArray(s.sample_rows)
        ? s.sample_rows.slice(0, 12).map(r => Array.isArray(r) ? r.slice(0, 18) : [])
        : []
    }))
  };

  const instructions = `
당신은 한국어 웹 애플리케이션의 엑셀 편집 기획 담당자입니다.
통합 문서 구조와 샘플 행을 사용하여 사용자의 요청을 해석합니다.
유효한 JSON 객체 하나만 반환하세요. 마크다운, 코드 펜스, JSON 외부의 주석은 허용되지 않습니다.

이 앱은 다음 유형의 액션만 지원합니다.
1) fill_formula
{
  "유형":"수식 채우기",
  "시트":"정확한 시트 이름",
  "target_column":"K",
  "start_row":2,
  "end_row":"마지막",
  "formula_template":"=IFERROR(H{row}/J{row},0)",
  "header": "달성률",
  "헤더_행":1
}
각 데이터 행 번호를 입력해야 하는 곳에는 {row}를 사용하십시오.

2) 요약 공식
{
  "유형":"요약_수식",
  "시트":"정확한 시트 이름",
  "함수 이름":"합계",
  "범위":"J2:J10",
  "target_cell":"J11"
}
허용되는 함수 이름: SUM, AVERAGE, COUNT, COUNTA, MAX, MIN.

3) set_formula_cell
{
  "유형":"set_formula_cell",
  "시트":"정확한 시트 이름",
  "세포":"K2",
  "공식":"=H2/J2"
}

4) set_value_cell
{
  "유형":"set_value_cell",
  "시트":"정확한 시트 이름",
  "세포":"K1",
  "value":"달성률"
}

5) insert_blank_rows_between
{
  "유형":"빈 행 사이 삽입",
  "시트":"정확한 시트 이름",
  "start_row":1,
  "end_row":10
}

출력 객체:
{
  "요약":"변경될 내용에 대한 간략한 한국어 설명"
  "참고": "선택 사항인 한국어로 된 간략한 주의 사항 또는 가정"
  "액션":[ ... ]
}

규칙:
- 절대로 임의로 시트 이름을 만들지 마세요. 통합 문서의 정확한 시트 이름을 사용하십시오.
- 헤더 이름과 샘플 데이터를 활용하여 소스 열을 추론하는 것이 좋습니다.
- 새 계산 열을 만들려면 일반적으로 기존 데이터의 오른쪽에 있는 적절한 첫 번째 빈 열을 선택하고, 명확한 한국어 헤더를 추가한 다음, 첫 번째 데이터 행부터 마지막 ​​행까지 데이터를 채웁니다.
- 1행이 헤더처럼 보이면 start_row 2를 사용하십시오. 그렇지 않으면 신중하게 추론하십시오.
- 나눗셈 결과가 0이 될 경우 IFERROR를 사용하십시오.
- 데이터를 삭제하지 마세요.
- 매크로, VBA, 외부 링크 또는 지원되지 않는 서식을 만들지 마십시오.
- 요청을 지원되는 작업으로 안전하게 표현할 수 없는 경우 다음을 반환합니다.
  {"summary":"현재 MVP에서 지원하지 않는 요청입니다.","note":"지원 가능한 범위로 요청해 주세요.","actions":[]}
- 계획은 최소한으로 유지하세요. 보통 1~3가지 행동으로 충분합니다.
`;

  const inputObject = {
    사용자 요청: 사용자 요청,
    워크북: 컴팩트워크북
  };

  노력하다 {
    const response = await fetch("https://api.openai.com/v1/responses", {
      방법: "POST",
      헤더: {
        "인증": `Bearer ${apiKey}`,
        "콘텐츠 유형": "application/json"
      },
      본문: JSON.stringify({
        모델: process.env.OPENAI_MODEL || "gpt-6-luna",
        지침,
        입력: JSON.stringify(inputObject),
        최대 출력 토큰 수: 1400,
        저장: false
      })
    });

    const data = await response.json();

    응답이 괜찮지 않으면 {
      const 메시지 = data?.error?.message || "OpenAI API 요청에 실패했습니다.";
      응답의 상태를 나타내는 JSON을 반환합니다. { 오류: 메시지 }
    }

    const outputText = extractOutputText(data);
    만약 출력텍스트가 아니라면 {
      return json(502, { error: "AI 응답에서 작업을 준비했습니다." });
    }

    계획을 세우자;
    노력하다 {
      계획 = JSON.parse(stripCodeFence(outputText));
    } 잡다 {
      json(502, {를 반환합니다.
        오류: "AI가 올바른 JSON 작업을 계획하지 않았습니다.",
        디버그: outputText.slice(0, 500)
      });
    }

    const validationError = validatePlan(plan, compactWorkbook.sheets.map(s => s.name));
    유효성 검사 오류가 발생하면
      json(422, { error: validationError })를 반환합니다.
    }

    json(200, { plan })을 반환합니다.
  } catch (오류) {
    return json(500, { error: error?.message || "AI 서버 처리 중 오류가 발생했습니다." });
  }
};

함수 extractOutputText(data) {
  const parts = [];
  for (const item of data?.output || []) {
    for (const content of item?.content || []) {
      만약 content?.type이 "output_text"이고 content.text의 typeof가 "string"이면,
        parts.push(content.text);
      }
    }
  }
  parts.join("\n").trim();를 반환합니다.
}

함수 stripCodeFence(text) {
  문자열(text)을 반환합니다.
    .손질()
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/i, "")
    .손질();
}

함수 validatePlan(plan, sheetNames) {
  if (!plan || typeof plan !== "object") return "작업 계획 형식이 올바르지 않습니다.";
  if (!Array.isArray(plan.actions)) return "작업 배열이 없습니다.";
  if (plan.actions.length > 5) return "한 번에 최대 5개의 작업만 허용합니다.";

  const allowed = new Set([
    "fill_formula",
    "요약 공식",
    "set_formula_cell",
    "set_value_cell",
    "빈 행을 사이에 삽입"
  ]);

  (plan.actions의 const action에 대해) {
    if (!action || !allowed.has(action.type)) return `지원하지 않는 작업 유형입니다: ${action?.type}`;
    if (!sheetNames.includes(action.sheet)) return `존재 없는 시트를 요청했습니다: ${action.sheet}`;

    만약 action.type이 "fill_formula"와 같다면,
      if (!/^[AZ]+$/i.test(String(action.target_column || ""))) return "잘못된 대상 열입니다.";
      action.formula_template의 typeof가 "string"이 아니거나 action.formula_template에 "{row}"가 포함되어 있지 않으면
        return "fill_formula의 Formula_template에는 {행}이 필요합니다.";
      }
    }
    만약 action.type이 "summary_formula"와 같다면,
      만약 ["SUM","AVERAGE","COUNT","COUNTA","MAX","MIN"]이 문자열(action.function_name || "")을 대문자로 변환한 것이 아니라면
        return "지원하지 않는 요약입니다.";
      }
      if (!/^[AZ]+\d+:[AZ]+\d+$/i.test(String(action.range || ""))) return "잘못된 함수 범위입니다.";
      if (!/^[AZ]+\d+$/i.test(String(action.target_cell || ""))) return "잘못된 결과 셀입니다.";
    }
    만약 action.type이 "set_formula_cell"이거나 action.type이 "set_value_cell"이라면,
      if (!/^[AZ]+\d+$/i.test(String(action.cell || ""))) return "잘못된 셀 주소입니다.";
    }
    만약 action.type이 "insert_blank_rows_between"과 같다면,
      const a = Number(action.start_row), b = Number(action.end_row);
      if (!Number.isInteger(a) || !Number.isInteger(b) || a < 1 || b <= a) return "빈 행 삽입 범위가 올바르지 않습니다.";
    }
  }
  null을 반환합니다.
}

함수 json(statusCode, body) {
  반품 {
    상태 코드,
    헤더: {
      "Content-Type": "application/json; charset=utf-8",
      "캐시 제어": "저장 안 함"
    },
    본문: JSON.stringify(body)
  };
}
