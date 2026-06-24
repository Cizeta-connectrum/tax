/**
 * 確定申告マネージャー & スマートレシート・申告書管理 - バックエンド
 */

const SHEET_NAME = "レシート";

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('確定申告マネージャー')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error("【APIキー未設定エラー】スクリプトプロパティに「GEMINI_API_KEY」を設定してください。");
  }
  return apiKey.replace(/[\s\t\n\r]/g, '');
}

function triggerAuthorizationTest() {
  try {
    UrlFetchApp.fetch("https://generativelanguage.googleapis.com/", {muteHttpExceptions: true});
    Logger.log("通信承認テスト成功。承認はすでに完了しています。");
  } catch(e) {
    Logger.log("通信テスト中にエラーが発生しました。承認されれば解消します: " + e.toString());
  }
}

function initSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["日付", "内容・店名", "金額", "カテゴリ", "区分", "メモ", "登録日時"]);
    sheet.setFrozenRows(1);
  } else {
    const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
    if (headers.length === 6 && headers[5] === "登録日時") {
      sheet.insertColumnBefore(5);
      sheet.getRange(1, 5).setValue("区分");
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const fillArray = Array(lastRow - 1).fill(["経費"]);
        sheet.getRange(2, 5, lastRow - 1, 1).setValues(fillArray);
      }
    }
  }
  return sheet;
}

function getAvailableModels(apiKey) {
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());
    if (data && data.models) {
      return data.models.filter(m => m.supportedGenerationMethods.includes("generateContent"));
    }
    return [];
  } catch (e) {
    return [];
  }
}

function analyzeImage(base64Data, mimeType) {
  const cleanApiKey = getApiKey();
  const actualMimeType = mimeType || "image/png";
  const availableModels = getAvailableModels(cleanApiKey);
  const priorityModels = ["gemini-2.5-flash", "gemini-1.5-flash", "gemini-2.0-flash", "gemini-2.5-pro"];
  const prompt = `Analyze this receipt document and extract items in Japanese JSON format. Return ONLY the JSON array: [{ "date": "YYYY-MM-DD", "amount": number, "shop": "shop name", "category": "消耗品費/旅費交通費/接待交際費/会議費/通信費/租税公課/広告宣伝費/支払手数料/福利厚生費/新聞図書費/地代家賃/減価償却費/雑費", "memo": "desc" }]`;

  return callGemini(cleanApiKey, availableModels, priorityModels, base64Data, prompt, actualMimeType);
}

/**
 * 徹底的に厳格化したPDF読み取り専用プロンプト
 * 数値の読み取り精度を最優先するため、レシート解析よりも高精度な gemini-2.5-pro を優先する
 */
function analyzeTaxDocument(base64Data, mimeType) {
  const cleanApiKey = getApiKey();
  const actualMimeType = mimeType || "application/pdf";
  const availableModels = getAvailableModels(cleanApiKey);
  const priorityModels = ["gemini-2.5-pro", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"];

  const prompt = `あなたはプロのデータ入力オペレーターです。添付された日本の確定申告書類（「確定申告書B（第一表・第二表）」「収支内訳書」のいずれか、または複数）を解析し、記載されている金額を一切の推測や計算を行わずに、そのまま抽出してJSON形式で返してください。
Markdownの装飾は一切不要です。純粋なJSON文字列のみを出力してください。

【出力JSONフォーマット】
{
  "year": 申告年度の西暦（例: 平成24年分なら 2012）,
  "items": [
    {
      "type": "revenue" | "salary" | "expense" | "deduction" | "income_detail" | "tax",
      "category": "指定の勘定科目名",
      "title": "抽出した項目名",
      "amount": 数値金額（カンマを除去した整数）
    }
  ]
}

【抽出対象とカテゴリの絶対ルール（以下のカテゴリ名のみを使用すること）】
■ 確定申告書B（第一表）の「収入金額等」ブロックから：
- 「事業（営業等）ア」の金額 ➔ type: "revenue", category: "売上・事業収入"
- 「給与 カ」の金額 ➔ type: "salary", category: "給与収入"

■ 確定申告書B（第一表）の「所得金額」ブロックから：
- 「事業（営業等）①」の金額 ➔ type: "revenue", category: "事業所得"
- 「給与 ⑥」の金額 ➔ type: "salary", category: "給与所得"

■ 確定申告書B（第一表）の「所得から差し引かれる金額」ブロックから：
- 「社会保険料控除 ⑫」 ➔ type: "deduction", category: "社会保険料控除"
- 「基礎控除 ㉔」 ➔ type: "deduction", category: "基礎控除"
- その他記載のある控除項目

■ 収支内訳書の「経費」ブロックから（必ず「経費の内訳」の表組み部分から1行ずつ抽出すること）：
- 経費の内訳表には通常、以下の順番で項目が印字されています：給料賃金、外注工賃、減価償却費、貸倒金、地代家賃、利子割引料、租税公課、荷造運賃、水道光熱費、旅費交通費、通信費、広告宣伝費、接待交際費、損害保険料、修繕費、消耗品費、福利厚生費（、雑費）。
- これらの項目名の右側（同じ行）に印字されている金額を、その項目名の category として抽出してください。行をずらして金額を読み取らないよう、各項目名とその右側の金額が必ず同じ行にあることを確認してください。
- 雑費の前後に、納税者が手書き・記入した独自の科目名（空欄行に書かれた項目。例：会議費、雑給、書籍費、諸会費 など）がある場合は、その科目名をそのまま category として使用し、絶対に「雑費」へ書き換えないでください（独自科目名と雑費は別の行として扱う）。
- 「経費計」「差引金額」など、内訳の合計値や差引計算結果は expense の項目として抽出しないでください（個別の科目金額のみを抽出し、合計行は除外することで二重計上を防ぐこと）。
- ページ上部や税務署整理欄にある番号付きのボックス（①〜㉑などの集計用グリッド）は、経費の内訳表とは別物です。このボックスの数値を経費の内訳科目として重複抽出しないでください。あくまで「経費の内訳」表組みの行のみを expense として抽出してください。

■ 第一表の数字グリッドだけで「事業（営業等）ア」「①」が0または空白に見える場合、必ず確定申告書B「第二表」の「所得の内訳」または「雑所得（業務）に関する事項」の表も確認してください。
- 「業務」「雑所得」などの行に「収入金額」「必要経費等」「差引金額」が記載されていれば：
  - 収入金額 ➔ type: "revenue", category: "売上・事業収入"
  - 差引金額（所得金額） ➔ type: "revenue", category: "事業所得"
  - 必要経費等 ➔ type: "expense", category: "雑費", title: "必要経費合計（第二表より）"
  この第二表の数値は、第一表の数字グリッドより読み取りやすい場合が多いため、両方を必ず確認し、実際に印字されている数値を採用してください。

■ 確定申告書B（第一表）の「税金の計算」ブロックから（㉖以降）：
- 「課税される所得金額 ㉖」 ➔ type: "tax", category: "課税される所得金額", title: "課税される所得金額"
- 「上の㉖に対する税額 ㉗」 ➔ type: "tax", category: "算出税額", title: "算出税額"
- 「差引所得税額」（控除後） ➔ type: "tax", category: "差引所得税額", title: "差引所得税額"
- 「復興特別所得税額」（第一表に記載がある場合のみ。平成24年分以前の古い様式にはこの項目自体が存在しないため、記載が無ければ抽出しないこと） ➔ type: "tax", category: "復興特別所得税額", title: "復興特別所得税額"
- 「所得税及び復興特別所得税の額」（合計。古い様式では「差引所得税額」と同額の場合がある） ➔ type: "tax", category: "所得税及び復興特別所得税の額", title: "所得税及び復興特別所得税の額"
- 「源泉徴収税額」 ➔ type: "tax", category: "源泉徴収税額", title: "源泉徴収税額"
  - 第一表の「源泉徴収税額」欄が空白・読み取り不能な場合は、必ず確定申告書B「第二表」の「所得の内訳（源泉徴収税額）」の表にある「源泉徴収税額の合計額」を採用してください（各収入源ごとの源泉徴収税額を全て合算した数値が印字されています）。
- 「申告納税額」（納める税金・還付される税金） ➔ type: "tax", category: "申告納税額", title: "申告納税額"
  - この欄の数値に「▲」「△」「－（マイナス）」などの記号が付いている場合は、還付される税金であることを示すため、amount は必ず負の数値（例: -201309）として返してください。記号が無ければ納める税金として正の数値で返してください。
これらの「税金の計算」欄に記載がある項目だけ抽出し、記載のない項目（古い様式に存在しない項目を含む）は除外してください。

■ 確定申告書B（第二表）の「所得の内訳（源泉徴収税額）」テーブルから：
- 表内の各行（給与（会社名）、給与 会社名、業務委託など）について、以下を抽出してください：
  - 種別・摘要 ➔ 企業名や「給与」「業務委託」などを title に含める
    （括弧形式「給与（会社A）」と、スペース形式「給与 会社A」の両形式に対応）
  - 収入金額 ➔ type: "income_detail", category: "給与" | "業務委託" | "その他"
  - 源泉徴収税額 ➔ type: "income_detail", category: "給与源泉" | "業務委託源泉" など

■ 確定申告書B（第二表）の「雑所得（公的年金等以外）、総合課税の配当所得・譲渡所得」セクションから：
- 「雑所得」の行に「収入金額」「必要経費等」「差引金額」が記載されていれば：
  - 収入金額 ➔ type: "income_detail", category: "雑所得", title: "雑所得収入"
  - 必要経費等 ➔ type: "income_detail", category: "雑所得", title: "雑所得経費"
  - 差引金額 ➔ type: "income_detail", category: "雑所得", title: "雑所得"
- 「配当所得」が記載されていれば ➔ type: "income_detail", category: "配当所得"
- 「譲渡所得」が記載されていれば ➔ type: "income_detail", category: "譲渡所得"

【厳重注意】
- 「収入金額（ア、カなど）」と「所得金額（①、⑥など）」を絶対に混同しないでください。
- 減価償却費などの経費の数値を、誤って給与収入（カ）などに分類しないでください。
- 計算は一切行わず、画像・PDFにある数値を「そのまま」抽出してください。存在しない項目は推測で0を入れず、itemsから除外してください。`;

  const result = callGemini(cleanApiKey, availableModels, priorityModels, base64Data, prompt, actualMimeType);
  return reconcileBusinessExpense(result);
}

/**
 * 収支内訳書が添付されておらず経費の内訳が一件も抽出されなかった場合、
 * 「事業所得＝事業収入－必要経費」という確定した会計上の等式から
 * 必要経費合計を逆算する（推測ではなく、抽出済みの実数値による引き算）
 */
function reconcileBusinessExpense(docResult) {
  if (!docResult || !docResult.items || !docResult.items.length) return docResult;

  const items = docResult.items;
  const sum = (pred) => items.filter(pred).reduce((s, i) => s + (Number(i.amount) || 0), 0);

  const businessRevenue = sum(i => i.type === 'revenue' && i.category === '売上・事業収入');
  const businessIncome = sum(i => i.type === 'revenue' && i.category === '事業所得');
  const existingExpense = sum(i => i.type === 'expense');

  if (businessRevenue > 0 && businessIncome >= 0 && existingExpense === 0) {
    const computedExpense = businessRevenue - businessIncome;
    if (computedExpense > 0) {
      items.push({
        type: 'expense',
        category: '雑費',
        title: '必要経費合計（収入－所得から計算）',
        amount: computedExpense
      });
    }
  }
  return docResult;
}

function extractJson(text) {
  try {
    return JSON.parse(text);
  } catch (e) {
    const jsonMatch = text.match(/\{[\s\S]*\}/) || text.match(/\[[\s\S]*\]/);
    if (jsonMatch) return JSON.parse(jsonMatch[0]);
    throw e;
  }
}

function isBinaryString(str) {
  if (!str) return true;
  let controlChars = 0;
  const checkLength = Math.min(str.length, 500);
  for (let i = 0; i < checkLength; i++) {
    const code = str.charCodeAt(i);
    if (code < 32 && code !== 9 && code !== 10 && code !== 13) {
      controlChars++;
    }
  }
  return (controlChars / checkLength) > 0.08;
}

function callGemini(apiKey, availableModels, priorityModels, base64Data, prompt, mimeType) {
  const actualMimeType = mimeType || "image/png";
  let payload;

  if (actualMimeType === "application/octet-stream" || actualMimeType.startsWith("text/")) {
    let fileText = "";
    try {
      const decodedBytes = Utilities.base64Decode(base64Data.split(',')[1]);
      fileText = Utilities.newBlob(decodedBytes).getDataAsString("UTF-8");
    } catch (e) {
      throw new Error("バイナリデータ形式のため、AIで直接中身を解析できませんでした。PDFファイルをアップロードしてください。");
    }

    if (isBinaryString(fileText)) {
      throw new Error("バイナリデータ形式のため、AIで直接中身を解析できませんでした。PDFファイルをアップロードしてください。");
    }

    const textPrompt = `${prompt}\n\n【解析対象の構造化データ（テキスト）】\n${fileText}`;
    payload = {
      contents: [{ parts: [{ text: textPrompt }] }],
      generationConfig: { temperature: 0, responseMimeType: "application/json" }
    };
  } else {
    payload = {
      contents: [{
        parts: [
          { text: prompt },
          { inlineData: { mimeType: actualMimeType, data: base64Data.split(',')[1] } }
        ]
      }],
      generationConfig: { temperature: 0, responseMimeType: "application/json" }
    };
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let lastError = "";
  let success = false;
  let resultJSON = null;

  if (availableModels && availableModels.length > 0) {
    for (const pModel of priorityModels) {
      const found = availableModels.find(m => m.name.includes(pModel));
      if (!found) continue;

      const targetModel = found.name;
      const url = `https://generativelanguage.googleapis.com/v1beta/${targetModel}:generateContent?key=${apiKey}`;
      
      try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        const result = JSON.parse(responseText);

        if (responseCode === 200 && !result.error) {
          const text = result.candidates[0].content.parts[0].text;
          resultJSON = extractJson(text);
          success = true;
          break;
        } else {
          lastError = result.error ? result.error.message : `HTTP ${responseCode}`;
        }
      } catch (e) {
        lastError = e.toString();
      }
    }
  }

  if (!success) {
    const fallbackModels = ["models/gemini-2.5-flash", "models/gemini-1.5-flash"];
    for (const fbModel of fallbackModels) {
      const url = `https://generativelanguage.googleapis.com/v1beta/${fbModel}:generateContent?key=${apiKey}`;
      try {
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        const result = JSON.parse(responseText);

        if (responseCode === 200 && !result.error) {
          const text = result.candidates[0].content.parts[0].text;
          resultJSON = extractJson(text);
          success = true;
          break;
        } else {
          lastError = result.error ? result.error.message : `HTTP ${responseCode}`;
        }
      } catch (e) {
        lastError = e.toString();
      }
    }
  }

  if (success && resultJSON) {
    return resultJSON;
  }

  throw new Error(`解析処理中にエラーが発生しました: ${lastError || "利用可能なAIモデルが応答しませんでした。"}`);
}

function getData() {
  try {
    const sheet = initSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    const colMap = { date: 1, title: 2, amount: 3, category: 4, type: 5, memo: 6 };
    headers.forEach((h, idx) => {
      const cleanH = h.toString().trim();
      if (cleanH === "日付") colMap.date = idx + 1;
      else if (cleanH === "内容・店名" || cleanH === "店名") colMap.title = idx + 1;
      else if (cleanH === "金額") colMap.amount = idx + 1;
      else if (cleanH === "カテゴリ") colMap.category = idx + 1;
      else if (cleanH === "区分") colMap.type = idx + 1;
      else if (cleanH === "メモ") colMap.memo = idx + 1;
    });

    const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    return values.map((row, index) => {
      const rawDate = row[colMap.date - 1];
      let dateObj = null;
      
      if (rawDate instanceof Date) {
        dateObj = rawDate;
      } else if (rawDate) {
        const dateStr = rawDate.toString().replace(/\//g, "-").trim();
        dateObj = new Date(dateStr);
      }

      const dateStr = (dateObj && !isNaN(dateObj)) ? Utilities.formatDate(dateObj, "JST", "yyyy-MM-dd") : "";
      const year = (dateObj && !isNaN(dateObj)) ? dateObj.getFullYear() : new Date().getFullYear();
      
      const rawType = row[colMap.type - 1] ? row[colMap.type - 1].toString().trim() : "経費";
      let type = "expense";
      if (rawType === "売上" || rawType === "revenue") type = "revenue";
      else if (rawType === "控除" || rawType === "deduction") type = "deduction";
      else if (rawType === "給与" || rawType === "salary") type = "salary";
      else if (rawType === "税金" || rawType === "tax") type = "tax";
      else if (rawType === "所得内訳" || rawType === "income_detail") type = "income_detail";

      return {
        id: index + 2,
        year: year,
        date: dateStr,
        title: row[colMap.title - 1] || "",
        amount: Number(row[colMap.amount - 1]) || 0,
        category: row[colMap.category - 1] || "",
        type: type,
        memo: row[colMap.memo - 1] || ""
      };
    });
  } catch (e) {
    console.error("データ取得中にエラーが発生しました: ", e);
    return [];
  }
}

function addData(item) {
  const sheet = initSheet();
  let typeJp = "経費";
  if (item.type === "revenue") typeJp = "売上";
  else if (item.type === "deduction") typeJp = "控除";
  else if (item.type === "salary") typeJp = "給与";
  else if (item.type === "tax") typeJp = "税金";

  sheet.appendRow([
    item.date,
    item.title || item.shop,
    Number(item.amount),
    item.category,
    typeJp,
    item.memo || "",
    new Date()
  ]);
  return getData();
}

function addBulkData(items) {
  const sheet = initSheet();
  const rows = items.map(item => {
    let typeJp = "経費";
    if (item.type === "revenue") typeJp = "売上";
    else if (item.type === "deduction") typeJp = "控除";
    else if (item.type === "salary") typeJp = "給与";
    else if (item.type === "tax") typeJp = "税金";
    else if (item.type === "income_detail") typeJp = "所得内訳";

    let formattedDate = item.date;
    try { formattedDate = new Date(item.date); } catch(e) {}

    return [
      formattedDate,
      item.title || item.shop,
      Number(item.amount),
      item.category,
      typeJp,
      item.memo || "",
      new Date()
    ];
  });
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  }
  return getData();
}

function updateData(rowIndex, item) {
  const sheet = initSheet();
  let typeJp = "経費";
  if (item.type === "revenue") typeJp = "売上";
  else if (item.type === "deduction") typeJp = "控除";
  else if (item.type === "salary") typeJp = "給与";
  else if (item.type === "tax") typeJp = "税金";
  
  sheet.getRange(rowIndex, 1, 1, 6).setValues([[
    item.date,
    item.title || item.shop,
    Number(item.amount),
    item.category,
    typeJp,
    item.memo || ""
  ]]);
  return getData();
}

function deleteData(rowIndex) {
  const sheet = initSheet();
  sheet.deleteRow(rowIndex);
  return getData();
}