/**
 * Node.js URL Checker for Google Sheets
 * 功能：讀取主控表配置，檢查目標試算表 URL，並回寫狀態與產生日誌。
 */

const { google } = require("googleapis");
const axios = require("axios");
const dayjs = require("dayjs");
const path = require("path");

// --- 1. 設定區域 ---
const CONFIG_SPREADSHEET_ID = "1UCkfb57VOi21oj4nTm8RCXCW0z1UH9g7"; // 請替換為你的主控試算表 ID
const CONFIG_SHEET_NAME = "配置";
const KEYWORD = "your_keyword_here"; // 請替換為你的關鍵字
const TIMEOUT_MS = 5000;

// 身份驗證設置
const auth = new google.auth.GoogleAuth({
  keyFile: path.join(__dirname, "credentials.json"),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});
const sheets = google.sheets({ version: "v4", auth });

// --- 2. 輔助函式 ---

// 將數字欄位轉換為 Excel 字母 (例如 1 -> A, 2 -> B)
function columnToLetter(column) {
  let temp,
    letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// 延遲函數 (避免 API 請求過快)
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// 檢查單個 URL 的函數
async function checkUrl(url, headers) {
  if (!url.match(/^https?:\/\//i)) {
    url = "https://" + url;
  }
  try {
    const response = await axios.get(url, {
      headers: headers,
      timeout: TIMEOUT_MS,
      validateStatus: () => true, // 讓 axios 不要在 404 時拋出錯誤，讓我們自己判斷
    });

    const statusCode = response.status;
    const content =
      typeof response.data === "string"
        ? response.data
        : JSON.stringify(response.data);

    if (statusCode === 404 || content.includes(KEYWORD)) {
      return { status: "失效", success: false };
    } else {
      return { status: "成功", success: true };
    }
  } catch (error) {
    console.log(`❌ 偵測到報錯: ${url} -> ${error.message}`);
    return { status: "失效", success: false, error: error.message };
  }
}

// --- 3. 主程式邏輯 ---

async function main() {
  console.log("🚀 開始執行 URL 檢查程式...");
  const now = dayjs();

  // 統計數據初始化 (修正了你原本代碼的變數問題)
  let stats = {
    total: 0,
    success: 0,
    failed: 0,
    failedURLs: [],
    successURLs: [],
    invalidURLs: [],
  };

  const headers = {
    "User-Agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
    Accept:
      "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
  };

  try {
    // 步驟 A: 獲取主控表配置與白名單
    console.log("正在讀取配置與白名單...");

    // 讀取白名單
    const whitelistRes = await sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG_SPREADSHEET_ID,
      range: "白名單!A:A",
    });
    const whitelistRows = whitelistRes.data.values || [];
    const whitelist = whitelistRows
      .map((row) => (row[0] ? row[0].trim().toLowerCase() : ""))
      .filter((url) => url !== "");

    // 讀取配置表 (假設從 A2 開始)
    const configRes = await sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG_SPREADSHEET_ID,
      range: `${CONFIG_SHEET_NAME}!A2:F`,
    });
    const configData = configRes.data.values || [];

    // 步驟 B: 遍歷配置並檢查
    for (const rowConfig of configData) {
      const spreadsheetId = rowConfig[1];
      const sheetName = rowConfig[2];
      const urlColIndex = parseInt(rowConfig[3], 10); // 例如 1 (A欄)
      const outColIndex = parseInt(rowConfig[4], 10); // 例如 2 (B欄)

      // 確保欄位索引是合法的數字，避免誤寫整欄導致資料被覆蓋
      if (
        !spreadsheetId ||
        !sheetName ||
        Number.isNaN(urlColIndex) ||
        Number.isNaN(outColIndex) ||
        urlColIndex < 1 ||
        outColIndex < 1
      ) {
        console.warn(
          `跳過設定：試算表ID=${spreadsheetId}，工作表=${sheetName}，urlColIndex=${rowConfig[3]}, outColIndex=${rowConfig[4]}`,
        );
        continue;
      }

      // 避免不慎把 URL 欄位當成結果欄位寫回去導致資料遺失
      if (urlColIndex === outColIndex) {
        console.warn(
          `跳過設定：urlColIndex 與 outColIndex 相同 (欄位 ${urlColIndex})，會覆寫原本的 URL。請修正設定。`,
        );
        continue;
      }

      console.log(`正在處理: 試算表[${spreadsheetId}] - 工作表[${sheetName}]`);

      try {
        // 讀取目標工作表的全部數據
        const targetSheetRes = await sheets.spreadsheets.values.get({
          spreadsheetId: spreadsheetId,
          range: `${sheetName}!A2:Z`, // 假設數據不超過 Z 欄，可自行調整
        });

        const rows = targetSheetRes.data.values;
        if (!rows || rows.length === 0) continue;

        const updates = []; // 準備批量更新的數據
        const startRow = 2; // 數據從第 2 行開始

        // 逐行檢查
        for (let i = 0; i < rows.length; i++) {
          const currentRowNum = startRow + i;
          // 欄位索引是從 0 開始，所以要減 1
          let url = rows[i][urlColIndex - 1];

          // 用來存入結果的變數
          let resultStatus = "";

          if (!url || typeof url !== "string" || url.trim() === "") {
            updates.push([""]); // 空值清空結果
            continue;
          }

          url = url.trim();
          const lowerUrl = url.toLowerCase();
          const isWhitelisted = whitelist.some((prefix) =>
            lowerUrl.startsWith(prefix),
          );

          if (isWhitelisted) {
            resultStatus = "成功";
            stats.success++;
            stats.successURLs.push([
              spreadsheetId,
              sheetName,
              currentRowNum,
              url,
            ]);
          } else {
            // 格式檢查
            const urlPattern = /^https?:\/\/([\w.-]+)(:\d+)?(\/.*)?$/i;
            if (!urlPattern.test(url) && !url.match(/^https?:\/\//i)) {
              // 簡易寬鬆檢查，若真的格式太差
              stats.invalidURLs.push([
                spreadsheetId,
                sheetName,
                currentRowNum,
                url,
              ]);
              // 這裡不一定要寫入 "失效"，看需求，先保持清空或標記
              resultStatus = "";
            } else {
              // 進行連線檢查
              stats.total++;
              // 這裡使用 await 逐個檢查，避免並發過高被封鎖
              // 如果想要更快，可以使用 Promise.all 分批處理
              process.stdout.write(
                `\r   檢查中 (${i + 1}/${rows.length}): ${url.substring(
                  0,
                  30,
                )}...`,
              );

              // 在 checkUrl 前面加上延遲，例如 1~2 秒
              await delay(1000 + Math.random() * 1000);
              const checkResult = await checkUrl(url, headers);
              resultStatus = checkResult.status;

              if (checkResult.success) {
                stats.success++;
                stats.successURLs.push([
                  spreadsheetId,
                  sheetName,
                  currentRowNum,
                  url,
                ]);
              } else {
                stats.failed++;
                stats.failedURLs.push([
                  spreadsheetId,
                  sheetName,
                  currentRowNum,
                  url,
                ]);
              }
            }
          }
          updates.push([resultStatus]);
        }
        console.log("該表檢查完成，正在寫入結果...\n   ");

        // 批量寫入結果到 Google Sheet (比 GAS 迴圈寫入快非常多)
        // 計算寫入的 Range，例如: 工作表!D2:D100
        const outColLetter = columnToLetter(outColIndex);
        const updateRange = `${sheetName}!${outColLetter}${startRow}:${outColLetter}${
          startRow + updates.length - 1
        }`;

        await sheets.spreadsheets.values.update({
          spreadsheetId: spreadsheetId,
          range: updateRange,
          valueInputOption: "RAW",
          resource: { values: updates },
        });
      } catch (err) {
        console.error(`   處理工作表時發生錯誤: ${err.message}`);
        stats.failedURLs.push([
          "試算表讀取錯誤",
          spreadsheetId,
          sheetName,
          "-",
          err.message,
        ]);
      }
    }

    // 步驟 C: 創建並寫入日誌表
    console.log("正在生成日誌報告...");
    const logSheetName = now.format("YYYY-MM-DD-HH:mm");

    // 1. 檢查日誌表是否存在，若存在需要刪除 (API 比較複雜，這裡採用如果存在就附加後綴的策略，或者先嘗試建立)
    // 為了模擬 GAS 的 deleteSheet，我們需要先獲取 sheetId
    const metaData = await sheets.spreadsheets.get({
      spreadsheetId: CONFIG_SPREADSHEET_ID,
    });
    const existingSheet = metaData.data.sheets.find(
      (s) => s.properties.title === logSheetName,
    );

    if (existingSheet) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: CONFIG_SPREADSHEET_ID,
        resource: {
          requests: [
            { deleteSheet: { sheetId: existingSheet.properties.sheetId } },
          ],
        },
      });
    }

    // 2. 建立新工作表
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: CONFIG_SPREADSHEET_ID,
      resource: {
        requests: [
          {
            addSheet: {
              properties: {
                title: logSheetName,
                gridProperties: { frozenRowCount: 1 },
              },
            },
          },
        ],
      },
    });

    // 3. 準備寫入日誌的數據
    let logData = [];

    // 統計區塊
    logData.push(["檢查日期", now.format("YYYY-MM-DD")]);
    logData.push(["檢查時間", now.format("HH:mm:ss")]);
    logData.push(["總檢查數量", stats.total]);
    logData.push(["成功數量", stats.success]);
    logData.push(["失效數量", stats.failed]);
    logData.push(["網址格式錯誤", stats.invalidURLs.length]);
    logData.push(["", ""]); // 空行
    logData.push(["", ""]); // 空行
    logData.push(["", ""]); // 空行

    // 詳細列表區塊
    if (stats.successURLs.length > 0) {
      logData.push(["✅ 成功", "試算表 ID", "工作表名稱", "行號", "URL"]);
      stats.successURLs.forEach((row) => logData.push(row));
      logData.push(["", "", "", "", ""]); // 空行
    }

    if (stats.failedURLs.length > 0) {
      logData.push(["❌ 失效", "試算表 ID", "工作表名稱", "行號", "URL"]);
      stats.failedURLs.forEach((row) => logData.push(row));
      logData.push(["", "", "", "", ""]); // 空行
    }

    if (stats.invalidURLs.length > 0) {
      logData.push([
        "⚠️ 網址格式錯誤",
        "試算表 ID",
        "工作表名稱",
        "行號",
        "URL",
      ]);
      stats.invalidURLs.forEach((row) => logData.push(row));
    }

    // 4. 寫入日誌數據
    await sheets.spreadsheets.values.update({
      spreadsheetId: CONFIG_SPREADSHEET_ID,
      range: `${logSheetName}!A1`,
      valueInputOption: "RAW",
      resource: { values: logData },
    });

    console.log(`✅ 所有作業完成！日誌已寫入工作表: ${logSheetName}`);
  } catch (error) {
    console.error("❌ 程式執行發生嚴重錯誤:", error);
  }
}

main();
