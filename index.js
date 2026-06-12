/**
 * Node.js URL Checker for Google Sheets (High-Speed Batched Version)
 * 功能：讀取主控表配置，並行/批量檢查目標試算表 URL，並回寫狀態與產生日誌。
 */

const { google } = require("googleapis");
const axios = require("axios");
const dayjs = require("dayjs");
const path = require("path");

// --- 1. 設定區域 ---
const CONFIG_SPREADSHEET_ID = "1gBhgTPvrYSGujOl6vKVk769NuvchsvJ3EfTLzFNECvE"; // 請替換為你的主控試算表 ID
const CONFIG_SHEET_NAME = "配置";
const KEYWORD = "your_keyword_here"; // 請替換為你的關鍵字
const TIMEOUT_MS = 5000;

// ⚡ 效能與速率限制設定 ⚡
const MAX_CONCURRENT_REQUESTS = 10; // 同時檢查的網址數量 (可依據需求調整，建議 10-20)
const BATCH_DELAY_MS = 500;         // 每批次之間的緩衝時間 (毫秒)，降低被防爬蟲封鎖的機率

// 身份驗證設置
const auth = new google.auth.GoogleAuth({
  keyFile: path.join(__dirname, "credentials.json"),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});
const sheets = google.sheets({ version: "v4", auth });

// --- 2. 輔助函式 ---

// 將數字欄位轉換為 Excel 字母
function columnToLetter(column) {
  let temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// 延遲函數
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
      validateStatus: () => true, // 讓 axios 不要在 404 時拋出錯誤
    });

    const statusCode = response.status;
    const content = typeof response.data === "string" ? response.data : JSON.stringify(response.data);

    if (statusCode === 404 || content.includes(KEYWORD)) {
      return { status: "失效", success: false };
    } else {
      return { status: "成功", success: true };
    }
  } catch (error) {
    return { status: "失效", success: false, error: error.message };
  }
}

// 異步控制核心：將陣列切成小 chunk，並行處理
async function processInBatches(tasks, batchSize, delayMs) {
  const results = [];
  for (let i = 0; i < tasks.length; i += batchSize) {
    const batch = tasks.slice(i, i + batchSize);
    console.log(`\t⚡ 正在並行檢查批次: ${i + 1} ~ ${Math.min(i + batchSize, tasks.length)} / 總數 ${tasks.length}`);
    
    // 同時啟動當前批次內的所有請求
    const batchResults = await Promise.all(batch.map(task => task()));
    results.push(...batchResults);
    
    if (i + batchSize < tasks.length && delayMs > 0) {
      await delay(delayMs); // 批次間的微小停頓
    }
  }
  return results;
}

// --- 3. 主程式邏輯 ---

async function main() {
  console.log("🚀 開始執行高傳輸 URL 檢查程式...");
  const now = dayjs();

  let stats = {
    total: 0,
    success: 0,
    failed: 0,
    failedURLs: [],
    successURLs: [],
    invalidURLs: [],
  };

  const headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
    Accept: "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
  };

  try {
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

    // 讀取配置表
    const configRes = await sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG_SPREADSHEET_ID,
      range: `${CONFIG_SHEET_NAME}!A2:F`,
    });
    const configData = configRes.data.values || [];

    // 遍歷配置並檢查
    for (const rowConfig of configData) {
      const spreadsheetId = rowConfig[1];
      const sheetName = rowConfig[2];
      const urlColIndex = parseInt(rowConfig[3], 10);
      const outColIndex = parseInt(rowConfig[4], 10);

      if (!spreadsheetId || !sheetName || Number.isNaN(urlColIndex) || Number.isNaN(outColIndex) || urlColIndex < 1 || outColIndex < 1) {
        console.warn(`跳過設定：不完整的欄位配置。`);
        continue;
      }

      if (urlColIndex === outColIndex) {
        console.warn(`跳過設定：urlColIndex 與 outColIndex 相同，避免覆寫。`);
        continue;
      }

      console.log(`\n📂 正在處理: 試算表[${spreadsheetId}] - 工作表[${sheetName}]`);

      try {
        const targetSheetRes = await sheets.spreadsheets.values.get({
          spreadsheetId: spreadsheetId,
          range: `${sheetName}!A2:Z`,
        });

        const rows = targetSheetRes.data.values;
        if (!rows || rows.length === 0) continue;

        const startRow = 2;
        const updates = new Array(rows.length).fill([""]); // 先預抓空間
        const taskQueue = []; // 用來存放需要連線連線檢查的任務資訊

        // 第一階段：快速篩選白名單與格式錯誤，建立「需連線檢查」的任務陣列
        for (let i = 0; i < rows.length; i++) {
          const currentRowNum = startRow + i;
          let url = rows[i][urlColIndex - 1];

          if (!url || typeof url !== "string" || url.trim() === "") {
            continue;
          }

          url = url.trim();
          const lowerUrl = url.toLowerCase();
          const isWhitelisted = whitelist.some((prefix) => lowerUrl.startsWith(prefix));

          if (isWhitelisted) {
            updates[i] = ["成功"];
            stats.success++;
            stats.successURLs.push([spreadsheetId, sheetName, currentRowNum, url]);
          } else {
            const urlPattern = /^https?:\/\/([\w.-]+)(:\d+)?(\/.*)?$/i;
            if (!urlPattern.test(url) && !url.match(/^https?:\/\//i)) {
              stats.invalidURLs.push([spreadsheetId, sheetName, currentRowNum, url]);
              updates[i] = [""]; 
            } else {
              stats.total++;
              // 將任務包裹成常規的 Function 推入佇列，但不立刻執行
              taskQueue.push({
                index: i,
                currentRowNum,
                url,
                run: () => checkUrl(url, headers)
              });
            }
          }
        }

        // 第二階段：開始高效並行批量檢查
        if (taskQueue.length > 0) {
          // 轉換為 Promise 執行陣列
          const tasks = taskQueue.map(task => async () => {
            const checkResult = await task.run();
            return { index: task.index, currentRowNum: task.currentRowNum, url: task.url, checkResult };
          });

          // 執行分流批量檢查
          const batchResults = await processInBatches(tasks, MAX_CONCURRENT_REQUESTS, BATCH_DELAY_MS);

          // 彙整並行檢查的結果
          for (const res of batchResults) {
            updates[res.index] = [res.checkResult.status];
            
            if (res.checkResult.success) {
              stats.success++;
              stats.successURLs.push([spreadsheetId, sheetName, res.currentRowNum, res.url]);
            } else {
              stats.failed++;
              stats.failedURLs.push([spreadsheetId, sheetName, res.currentRowNum, res.url, res.checkResult.error || ""]);
            }
          }
        }

        console.log("-> 該表連線檢查完成，正在批次寫入 Google Sheets...");

        // 批量寫入結果到 Google Sheet
        const outColLetter = columnToLetter(outColIndex);
        const updateRange = `${sheetName}!${outColLetter}${startRow}:${outColLetter}${startRow + updates.length - 1}`;

        await sheets.spreadsheets.values.update({
          spreadsheetId: spreadsheetId,
          range: updateRange,
          valueInputOption: "RAW",
          resource: { values: updates },
        });

      } catch (err) {
        console.error(`❌ 處理工作表時發生錯誤: ${err.message}`);
        stats.failedURLs.push(["試算表讀取錯誤", spreadsheetId, sheetName, "-", err.message]);
      }
    }

    // 步驟 C: 創建並寫入日誌表
    console.log("\n📊 正在生成日誌報告...");
    const logSheetName = now.format("YYYY-MM-DD-HH-mm"); // 修正原始碼中冒號在工作表名稱不合法的問題

    const metaData = await sheets.spreadsheets.get({ spreadsheetId: CONFIG_SPREADSHEET_ID });
    const existingSheet = metaData.data.sheets.find((s) => s.properties.title === logSheetName);

    if (existingSheet) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: CONFIG_SPREADSHEET_ID,
        resource: { requests: [{ deleteSheet: { sheetId: existingSheet.properties.sheetId } }] },
      });
    }

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: CONFIG_SPREADSHEET_ID,
      resource: {
        requests: [
          {
            addSheet: {
              properties: { title: logSheetName, gridProperties: { frozenRowCount: 1 } },
            },
          },
        ],
      },
    });

    let logData = [
      ["檢查日期", now.format("YYYY-MM-DD")],
      ["檢查時間", now.format("HH:mm:ss")],
      ["總檢查數量", stats.total],
      ["成功數量", stats.success],
      ["失效數量", stats.failed],
      ["網址格式錯誤", stats.invalidURLs.length],
      ["", ""], ["", ""], ["", ""]
    ];

    if (stats.failedURLs.length > 0) {
      logData.push(["❌ 失效", "試算表 ID", "工作表名稱", "行號", "URL", "錯誤原因"]);
      stats.failedURLs.forEach((row) => logData.push(row));
      logData.push(["", "", "", "", "", ""]);
    }
    
    if (stats.successURLs.length > 0) {
      logData.push(["✅ 成功", "試算表 ID", "工作表名稱", "行號", "URL"]);
      stats.successURLs.forEach((row) => logData.push(row));
      logData.push(["", "", "", "", ""]);
    }

    if (stats.invalidURLs.length > 0) {
      logData.push(["⚠️ 網址格式錯誤", "試算表 ID", "工作表名稱", "行號", "URL"]);
      stats.invalidURLs.forEach((row) => logData.push(row));
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId: CONFIG_SPREADSHEET_ID,
      range: `${logSheetName}!A1`,
      valueInputOption: "RAW",
      resource: { values: logData },
    });

    console.log(`\n🎉 所有作業完成！詳細日誌已記錄在工作表: ${logSheetName}`);
  } catch (error) {
    console.error("❌ 程式執行發生嚴重錯誤:", error);
  }
}

main();
