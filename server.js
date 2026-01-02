const express = require("express");
const axios = require("axios");
const xml2js = require("xml2js");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(__dirname));

// reports 目錄
const reportsDir = path.join(__dirname, "reports");
if (!fs.existsSync(reportsDir)) fs.mkdirSync(reportsDir);

// CORS：允許 GitHub Pages 域名及本地端
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*"); // 你也可改成 https://morris8231.github.io
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.sendStatus(204);
  next();
});

// ====== 進度狀態（單一任務版） ======
let clients = new Set();
let isRunning = false;

function createEmptyProgress() {
  return {
    current: 0,
    total: 0,
    percent: 0,
    labels: [],
    counts: [],
    complete: false,
    error: false,
    message: "",
    reportFile: "",
    summaryRowCount: 0,
    rawRowCount: 0,
  };
}

let progressState = createEmptyProgress();

function broadcastProgress() {
  const payload = `data: ${JSON.stringify(progressState)}\n\n`;
  for (const res of clients) {
    try {
      res.write(payload);
    } catch {
      clients.delete(res);
    }
  }
}

function setProgressTotal(total) {
  progressState.total = total;
  progressState.percent = total ? Math.round((progressState.current / total) * 100) : 0;
  broadcastProgress();
}

function updateProgress(label, count, message = "") {
  progressState.current += 1;
  progressState.labels.push(label);
  progressState.counts.push(count);
  progressState.percent = progressState.total
    ? Math.min(100, Math.round((progressState.current / progressState.total) * 100))
    : 0;
  if (message) progressState.message = message;
  broadcastProgress();
}

function markDone(ok = true, message = "", extra = {}) {
  progressState.complete = true;
  progressState.error = !ok;
  if (message) progressState.message = message;
  Object.assign(progressState, extra);
  broadcastProgress();
}

// ====== SSE ======
app.get("/progress", (req, res) => {
  res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
  res.setHeader("Cache-Control", "no-cache, no-transform");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("Access-Control-Allow-Origin", "*");
  // 某些 proxy 會 buffer SSE；加這行可提升穩定性（若環境支援）
  res.setHeader("X-Accel-Buffering", "no");

  res.flushHeaders?.();

  clients.add(res);

  // 先送一次快照
  res.write(`data: ${JSON.stringify(progressState)}\n\n`);

  // keep-alive ping，避免中間層斷線
  const ping = setInterval(() => {
    try {
      res.write(`: ping\n\n`);
    } catch {
      clearInterval(ping);
      clients.delete(res);
    }
  }, 15000);

  req.on("close", () => {
    clearInterval(ping);
    clients.delete(res);
  });
});

// ====== generate ======
app.post("/generate", (req, res) => {
  if (isRunning) return res.status(409).json({ message: "已有報表正在產生中，請稍候。" });

  const { startDate, endDate } = req.body || {};
  if (!startDate || !endDate) return res.status(400).json({ message: "缺少日期" });

  // 每次開始都 reset
  progressState = createEmptyProgress();
  broadcastProgress();

  // （可選）先計算期數 periods，再設置 total；也可以在 generateJob 裡先設
  // const periods = getPeriods(startDate, endDate);
  // setProgressTotal(periods.length);

  // 立刻回應前端：開始了
  res.json({ started: true });

  (async () => {
    isRunning = true;
    try {
      await generateJob(startDate, endDate, { updateProgress, setProgressTotal, markDone });
    } catch (err) {
      console.error(err);
      markDone(false, err?.message || "產生失敗");
    } finally {
      isRunning = false;
    }
  })();
});

app.get("/history", (req, res) => {
  const historyFile = path.join(reportsDir, "history.json");
  if (!fs.existsSync(historyFile)) return res.json([]);
  try {
    const data = JSON.parse(fs.readFileSync(historyFile, "utf8"));
    const cleaned = Array.isArray(data)
      ? data.map((x) => ({
          file: x.file,
          summaryCount: x.summaryCount,
          rawCount: x.rawCount,
          created: x.created,
        }))
      : [];
    res.json(cleaned);
  } catch {
    res.json([]);
  }
});

app.get("/download/:file", (req, res) => {
  const filePath = path.join(reportsDir, req.params.file);
  if (!fs.existsSync(filePath)) return res.status(404).send("檔案不存在");
  res.download(filePath);
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`Server started on port ${PORT}`);
});

// ====== 產生報表邏輯 ======
async function generateJob(startDate, endDate, helpers) {
  const { updateProgress, setProgressTotal, markDone } = helpers;

  // 1) 先算 periods（每月 1 / 16）=> periods[]
  //    例如：
  //    const periods = getPeriods(startDate, endDate);
  //    setProgressTotal(periods.length);

  // 2) 依序處理每個 period：
  //    - fetch 政府資料
  //    - parse XML / JSON
  //    - 加入 ExcelJS workbook
  //    - updateProgress(label, count)

  // 3) 最後寫入報表檔並更新 history.json
  //    markDone(true, "完成", {
  //      reportFile: fileName,
  
  //      summaryRowCount: xxx,
  //      rawRowCount: yyy,
  //    });
}
