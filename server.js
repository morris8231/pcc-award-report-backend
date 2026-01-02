// server.js
// 政府採購網 OpenData 決標資料報表產生器
// 將本檔案放到 pcc-award-report-backend 根目錄並 commit 部署

const express = require('express');
const cors = require('cors');
const axios = require('axios');
const xml2js = require('xml2js');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
// Render 會指定 PORT，若本地執行則預設 3000
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

// 報表暫存資料夾
const reportsDir = path.join(__dirname, 'reports');
if (!fs.existsSync(reportsDir)) {
  fs.mkdirSync(reportsDir, { recursive: true });
}

// ========== 進度狀態與 SSE ==========
const clients = new Set();
let isRunning = false;
let progressState = createEmptyProgress();

function createEmptyProgress() {
  return {
    current: 0,
    total: 0,
    percent: 0,
    labels: [],
    counts: [],
    complete: false,
    error: false,
    message: '',
    reportFile: '',
    summaryRowCount: 0,
    rawRowCount: 0
  };
}

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
  progressState.percent = total ? 0 : 100;
  broadcastProgress();
}

function updateProgress(label, count, message = '') {
  progressState.current += 1;
  progressState.labels.push(label);
  progressState.counts.push(count);
  progressState.percent = progressState.total
    ? Math.min(100, Math.round((progressState.current / progressState.total) * 100))
    : 0;
  if (message) progressState.message = message;
  broadcastProgress();
}

function markDone(ok = true, msg = '', extra = {}) {
  progressState.complete = true;
  progressState.error = !ok;
  if (msg) progressState.message = msg;
  Object.assign(progressState, extra);
  progressState.percent = 100;
  broadcastProgress();
}

app.get('/progress', (req, res) => {
  // 設定 SSE header
  res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
  res.setHeader('Cache-Control', 'no-cache, no-transform');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('Access-Control-Allow-Origin', '*');
  // 有些 proxy 如 Cloudflare 可能會 buffer SSE，這行可增加穩定性
  res.setHeader('X-Accel-Buffering', 'no');

  res.flushHeaders?.();

  // 新增客戶端並立即傳送當前進度
  clients.add(res);
  res.write(`data: ${JSON.stringify(progressState)}\n\n`);

  // keep-alive ping，每 15 秒傳一次
  const ping = setInterval(() => {
    try {
      res.write(`: ping\n\n`);
    } catch {
      clearInterval(ping);
      clients.delete(res);
    }
  }, 15000);

  req.on('close', () => {
    clearInterval(ping);
    clients.delete(res);
  });
});

// ========== API ==========
app.post('/generate', (req, res) => {
  if (isRunning) {
    return res.status(409).json({ message: '目前已有報表正在產生中，請稍後再試。' });
  }

  const { startDate, endDate } = req.body || {};
  if (!startDate || !endDate) {
    return res.status(400).json({ message: '缺少 startDate / endDate' });
  }

  // 立即回覆讓前端開始監聽進度
  res.json({ started: true });

  (async () => {
    isRunning = true;
    try {
      await generateJob(startDate, endDate);
    } catch (err) {
      console.error(err);
      markDone(false, err?.message || '產生失敗');
    } finally {
      isRunning = false;
    }
  })();
});

app.get('/history', (req, res) => {
  const historyFile = path.join(reportsDir, 'history.json');
  if (!fs.existsSync(historyFile)) {
    return res.json([]);
  }
  try {
    const data = JSON.parse(fs.readFileSync(historyFile, 'utf8'));
    const cleaned = Array.isArray(data)
      ? data.map(x => ({
          file: x.file,
          summaryCount: x.summaryCount,
          rawCount: x.rawCount,
          created: x.created
        }))
      : [];
    res.json(cleaned);
  } catch {
    res.json([]);
  }
});

app.get('/download/:file', (req, res) => {
  const filePath = path.join(reportsDir, req.params.file);
  if (!fs.existsSync(filePath)) {
    return res.status(404).send('檔案不存在');
  }
  res.download(filePath);
});

app.listen(PORT, () => {
  console.log(`Server started at port ${PORT}`);
});

// ========== 核心流程：產生報表 ==========
async function generateJob(startDate, endDate) {
  const files = buildFileList(startDate, endDate);
  // 重設進度狀態並設定總數
  progressState = createEmptyProgress();
  setProgressTotal(files.length);

  const jsonRecords = [];
  const rawRows = [];
  const bidderMap = Object.create(null);
  const attemptedTokens = [];

  // 依序處理每個檔案
  for (const { fileName } of files) {
    const token = fileNameToToken(fileName);
    attemptedTokens.push(token);

    let xmlText = '';
    try {
      xmlText = await downloadAwardXml(fileName);
    } catch (e) {
      updateProgress(fileName, 0, `下載失敗：${fileName}`);
      continue;
    }

    let xmlObj;
    try {
      xmlObj = await xml2js.parseStringPromise(xmlText, {
        explicitArray: false,
        trim: true,
        mergeAttrs: true
      });
    } catch {
      updateProgress(fileName, 0, `解析失敗：${fileName}`);
      continue;
    }

    const tenders = extractTenders(xmlObj);
    for (const tender of tenders) {
      const bidderName = safeGet(tender, ['BIDDER_LIST', 'BIDDER_SUPP_NAME']) || '';
      const awardDate = safeGet(tender, ['AWARD_NOTICE_DATE']) ||
                        safeGet(tender, ['AWARD_DATE']) || '';
      const tenderNo = safeGet(tender, ['TENDER_NO']) || '';
      const tenderName = safeGet(tender, ['TENDER_NAME']) || '';
      const orgName =
        safeGet(tender, ['ORG_NAME']) ||
        safeGet(tender, ['PROCURING_ENTITY_NAME']) ||
        safeGet(tender, ['PROCURING_ENTITY']) ||
        '';

      const priceText = String(
        safeGet(tender, ['TENDER_AWARD_PRICE']) ||
        safeGet(tender, ['AWARD_PRICE']) ||
        ''
      ).replace(/,/g, '');

      const price = priceText ? Number(priceText) : 0;
      const priceMillion = price ? price / 1_000_000 : 0;

      jsonRecords.push({ sourceFile: fileName, tender });

      rawRows.push({
        sourceFile: fileName,
        tenderNo,
        tenderName,
        orgName,
        bidderName,
        awardDate,
        awardPrice: price,
        awardPriceMillion: Number.isFinite(priceMillion) ? priceMillion : 0
      });

      if (bidderName) {
        if (!bidderMap[bidderName]) {
          bidderMap[bidderName] = { count: 0, sumMillion: 0, latestDate: '', latestPriceMillion: 0 };
        }
        bidderMap[bidderName].count += 1;
        bidderMap[bidderName].sumMillion += priceMillion;
        if (
          awardDate &&
          (!bidderMap[bidderName].latestDate ||
           new Date(awardDate) > new Date(bidderMap[bidderName].latestDate))
        ) {
          bidderMap[bidderName].latestDate = awardDate;
          bidderMap[bidderName].latestPriceMillion = priceMillion;
        }
      }
    }

    updateProgress(fileName, tenders.length, '');
  }

  // 彙整 summary
  const summaryRows = Object.entries(bidderMap).map(([name, info]) => ({
    companyName: name,
    awardNoticeDate: info.latestDate || '',
    latestPriceMillion: Number.isFinite(info.latestPriceMillion)
      ? Number(info.latestPriceMillion.toFixed(1))
      : 0,
    cumulativeCount: info.count,
    cumulativeSumMillion: Number.isFinite(info.sumMillion)
      ? Number(info.sumMillion.toFixed(1))
      : 0
  }));
  summaryRows.sort((a, b) => new Date(b.awardNoticeDate) - new Date(a.awardNoticeDate));

  // 組報表檔名
  const sortedTokens = attemptedTokens.filter(Boolean).sort();
  const startToken = sortedTokens[0] || `report_${Date.now()}`;
  const endToken = sortedTokens[sortedTokens.length - 1] || startToken;
  const baseName = `${startToken}_${endToken}`;

  // 寫入 JSON 檔
  const jsonFileName = `${baseName}.json`;
  fs.writeFileSync(path.join(reportsDir, jsonFileName), JSON.stringify(jsonRecords, null, 2), 'utf8');

  // 建立 Excel 報表
  const workbook = new ExcelJS.Workbook();
  const sheet1 = workbook.addWorksheet('Summary');
  sheet1.columns = [
    { header: '公司名稱', key: 'companyName', width: 40 },
    { header: '決標公告日期', key: 'awardNoticeDate', width: 18 },
    { header: '最新得標價格(百萬)', key: 'latestPriceMillion', width: 20 },
    { header: '累積得標次數', key: 'cumulativeCount', width: 18 },
    { header: '累積得標金額', key: 'cumulativeSumMillion', width: 18 }
  ];
  summaryRows.forEach(r => sheet1.addRow(r));

  const sheet2 = workbook.addWorksheet('RawData');
  sheet2.columns = [
    { header: 'sourceFile', key: 'sourceFile', width: 20 },
    { header: 'tenderNo', key: 'tenderNo', width: 22 },
    { header: 'tenderName', key: 'tenderName', width: 50 },
    { header: 'orgName', key: 'orgName', width: 30 },
    { header: 'bidderName', key: 'bidderName', width: 30 },
    { header: 'awardDate', key: 'awardDate', width: 18 },
    { header: 'awardPrice', key: 'awardPrice', width: 18 },
    { header: 'awardPriceMillion', key: 'awardPriceMillion', width: 18 }
  ];
  rawRows.forEach(r => sheet2.addRow(r));

  const xlsxFileName = `${baseName}.xlsx`;
  await workbook.xlsx.writeFile(path.join(reportsDir, xlsxFileName));

  // 更新歷史紀錄
  const historyFile = path.join(reportsDir, 'history.json');
  let history = [];
  if (fs.existsSync(historyFile)) {
    try {
      history = JSON.parse(fs.readFileSync(historyFile, 'utf8'));
      if (!Array.isArray(history)) history = [];
    } catch {
      history = [];
    }
  }
  history.unshift({
    file: xlsxFileName,
    json: jsonFileName,
    summaryCount: summaryRows.length,
    rawCount: rawRows.length,
    created: new Date().toISOString()
  });
  fs.writeFileSync(historyFile, JSON.stringify(history, null, 2), 'utf8');

  // 設定完成狀態並傳送最後進度
  markDone(true, `完成：${xlsxFileName}（Summary ${summaryRows.length} / Raw ${rawRows.length}）`, {
    reportFile: xlsxFileName,
    summaryRowCount: summaryRows.length,
    rawRowCount: rawRows.length
  });
}

// ========== 工具函式 ==========
function buildFileList(start, end) {
  const s = new Date(start);
  const e = new Date(end);
  if (Number.isNaN(s.getTime()) || Number.isNaN(e.getTime()) || s > e) return [];

  const cur = new Date(s.getFullYear(), s.getMonth(), 1);
  const set = new Map();

  while (cur <= e) {
    const y = cur.getFullYear();
    const m = String(cur.getMonth() + 1).padStart(2, '0');

    const firstStart = new Date(y, cur.getMonth(), 1);
    const firstEnd = new Date(y, cur.getMonth(), 15);
    const secondStart = new Date(y, cur.getMonth(), 16);
    const secondEnd = new Date(y, cur.getMonth() + 1, 0);

    if (rangesOverlap(s, e, firstStart, firstEnd)) {
      set.set(`award_${y}${m}01.xml`, { fileName: `award_${y}${m}01.xml` });
    }
    if (rangesOverlap(s, e, secondStart, secondEnd)) {
      set.set(`award_${y}${m}02.xml`, { fileName: `award_${y}${m}02.xml` });
    }

    cur.setMonth(cur.getMonth() + 1);
  }

  return Array.from(set.values());
}

function rangesOverlap(aStart, aEnd, bStart, bEnd) {
  return aStart <= bEnd && bStart <= aEnd;
}

function fileNameToToken(fileName) {
  return String(fileName || '')
    .replace(/^award_/, '')
    .replace(/\.xml$/i, '');
}

async function downloadAwardXml(fileName) {
  // 依序嘗試 planpe 與 web 主站，通常至少有一個可下載
  const hosts = ['https://planpe.pcc.gov.tw', 'https://web.pcc.gov.tw'];
  let lastErr = null;
  for (const host of hosts) {
    const url = `${host}/tps/tp/OpenData/downloadFile?fileName=${encodeURIComponent(fileName)}`;
    try {
      const resp = await axios.get(url, {
        responseType: 'arraybuffer',
        timeout: 45000,
        maxRedirects: 5,
        validateStatus: s => s >= 200 && s < 400,
        headers: {
          Referer: `${host}/tps/tp/OpenData/showList`,
          'User-Agent': 'Mozilla/5.0 (Node.js)',
          Accept: 'application/xml,text/xml,*/*',
          'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.7',
          'Cache-Control': 'no-cache',
          Pragma: 'no-cache'
        }
      });
      const text = Buffer.from(resp.data).toString('utf8').trim();
      if (text.startsWith('<!DOCTYPE html') || text.includes('<html')) {
        throw new Error(`Got HTML instead of XML from ${host}`);
      }
      return text;
    } catch (e) {
      lastErr = e;
      continue;
    }
  }
  const status = lastErr?.response?.status;
  throw new Error(`下載失敗：${fileName}（planpe/web 都失敗，status=${status ?? 'n/a'}）`);
}

function extractTenders(obj) {
  const results = [];
  (function walk(node) {
    if (!node) return;
    if (Array.isArray(node)) return node.forEach(walk);
    if (typeof node !== 'object') return;
    for (const [k, v] of Object.entries(node)) {
      if (k === 'TENDER') {
        if (Array.isArray(v)) results.push(...v);
        else if (v) results.push(v);
      } else {
        walk(v);
      }
    }
  })(obj);
  return results.filter(Boolean);
}

function safeGet(obj, pathArr) {
  let cur = obj;
  for (const k of pathArr) {
    if (!cur || typeof cur !== 'object') return '';
    cur = cur[k];
  }
  return cur ?? '';
}
