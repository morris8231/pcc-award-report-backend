// server.js
// 政府採購網 OpenData 決標資料：award_YYYYMM01.xml / award_YYYYMM02.xml
// 改良版：適用於 Render (加入 CORS 與 PORT 設定)

const express = require('express');
const axios = require('axios');
const xml2js = require('xml2js');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const cors = require('cors'); // 新增 CORS

const app = express();
// Render 會自動分配 PORT，若無則用 3000
constjh PORT = process.env.PORT || 3000;

app.use(cors()); // 允許跨域請求 (Frontend 和 Backend 不同網址時必須)
app.use(express.json());
// 雲端版後端不需要 serve static，這行可以保留但不重要
app.use(express.static(__dirname));

// reports dir
// 注意：Render 免費版硬碟是暫時的，重啟後檔案會消失，但執行期間是可以用的
const reportsDir = path.join(__dirname, 'reports');
if (!fs.existsSync(reportsDir)) fs.mkdirSync(reportsDir, { recursive: true });

// ======================
// SSE 進度
// ======================
let clients = [];
let isRunning = false;

let progressState = {
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

function sendProgress() {
  const payload = `data: ${JSON.stringify(progressState)}\n\n`;
  clients.forEach((c) => c.write(payload));
}

function updateProgress(label, count, message = '') {
  progressState.current += 1;
  progressState.labels.push(label);
  progressState.counts.push(count);
  progressState.percent = progressState.total
    ? Math.min(100, Math.round((progressState.current / progressState.total) * 100))
    : 100;
  if (message) progressState.message = message;
  sendProgress();
}

app.get('/progress', (req, res) => {
  res.set({
    'Content-Type': 'text/event-stream; charset=utf-8',
    'Cache-Control': 'no-cache',
    Connection: 'keep-alive'
  });
  res.flushHeaders?.();

  clients.push(res);
  sendProgress();

  req.on('close', () => {
    clients = clients.filter((c) => c !== res);
  });
});

// ======================
// API
// ======================
app.post('/generate', (req, res) => {
  if (isRunning) return res.status(409).json({ message: '目前已有報表正在產生中，請稍後再試。' });

  const { startDate, endDate } = req.body || {};
  if (!startDate || !endDate) return res.status(400).json({ message: '缺少 startDate / endDate' });

  // 立刻回應，讓前端開始聽 SSE
  res.json({ started: true });

  (async () => {
    isRunning = true;
    try {
      await generateJob(startDate, endDate);
    } catch (err) {
      console.error(err);
      progressState.complete = true;
      progressState.error = true;
      progressState.message = err?.message || '產生失敗';
      sendProgress();
    } finally {
      isRunning = false;
    }
  })();
});

app.get('/history', (req, res) => {
  const historyFile = path.join(reportsDir, 'history.json');
  if (!fs.existsSync(historyFile)) return res.json([]);
  try {
    const data = JSON.parse(fs.readFileSync(historyFile, 'utf8'));
    const cleaned = Array.isArray(data)
      ? data.map((x) => ({
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
  if (!fs.existsSync(filePath)) return res.status(404).send('檔案不存在');
  res.download(filePath);
});

// 這裡要用 listen PORT (變數)
app.listen(PORT, () => {
  console.log(`Server started at port ${PORT}`);
});

// ======================
// 核心流程 (保持不變)
// ======================
async function generateJob(startDate, endDate) {
  const files = buildFileList(startDate, endDate);

  progressState = {
    current: 0,
    total: files.length,
    percent: files.length ? 0 : 100,
    labels: [],
    counts: [],
    complete: false,
    error: false,
    message: '開始處理…',
    reportFile: '',
    summaryRowCount: 0,
    rawRowCount: 0
  };
  sendProgress();

  const jsonRecords = [];
  const rawRows = [];
  const bidderMap = Object.create(null);
  const attemptedTokens = []; 

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
      const awardDate = safeGet(tender, ['AWARD_NOTICE_DATE']) || safeGet(tender, ['AWARD_DATE']) || '';
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
        awardPriceMillion: Number.isFinite(priceMillion) ? Number(priceMillion.toFixed(6)) : 0
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

  const summaryRows = Object.entries(bidderMap).map(([name, info]) => ({
    companyName: name,
    awardNoticeDate: info.latestDate || '',
    latestPriceMillion: Number.isFinite(info.latestPriceMillion) ? Number(info.latestPriceMillion.toFixed(1)) : 0,
    cumulativeCount: info.count,
    cumulativeSumMillion: Number.isFinite(info.sumMillion) ? Number(info.sumMillion.toFixed(1)) : 0
  }));
  summaryRows.sort((a, b) => new Date(b.awardNoticeDate) - new Date(a.awardNoticeDate));

  const sortedTokens = attemptedTokens.filter(Boolean).sort(); 
  const startToken = sortedTokens[0] || `report_${Date.now()}`;
  const endToken = sortedTokens[sortedTokens.length - 1] || startToken;
  const baseName = `${startToken}_${endToken}`;

  const jsonFileName = `${baseName}.json`;
  fs.writeFileSync(path.join(reportsDir, jsonFileName), JSON.stringify(jsonRecords, null, 2), 'utf8');

  const workbook = new ExcelJS.Workbook();

  const sheet1 = workbook.addWorksheet('Summary');
  sheet1.columns = [
    { header: '公司名稱', key: 'companyName', width: 40 },
    { header: '決標公告日期', key: 'awardNoticeDate', width: 18 },
    { header: '最新得標價格(百萬)', key: 'latestPriceMillion', width: 20 },
    { header: '累積得標次數', key: 'cumulativeCount', width: 18 },
    { header: '累積得標金額', key: 'cumulativeSumMillion', width: 18 }
  ];
  summaryRows.forEach((r) => sheet1.addRow(r));

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
  rawRows.forEach((r) => sheet2.addRow(r));

  const xlsxFileName = `${baseName}.xlsx`;
  await workbook.xlsx.writeFile(path.join(reportsDir, xlsxFileName));

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

  progressState.complete = true;
  progressState.error = false;
  progressState.percent = 100;
  progressState.reportFile = xlsxFileName;
  progressState.summaryRowCount = summaryRows.length;
  progressState.rawRowCount = rawRows.length;
  progressState.message = `完成：${xlsxFileName}（Summary ${summaryRows.length} / Raw ${rawRows.length}）`;
  sendProgress();
}

function buildFileList(start, end) {
  const s = new Date(start);
  const e = new Date(end);
  if (Number.isNaN(s.getTime()) || Number.isNaN(e.getTime()) || s > e) return [];

  const cur = new Date(s.getFullYear(), s.getMonth(), 1);
  const set = new Map();

  while (cur <= e) {
    const y = cur.getFullYear();
    const m = String(cur.getMonth() + 1).padStart(2, '0');

    const firstHalfStart = new Date(y, cur.getMonth(), 1);
    const firstHalfEnd = new Date(y, cur.getMonth(), 15);
    const secondHalfStart = new Date(y, cur.getMonth(), 16);
    const secondHalfEnd = new Date(y, cur.getMonth() + 1, 0);

    if (rangesOverlap(s, e, firstHalfStart, firstHalfEnd)) {
      set.set(`award_${y}${m}01.xml`, { fileName: `award_${y}${m}01.xml` });
    }
    if (rangesOverlap(s, e, secondHalfStart, secondHalfEnd)) {
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
  const hosts = ['https://planpe.pcc.gov.tw', 'https://web.pcc.gov.tw'];
  let lastErr = null;

  for (const host of hosts) {
    const url = `${host}/tps/tp/OpenData/downloadFile?fileName=${encodeURIComponent(fileName)}`;
    try {
      const resp = await axios.get(url, {
        responseType: 'arraybuffer',
        timeout: 45000,
        maxRedirects: 5,
        validateStatus: (s) => s >= 200 && s < 400,
        headers: {
          Referer: `${host}/tps/tp/OpenData/showList`,
          'User-Agent':
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
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
