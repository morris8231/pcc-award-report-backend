const express = require('express');
const axios = require('axios');
const xml2js = require('xml2js');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(__dirname));

// reports 目錄
const reportsDir = path.join(__dirname, 'reports');
if (!fs.existsSync(reportsDir)) fs.mkdirSync(reportsDir);

// CORS：允許 GitHub Pages 域名及本地端
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*'); // 可限定為 github.io
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  next();
});

// SSE 客戶端列表
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
  res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.flushHeaders?.();

  clients.push(res);
  sendProgress();
  req.on('close', () => {
    clients = clients.filter((c) => c !== res);
  });
});

app.post('/generate', (req, res) => {
  if (isRunning) return res.status(409).json({ message: '已有報表正在產生中，請稍候。' });

  const { startDate, endDate } = req.body || {};
  if (!startDate || !endDate) return res.status(400).json({ message: '缺少日期' });

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

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server started on port ${PORT}`);
});

// 產生報表邏輯（與您提供的版本一致，欄位名稱已修改）
async function generateJob(startDate, endDate) {
  // ... 保留您之前的程式
}
