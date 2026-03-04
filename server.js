const express = require('express');
const nodemailer = require('nodemailer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const RESULTS_FILE = path.join(__dirname, 'results.json');
const LOCKS_FILE   = path.join(__dirname, 'quiz_locks.json');
const ADMIN_USER   = 'Yael Maimon';
const ADMIN_PASS   = 'Yam_0604';

function loadLocks() {
  try {
    if (fs.existsSync(LOCKS_FILE)) return JSON.parse(fs.readFileSync(LOCKS_FILE, 'utf8'));
  } catch (e) {}
  const locks = {};
  for (let i = 1; i <= 10; i++) locks[i] = false; // default: all locked
  return locks;
}
function saveLocks(locks) {
  fs.writeFileSync(LOCKS_FILE, JSON.stringify(locks, null, 2), 'utf8');
}

function loadResults() {
  try {
    if (fs.existsSync(RESULTS_FILE)) {
      return JSON.parse(fs.readFileSync(RESULTS_FILE, 'utf8'));
    }
  } catch (e) { /* ignore */ }
  return [];
}

function saveResults(results) {
  fs.writeFileSync(RESULTS_FILE, JSON.stringify(results, null, 2), 'utf8');
}

// Submit quiz results
app.post('/api/submit', async (req, res) => {
  try {
    const { studentName, studentId, quizTitle, answers, score, totalQuestions, timestamp } = req.body;

    const result = {
      studentName,
      studentId,
      quizTitle,
      score,
      totalQuestions,
      percentage: Math.round((score / totalQuestions) * 100),
      answers,
      timestamp: timestamp || new Date().toISOString()
    };

    // Save to file
    const results = loadResults();
    results.push(result);
    saveResults(results);

    // Generate Excel and send email if configured
    if (process.env.EMAIL_TO) {
      try {
        await sendEmailWithExcel(result);
      } catch (emailErr) {
        console.error('Email sending failed:', emailErr.message);
      }
    }

    res.json({ success: true, message: 'Results submitted successfully', result });
  } catch (err) {
    console.error('Submit error:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

// Get lock states (public — students poll this)
app.get('/api/locks', (req, res) => {
  res.json(loadLocks());
});

// Update lock states (admin only)
app.post('/api/locks', (req, res) => {
  const { adminUser, adminPass, locks } = req.body;
  if (adminUser !== ADMIN_USER || adminPass !== ADMIN_PASS) {
    return res.status(401).json({ error: 'אין הרשאה' });
  }
  saveLocks(locks);
  res.json({ success: true, locks });
});

// Get all results (for admin)
app.get('/api/results', (req, res) => {
  const results = loadResults();
  res.json(results);
});

// Download all results as Excel
app.get('/api/results/excel', (req, res) => {
  const results = loadResults();
  if (results.length === 0) {
    return res.status(404).json({ message: 'No results yet' });
  }

  const wb = generateExcelWorkbook(results);
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=quiz_results.xlsx');
  res.send(buf);
});

function generateExcelWorkbook(results) {
  const wb = XLSX.utils.book_new();

  // Summary sheet
  const summaryData = results.map(r => ({
    'Student Name': r.studentName,
    'Student ID': r.studentId,
    'Quiz': r.quizTitle,
    'Score': r.score,
    'Total': r.totalQuestions,
    'Percentage': r.percentage + '%',
    'Date': new Date(r.timestamp).toLocaleString('he-IL')
  }));
  const ws1 = XLSX.utils.json_to_sheet(summaryData);
  ws1['!cols'] = [
    { wch: 20 }, { wch: 15 }, { wch: 40 }, { wch: 8 }, { wch: 8 }, { wch: 12 }, { wch: 20 }
  ];
  XLSX.utils.book_append_sheet(wb, ws1, 'Summary');

  // Detailed answers sheet
  const detailData = [];
  results.forEach(r => {
    if (r.answers) {
      r.answers.forEach((a, i) => {
        detailData.push({
          'Student Name': r.studentName,
          'Student ID': r.studentId,
          'Quiz': r.quizTitle,
          'Question #': i + 1,
          'Question': a.question,
          'Student Answer': a.studentAnswer,
          'Correct Answer': a.correctAnswer,
          'Is Correct': a.isCorrect ? 'V' : 'X'
        });
      });
    }
  });
  if (detailData.length > 0) {
    const ws2 = XLSX.utils.json_to_sheet(detailData);
    ws2['!cols'] = [
      { wch: 20 }, { wch: 15 }, { wch: 40 }, { wch: 10 }, { wch: 50 }, { wch: 40 }, { wch: 40 }, { wch: 10 }
    ];
    XLSX.utils.book_append_sheet(wb, ws2, 'Detailed Answers');
  }

  return wb;
}

async function sendEmailWithExcel(result) {
  const transporter = nodemailer.createTransport({
    service: process.env.EMAIL_SERVICE || 'gmail',
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS
    }
  });

  const wb = generateExcelWorkbook([result]);
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: process.env.EMAIL_TO,
    subject: `Quiz Result: ${result.studentName} - ${result.quizTitle} (${result.percentage}%)`,
    html: `
      <div dir="rtl" style="font-family: Arial, sans-serif;">
        <h2>New Quiz Submission</h2>
        <table style="border-collapse: collapse; width: 100%;">
          <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Student</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${result.studentName}</td></tr>
          <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>ID</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${result.studentId}</td></tr>
          <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Quiz</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${result.quizTitle}</td></tr>
          <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Score</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${result.score}/${result.totalQuestions} (${result.percentage}%)</td></tr>
          <tr><td style="padding: 8px; border: 1px solid #ddd;"><strong>Date</strong></td><td style="padding: 8px; border: 1px solid #ddd;">${new Date(result.timestamp).toLocaleString('he-IL')}</td></tr>
        </table>
      </div>
    `,
    attachments: [{
      filename: `quiz_${result.studentId}_${Date.now()}.xlsx`,
      content: buf
    }]
  };

  await transporter.sendMail(mailOptions);
  console.log(`Email sent for ${result.studentName}`);
}

app.listen(PORT, () => {
  console.log(`\n========================================`);
  console.log(`  MIS Quiz Server Running`);
  console.log(`  http://localhost:${PORT}`);
  console.log(`========================================\n`);
  if (!process.env.EMAIL_TO) {
    console.log('  Note: Email not configured.');
    console.log('  Edit .env file to enable email.\n');
  }
});
