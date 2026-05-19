require('dotenv').config({ override: true });

const path      = require('path');
const fs        = require('fs');
const os        = require('os');
const { execFile } = require('child_process');
const express   = require('express');
const session   = require('express-session');
const multer    = require('multer');
const mammoth   = require('mammoth');
const XLSX      = require('xlsx');
const { google } = require('googleapis');
const Anthropic  = require('@anthropic-ai/sdk');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const { v4: uuidv4 } = require('uuid');

// ── multer（メモリ上で処理、最大5ファイル×30MB）──
const uploadFiles = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 30 * 1024 * 1024, files: 5 },
});

const app  = express();
const PORT = process.env.PORT || 3000;
const ROOT = __dirname;
const OUTPUT_DIR = path.join(ROOT, 'output');
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

// ── ミドルウェア ──
app.use(express.json({ limit: '5mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: process.env.SESSION_SECRET || 'dev-secret-change-me',
  resave: false,
  saveUninitialized: false,
  cookie: { httpOnly: true, maxAge: 1000 * 60 * 60 * 8 },
}));
app.use(express.static(path.join(ROOT, 'public')));

// ── LLM（Gemini優先、なければAnthropic） ──
const USE_GEMINI = !!process.env.GEMINI_API_KEY;
const anthropic = USE_GEMINI ? null : new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const gemini    = USE_GEMINI ? new GoogleGenerativeAI(process.env.GEMINI_API_KEY) : null;
const CLAUDE_MODEL = 'claude-sonnet-4-20250514';
// 混雑時フォールバック用の複数モデル
const GEMINI_MODELS = ['gemini-2.5-flash', 'gemini-2.5-flash-lite', 'gemini-2.0-flash'];

async function callLLM({ systemPrompt, userPrompt, maxTokens = 8000, json = false }) {
  if (USE_GEMINI) {
    let lastErr;
    for (const modelName of GEMINI_MODELS) {
      try {
        const genCfg = {
          maxOutputTokens: maxTokens,
          temperature: 0.3,
        };
        // JSON モード：構造化出力を強制
        if (json) genCfg.responseMimeType = 'application/json';
        // Gemini 2.5 系の「思考」を無効化（構造化出力では不要で、トークンを消費して本文が切れる原因になる）
        if (modelName.startsWith('gemini-2.5')) {
          genCfg.thinkingConfig = { thinkingBudget: 0 };
        }
        const model = gemini.getGenerativeModel({
          model: modelName,
          systemInstruction: systemPrompt,
          generationConfig: genCfg,
        });
        const result = await model.generateContent(userPrompt);
        const resp = result.response;
        const text = resp.text();
        const finishReason = resp.candidates?.[0]?.finishReason;
        console.log(`[Gemini] model=${modelName} finish=${finishReason} chars=${text.length}`);
        if (!text || !text.trim()) {
          // 空レスポンス（safetyブロック or MAX_TOKENS）
          throw new Error(`Gemini empty response (finish=${finishReason})`);
        }
        if (finishReason === 'MAX_TOKENS') {
          console.warn(`[Gemini] ${modelName} output truncated (MAX_TOKENS). length=${text.length}`);
          // 途中で切れている → 次モデルへフォールバックせず、その文字列を返してJSON修復に任せる
        }
        return text;
      } catch (e) {
        lastErr = e;
        const msg = String(e?.message || e);
        if (/503|429|unavailable|overload|quota|high demand|empty response/i.test(msg)) {
          console.warn(`[Gemini] ${modelName} failed (${msg.slice(0,100)}), trying next...`);
          continue;
        }
        throw e;
      }
    }
    throw lastErr || new Error('All Gemini models failed');
  }
  const msg = await anthropic.messages.create({
    model: CLAUDE_MODEL,
    max_tokens: maxTokens,
    system: systemPrompt,
    messages: [{ role: 'user', content: userPrompt }],
  });
  return msg.content.filter(b => b.type === 'text').map(b => b.text).join('');
}

function hasLLMKey() {
  return !!(process.env.GEMINI_API_KEY || process.env.ANTHROPIC_API_KEY);
}

/**
 * LLMが返すJSONを緩くパースする。
 * - ```json ... ``` コードブロック除去
 * - 先頭/末尾のゴミテキスト除去
 * - MAX_TOKENSで途中切断された場合、開き括弧を数えて補完して救済
 */
function parseJsonLoose(raw) {
  if (!raw || !raw.trim()) throw new Error('empty response');
  let s = raw.replace(/```(?:json)?\s*/gi, '').replace(/```/g, '').trim();
  // 最初の { から最後の } までを抽出（完全な場合）
  const first = s.indexOf('{');
  if (first === -1) throw new Error('no { in response');
  const last = s.lastIndexOf('}');
  if (last > first) {
    const candidate = s.slice(first, last + 1);
    try { return JSON.parse(candidate); } catch {}
  }
  // 途中切断救済：括弧を数えて不足分を補完
  let body = s.slice(first);
  let depth = 0, inStr = false, esc = false, bracketDepth = 0;
  let lastValidEnd = -1;
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (esc) { esc = false; continue; }
    if (ch === '\\') { esc = true; continue; }
    if (ch === '"') inStr = !inStr;
    if (inStr) continue;
    if (ch === '{') depth++;
    else if (ch === '}') { depth--; if (depth === 0) lastValidEnd = i; }
    else if (ch === '[') bracketDepth++;
    else if (ch === ']') bracketDepth--;
  }
  if (lastValidEnd > 0) {
    try { return JSON.parse(body.slice(0, lastValidEnd + 1)); } catch {}
  }
  // 補完試行：文字列中なら閉じ、配列・オブジェクトも閉じる
  let repair = body;
  if (inStr) repair += '"';
  while (bracketDepth-- > 0) repair += ']';
  while (depth-- > 0) repair += '}';
  // 末尾が ", " や ": " のまま切れてる場合は削除してから閉じ直す
  repair = repair.replace(/,\s*([}\]])/g, '$1');
  try { return JSON.parse(repair); } catch (e) {
    throw new Error('JSON parse failed after repair: ' + e.message);
  }
}

// ── Google OAuth ──
function buildOAuth() {
  return new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI,
  );
}
const GOOGLE_SCOPES = [
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/userinfo.email',
];

// ════════════════════════════════════════════
// ユーティリティ
// ════════════════════════════════════════════

function runPython(inputData) {
  return new Promise((resolve, reject) => {
    const tmpFile = path.join(os.tmpdir(), `ca_${uuidv4()}.json`);
    const sessionId = uuidv4().slice(0, 8);
    const outDir = path.join(OUTPUT_DIR, sessionId);
    fs.mkdirSync(outDir, { recursive: true });

    inputData.outputDir  = outDir;
    inputData.sessionId  = sessionId;
    fs.writeFileSync(tmpFile, JSON.stringify(inputData, null, 2), 'utf8');

    const script = path.join(ROOT, 'generators', 'generate_docs.py');
    const pythonBin = process.env.PYTHON_BIN || 'python3';
    execFile(pythonBin, [script, tmpFile], { timeout: 120_000 }, (err, stdout, stderr) => {
      fs.unlinkSync(tmpFile);
      if (err) return reject(new Error(stderr || err.message));
      try {
        const result = JSON.parse(stdout.trim());
        result.forEach(r => { r.sessionId = sessionId; });
        resolve(result);
      } catch (e) {
        reject(new Error(`Python出力のパース失敗: ${stdout}`));
      }
    });
  });
}

function formatStamp(d) {
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth()+1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
}

// ════════════════════════════════════════════
// ファイルからテキスト抽出ユーティリティ
// ════════════════════════════════════════════

async function extractTextFromBuffer(buffer, mimetype, originalname) {
  const lname = originalname.toLowerCase();

  // ── PDF ──
  if (mimetype === 'application/pdf' || lname.endsWith('.pdf')) {
    const pdfParse = require('pdf-parse');
    const result = await pdfParse(buffer);
    return { type: 'PDF', name: originalname, text: result.text.trim().slice(0, 18000) };
  }

  // ── Word (.docx) ──
  if (lname.endsWith('.docx') ||
      mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
    const result = await mammoth.extractRawText({ buffer });
    return { type: 'Word', name: originalname, text: result.value.trim().slice(0, 18000) };
  }

  // ── Excel (.xlsx / .xls) ──
  if (lname.endsWith('.xlsx') || lname.endsWith('.xls') ||
      mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      mimetype === 'application/vnd.ms-excel') {
    const wb = XLSX.read(buffer, { type: 'buffer', cellText: true });
    let text = '';
    for (const sname of wb.SheetNames) {
      // 設定シート（ドロップダウン用）はスキップ
      if (sname === '設定' || sname === 'Sheet1' || sname === 'Sheet2' || sname === 'Sheet3') continue;
      const ws = wb.Sheets[sname];
      if (!ws['!ref']) continue;
      const range = XLSX.utils.decode_range(ws['!ref']);
      text += `【シート: ${sname}】\n`;
      for (let r = range.s.r; r <= range.e.r; r++) {
        const cols = [];
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (cell && cell.v !== undefined && String(cell.v).trim()) {
            cols.push(String(cell.v).trim().replace(/\n/g, ' '));
          }
        }
        if (cols.length) text += cols.join(' | ') + '\n';
      }
      text += '\n';
    }
    return { type: 'Excel', name: originalname, text: text.trim().slice(0, 18000) };
  }

  // ── テキスト（Zoom文字起こし .txt / .vtt / .srt） ──
  if (lname.endsWith('.txt') || lname.endsWith('.vtt') || lname.endsWith('.srt') ||
      mimetype?.startsWith('text/')) {
    let text = buffer.toString('utf8').trim();
    // VTT/SRT のタイムコード行を除去して読みやすくする
    if (lname.endsWith('.vtt') || lname.endsWith('.srt')) {
      text = text
        .replace(/WEBVTT\n*/g, '')
        .replace(/^\d+\n/gm, '')
        .replace(/\d{2}:\d{2}:\d{2}[.,]\d{3}\s*-->\s*\d{2}:\d{2}:\d{2}[.,]\d{3}\n?/gm, '')
        .replace(/\n{3,}/g, '\n\n');
    }
    return { type: 'テキスト', name: originalname, text: text.slice(0, 20000) };
  }

  return { type: '不明', name: originalname, text: '' };
}

// ════════════════════════════════════════════
// POST /api/parse-files  複数ファイル → AIでフォーム自動入力
// ════════════════════════════════════════════

app.post('/api/parse-files', uploadFiles.array('files', 5), async (req, res) => {
  try {
    if (!req.files?.length) return res.status(400).json({ error: 'ファイルが添付されていません。' });
    if (!hasLLMKey()) {
      return res.status(500).json({ error: 'GEMINI_API_KEY または ANTHROPIC_API_KEY を .env に設定してください。' });
    }

    // ── 各ファイルのテキスト抽出 ──
    const extracted = [];
    for (const f of req.files) {
      try {
        const result = await extractTextFromBuffer(f.buffer, f.mimetype, f.originalname);
        if (result.text) extracted.push(result);
      } catch (e) {
        extracted.push({ type: '読取失敗', name: f.originalname, text: '' });
      }
    }

    const totalText = extracted.map(e => `=== ${e.type}「${e.name}」===\n${e.text}`).join('\n\n');
    if (!totalText.trim()) return res.status(400).json({ error: 'テキストを抽出できませんでした。' });

    // ── Claude に構造化データ抽出を依頼 ──
    const systemPrompt = `あなたは日本の人材紹介会社で使う採用管理ツールのデータ入力アシスタントです。
提供されたファイル群（履歴書・職務経歴書・ヒヤリングシート・面談/Zoom文字起こし等）から
求職者情報を読み取り、指定のJSONスキーマに変換してください。

【基本ルール】
・情報がないフィールドは空文字・null・空配列（推測は禁止。ただし下記の「補完すべき項目」は例外）
・日付は西暦数値（例: birthYear: 1995）
・テキストは自然な日本語で整形

【補完すべき項目（不明でも必ず埋める）】
・candidateNameKana: 氏名のフリガナ（**ひらがな、姓と名の間は半角スペース**）
  不明でも漢字から最も一般的な読みを推定（例「夕菜」は「ゆな」が最も一般的。「ゆうな」ではない）
  例「早川 夕菜」→「はやかわ ゆな」、「田中 太郎」→「たなか たろう」
・addressKana: 現住所の読み（**ひらがな、県/市/町の区切りに半角スペース**）。番地は含めない
  地名の濁点・長音も正確に（例「吉作」→「よしづくり」、「寄居」→「よりい」、「三田」→「みた」）
  例「富山県富山市吉作486-177」→「とやまけん とやまし よしづくり」
  例「東京都新宿区西新宿1-2-3」→「とうきょうと しんじゅくく にししんじゅく」
・nearestLine: 住所から推定される最寄り路線名（例「あいの風とやま鉄道線」「JR高崎線」）。末尾に「線」を必ず付ける
・nearestStation: 最寄駅名（例「呉羽駅」「新町駅」）。末尾に「駅」を必ず付ける

【推測してはいけない項目（資料にない場合は空）】
・commuteTime: 通勤時間（資料に明記がない限り空。推測で埋めない）

【学歴・免許の注意】
・educationHistory: **古い順→新しい順**（時系列昇順）で返す。入学→卒業の順
・学校名は資料のまま（例「私立高岡向陵高等学校」の「私立」を省略しない）
・licenses: 資料に記載があれば正式名称で（例「普通自動車第一種運転免許」）、取得年月も licenses の代わりに licenseHistory に記載
・licenseHistory: [{ "year": 2025, "month": 3, "name": "普通自動車第一種運転免許 取得" }] 形式で年月付き免許を返す（任意）

※ 最寄駅・路線は郵便番号や住所から日本の地理知識で最も妥当なものを1つ推定すること。複数候補あるなら最も一般的な通勤利用駅。
※ ふりがな系は「わからない」で空欄にせず、漢字から合理的に推定した**ひらがな**を必ず入れること。カタカナではなく必ずひらがなで出力。`;

    const userPrompt = `以下のファイル群から求職者情報を抽出し、JSONのみを返してください（コードブロック不要）。

${totalText.slice(0, 40000)}

# 出力スキーマ（このキー名・型を厳守）
{
  "candidateName": "氏名（漢字）",
  "candidateNameKana": "フリガナ（ひらがな）",
  "gender": "女性 または 男性 または 空文字",
  "birthYear": 1995,
  "birthMonth": 5,
  "birthDay": 10,
  "age": 29,
  "lastEducation": "最終学歴（例: ○○大学 ○○学部 卒）",
  "email": "メアド",
  "mobile": "電話番号",
  "zipCode": "123-4567",
  "address": "住所（〒除く、番地・建物名含む完全な住所）",
  "addressKana": "住所フリガナ（ひらがな・番地手前まで）",
  "addressMain": "住所（番地手前まで）",
  "workHistory": [
    {
      "company": "会社名",
      "employmentType": "正社員 / 派遣社員 / 契約社員 / アルバイト / パート のどれか",
      "startYear": 2020, "startMonth": 4,
      "endYear": 2023, "endMonth": 3,
      "isCurrent": false,
      "businessContent": "事業内容（業種・サービス概要）",
      "jobContent": "業務内容（必ず3点の箇条書き。各行頭に'・'を付け、\\nで区切る。例: '・配車計画の立案と管理\\n・ドライバーへの指示・進捗確認\\n・顧客対応（電話応対・配送依頼受付）'。短い記述や1語の業務（例：「配車業務」のみ）の場合は、その業務から合理的に連想できる具体的タスクを3点に展開すること。推測を広げすぎない範囲で自然な関連業務を列挙する。元資料に既に複数業務がある場合はそれを優先して3点にまとめる)",
      "tenure": "3年"
    }
  ],
  "educationHistory": [
    { "year": 2018, "month": 3, "content": "○○大学 ○○学部 卒業" }
  ],
  "licenses": ["普通自動車第一種運転免許", "MOS Excel"],
  "licenseHistory": [
    { "year": 2025, "month": 3, "name": "普通自動車第一種運転免許 取得" }
  ],
  "excelSkills": ["入力","表作成","グラフ","SUM/AVE","IF関数","VLOOKUP","ピボット","マクロ","VBA"] のうち該当するもの,
  "excelLevel": "実務 または 独学 または 経験なし または 空文字",
  "wordSkills": ["入力","文書作成","編集","作表","書式設定","表/図挿入","差込印刷"] のうち該当するもの,
  "wordLevel": "実務 または 独学 または 経験なし または 空文字",
  "phoneExperience": "社外経験あり / 社内経験あり / 経験なし / 空文字",
  "emailExperience": "Outlook経験有 / その他経験有 / 経験なし / 空文字",
  "resignReason": "転職理由（面談メモ・文字起こしから）",
  "whyOffice": "なぜ事務を希望するか",
  "strengths": "強み（面談メモ・文字起こしから）",
  "weaknesses": "弱み",
  "priorities": ["給与・年収アップ","ワークライフバランス","土日祝休み","残業少なめ","研修・教育制度","キャリアアップ","安定性・大手","福利厚生の充実","在宅・リモートワーク","職場環境・人間関係","正社員登用あり","やりがい・成長"] のうち該当するもの（最大5つ）,
  "careerVision": "キャリアビジョン",
  "availableFrom": "即日 / 1ヶ月以内 / 2ヶ月以内 / 3ヶ月以内 / 6ヶ月以内 / 要相談 / 空文字",
  "currentSalary": "現年収・月給（文字列）",
  "desiredSalary": "希望年収（文字列）",
  "desiredArea": "希望勤務地",
  "nearestLine": "最寄り線名",
  "nearestStation": "最寄り駅名",
  "commuteTime": 30,
  "relocationPlan": false,
  "jobPreference": "就業条件の補足（自由記述）",
  "recommendations": [],
  "interviewDates": [
    { "date": "YYYY-MM-DD", "timeStart": "HH:MM (24h・15分刻みのみ：00/15/30/45 のいずれか。例 10:00, 10:15, 10:30, 10:45)", "timeEnd": "HH:MM (同上の15分刻み)", "note": "備考（任意）" }
  ],
  "notes": "備考・面談メモ（整理した内容）"
}`;

    const raw = (await callLLM({ systemPrompt, userPrompt, maxTokens: 8192, json: true })).trim();
    let parsed;
    try {
      parsed = parseJsonLoose(raw);
    } catch (e) {
      console.error('[parse-files] JSON parse failed. Raw (head):', raw.slice(0, 500));
      console.error('[parse-files] Raw (tail):', raw.slice(-500));
      throw new Error('AIがJSONを返せませんでした。（モデル応答を破棄）詳細はサーバーログ参照。');
    }

    // AIが返した面接日時を15分刻みに丸め込み（保険）
    if (parsed && Array.isArray(parsed.interviewDates)) {
      parsed.interviewDates = parsed.interviewDates.map(iv => ({
        ...iv,
        timeStart: snapTo15Min(iv.timeStart || ''),
        timeEnd:   snapTo15Min(iv.timeEnd   || ''),
      }));
    }

    return res.json({
      ok: true,
      formData: parsed,
      filesSummary: extracted.map(e => ({ name: e.name, type: e.type, chars: e.text.length })),
    });
  } catch (err) {
    console.error('[parse-files] error:', err);
    return res.status(500).json({ error: 'ファイル解析に失敗しました: ' + err.message });
  }
});

// ════════════════════════════════════════════
// AI ナレーティブ生成
// ════════════════════════════════════════════

async function generateNarrative(formData) {
  const {
    candidateName, age, lastEducation, workHistory = [],
    excelSkills = [], wordSkills = [], phoneExperience, emailExperience,
    strengths, weaknesses, resignReason, whyOffice,
    priorities = [], careerVision, currentSalary, desiredSalary,
    recommendations = [],
  } = formData;

  const workSummary = workHistory
    .map(w => `${w.company}（${w.startYear}年${w.startMonth}月〜${w.isCurrent ? '現在' : `${w.endYear}年${w.endMonth}月`}）${w.employmentType} / ${w.businessContent} / ${w.jobContent}`)
    .join('\n');

  const clientLabel = recommendations.length > 0 ? recommendations.join('・') : '一般派遣';

  const systemPrompt = `あなたは日本の人材紹介会社の経験豊富なキャリアアドバイザーです。
事務職希望の求職者情報から、履歴書・職務経歴書に使える日本語文書を作成してください。
自然なビジネス日本語で、誇張なく、実績や数字は与えられた情報のみを使ってください。

【重要ルール】
・workSummary（職務要約）と selfPr（自己PR）は役割を明確に分けること
  - workSummary = **事実ベースの要約のみ**（何社でどんな業種に何年、どの業務を担当したかを客観的に記述）
    自己PRっぽい表現（「強みを活かして〜」「〜に貢献できます」「〜力があります」等）は絶対に書かない
  - selfPr = アピール文（強み・姿勢・貢献意欲を含む）
・workSummary は「〜年間、〜として〜業務に従事」のような事実羅列にする`;

  const userPrompt = `以下の情報から3つの文章を作成してください。JSON形式で返答してください。

# 求職者情報
- 氏名: ${candidateName}
- 年齢: ${age}歳
- 最終学歴: ${lastEducation || '未記入'}
- 職歴:
${workSummary || '未記入'}
- Excelスキル: ${excelSkills.join('・') || '未記入'}
- Wordスキル: ${wordSkills.join('・') || '未記入'}
- 電話応対: ${phoneExperience || '未記入'}
- ビジネスメール: ${emailExperience || '未記入'}
- 強み: ${strengths || '未記入'}
- 弱み: ${weaknesses || '未記入'}
- 転職理由: ${resignReason || '未記入'}
- なぜ事務を希望: ${whyOffice || '未記入'}
- 転職優先条件: ${priorities.join('・') || '未記入'}
- キャリアビジョン: ${careerVision || '未記入'}
- 現年収: ${currentSalary || '未記入'}
- 希望年収: ${desiredSalary || '未記入'}
- 提案先: ${clientLabel}

# 出力仕様（JSON・コードブロック不要）
{
  "workSummary": "職務要約（150〜300文字。事実のみ。『〜社で〜年間、〜業務に従事』のような客観的事実の羅列。自己PR的な誇張・主観表現は禁止）",
  "skills": "活かせる経験・スキル（箇条書き2〜4点を改行区切り、各行先頭に'・'は付けない：コード側で付与）",
  "selfPr": "自己PR（300〜500文字。具体的エピソードを交え、事務職として貢献できる点を明示。ここで強みや意欲を述べる）"
}`;

  const raw = (await callLLM({ systemPrompt, userPrompt, maxTokens: 4096, json: true })).trim();
  return parseJsonLoose(raw);
}

// ════════════════════════════════════════════
// API: 書類生成（AI → Python → ファイル）
// ════════════════════════════════════════════

// 'HH:MM' を15分刻みに丸めて返す（最寄りの15分へ）
function snapTo15Min(hhmm) {
  if (!hhmm) return '';
  const m = String(hhmm).match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return hhmm;
  let h = parseInt(m[1], 10);
  let min = Math.round(parseInt(m[2], 10) / 15) * 15;
  if (min === 60) { min = 0; h += 1; }
  if (h >= 24) { h = 23; min = 45; }
  return `${String(h).padStart(2,'0')}:${String(min).padStart(2,'0')}`;
}

function normalizeInterviewDates(arr = []) {
  return (arr || []).map(iv => ({
    ...iv,
    timeStart: snapTo15Min(iv.timeStart || ''),
    timeEnd:   snapTo15Min(iv.timeEnd   || ''),
  }));
}

app.post('/api/generate', async (req, res) => {
  try {
    const formData = req.body;
    if (!formData.candidateName) {
      return res.status(400).json({ error: '氏名は必須です。' });
    }
    if (!hasLLMKey()) {
      return res.status(500).json({ error: 'GEMINI_API_KEY または ANTHROPIC_API_KEY を .env に設定してください。' });
    }

    // 面接希望日時を15分刻みに正規化（保険）
    if (formData.interviewDates) {
      formData.interviewDates = normalizeInterviewDates(formData.interviewDates);
    }

    // Step1: AIで職務要約・スキル・自己PRを生成
    const narrative = await generateNarrative(formData);

    // Step2: 全データをPythonへ渡してファイル生成
    const inputForPython = { ...formData, ...narrative };
    const files = await runPython(inputForPython);

    // Step3: 推薦先ごとの推薦文テキストを生成（メール用）
    let recommendationTexts = [];
    try {
      recommendationTexts = await generateRecommendationTexts(formData, narrative);
    } catch (e) {
      console.error('[recommendation-texts] error:', e);
    }

    return res.json({ ok: true, files, narrative, recommendationTexts });
  } catch (err) {
    console.error('[generate] error:', err);
    return res.status(500).json({ error: 'エラーが発生しました: ' + err.message });
  }
});

// ════════════════════════════════════════════
// 推薦文テキスト生成（各推薦先のメール原稿）
// ════════════════════════════════════════════

function prefectureOf(address) {
  if (!address) return '';
  const m = String(address).match(/^(.+?[都道府県])/);
  return m ? m[1] : '';
}

// 日付文字列を "M月D日（曜）" 形式に整形
function formatDateJP(dateStr) {
  if (!dateStr) return '';
  // YYYY-MM-DD / YYYY/MM/DD どちらも受け付け
  const m = String(dateStr).match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (!m) return dateStr;
  const y = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10);
  const d = parseInt(m[3], 10);
  const dow = ['日','月','火','水','木','金','土'][new Date(y, mo-1, d).getDay()];
  return `${mo}月${d}日（${dow}）`;
}

function formatInterviewDatesJP(interviewDates = []) {
  // "5月19日（月）　終日可能" / "5月19日（月）　10:00〜12:00" 形式
  const out = [];
  for (const iv of interviewDates) {
    const datePretty = formatDateJP(iv.date || '');
    if (!datePretty) continue;
    const ts = (iv.timeStart || '').trim();
    const te = (iv.timeEnd || '').trim();
    const note = (iv.note || '').trim();
    let timeStr;
    // note が時間情報ならそれ優先
    if (note && /終日|午前|午後|時/.test(note)) {
      timeStr = note;
    } else if (!ts && !te) {
      timeStr = '終日可能';
    } else if ((ts === '09:00' || ts === '9:00') && /^1[789]:\d{2}|^19:45$/.test(te)) {
      timeStr = '終日可能';
    } else if (ts && te) {
      timeStr = `${ts}〜${te}`;
    } else if (ts && !te) {
      timeStr = `${ts}〜`;
    } else if (!ts && te) {
      timeStr = `〜${te}`;
    } else {
      timeStr = '';
    }
    out.push(`${datePretty}　${timeStr}`);
  }
  return out.join('\n');
}

function buildMiraeruText(f) {
  const name = f.candidateName || '';
  const kana = f.candidateNameKana || '';
  const bd = (f.birthYear ? `${f.birthYear}年${f.birthMonth || ''}月${f.birthDay || ''}日生` : '');
  const pref = prefectureOf(f.address);
  const pref_desired = prefectureOf(f.desiredArea) || pref;
  const relocation = (String(f.relocationPlan) === 'true' || f.relocationPlan === true) ? 'あり' : '無し';
  const hope = (f.jobPreference || '').trim();
  const interview = formatInterviewDatesJP(f.interviewDates || []);
  const nameWithWideSp = name.replace(/\s+/g, '　');
  const kanaWithWideSp = kana.replace(/\s+/g, '　');
  return `【ミラエール提案・候補者情報】

■氏名（姓名の間に全角スペース）：${nameWithWideSp}
■ふりがな（姓名の間に全角スペース）：${kanaWithWideSp}
■生年月日：${bd}
■性別：${f.gender || ''}
■メールアドレス：${f.email || ''}
■最終学歴（〇〇卒の形で記載）：${f.lastEducation || ''}
■現住所（都道府県のみ）：${pref}
■希望勤務地（都道府県のみ）：${pref_desired}
■転居予定：${relocation}${hope ? `\n備考：${hope}` : ''}

【面接希望日】
${interview || '（未記入）'}`;
}

async function generateSuisenbun(f, narrative) {
  // 推薦文（200〜350字程度）を AI で生成
  const systemPrompt = `あなたは日本の人材紹介会社のCA（キャリアアドバイザー）です。
推薦先企業へメール本文として添える「推薦文」を作成してください。
・誇張・事実に無い記述は書かない
・3〜5文、約200〜350字
・定型として「推薦者は」から始める
・最後は「貴社にて○○を発揮し貢献できる」「ぜひ一度面接を宜しくお願い致します。」で締める
・テンプレ例：
『推薦者は、とても愛嬌がよく素直な方です。面談の際にはコミュニケーション能力の高さを感じました。未経験ではありますが、スキルを向上させたいと意欲を示していました。推薦者のコミュニケーション能力、成長意欲は貴社にて存分に発揮し、貢献することができると感じました。ぜひ一度面接を宜しくお願い致します。』`;

  const userPrompt = `以下の求職者情報から推薦文を作成してください。純粋な日本語の本文のみを返し、JSONや囲み記号は不要。

氏名: ${f.candidateName || ''}（${f.age || '?'}歳）
強み: ${f.strengths || '（不明）'}
転職理由: ${f.resignReason || '（不明）'}
事務希望理由: ${f.whyOffice || '（不明）'}
スキル: Excel=${(f.excelSkills || []).join('・')} / Word=${(f.wordSkills || []).join('・')} / 電話=${f.phoneExperience || ''} / メール=${f.emailExperience || ''}
職務要約: ${narrative?.workSummary || ''}
備考メモ: ${f.notes || ''}
`;

  try {
    const raw = (await callLLM({ systemPrompt, userPrompt, maxTokens: 800 })).trim();
    return raw.replace(/```[^\n]*\n?/g, '').replace(/```/g, '').trim();
  } catch (e) {
    // フォールバック（AIエラー時の固定文）
    return '推薦者は、とても素直で意欲的な方です。面談の際にはコミュニケーション能力の高さを感じました。前職での経験を活かし、貴社にて貢献することができると感じました。ぜひ一度面接を宜しくお願い致します。';
  }
}

function buildManpowerText(f, suisenbun) {
  const name = f.candidateName || '';
  const kana = f.candidateNameKana || '';
  const age = f.age || '';
  const tel = f.mobile || '';
  const mail = f.email || '';
  const edu = f.lastEducation || '';
  const pref = prefectureOf(f.address);
  const pref_desired = prefectureOf(f.desiredArea) || pref;
  const available = f.availableFrom || '';
  const interview = formatInterviewDatesJP(f.interviewDates || []);
  return `マンパワーグループ株式会社　様

お世話になっております。
1名貴社に推薦させていただきたい方がおりますので、
基本情報を下記にて明記致しますのでご確認いただけますと幸いです。

名前：${name}
ふりがな：${kana}
年齢： ${age}歳
電話番号：${tel}
MAIL：${mail}
最終学歴：${edu}
現住所都道府県名：${pref}
希望勤務地：${pref_desired}
就業可能時期：${available}
ステータス：中途
選考方法希望：WEB


日程：
${interview || '（未記入）'}


【推薦文】
${suisenbun}

以上でございます。

お忙しいところ恐縮ですが
ご検討の程、よろしくお願い致します。`;
}

async function generateRecommendationTexts(formData, narrative) {
  const recommendations = formData.recommendations || [];
  const results = [];
  // ミラエール
  if (recommendations.includes('ミラエール')) {
    results.push({ label: 'ミラエール 提案メール', text: buildMiraeruText(formData) });
  }
  // マンパワー
  if (recommendations.includes('マンパワー')) {
    const sb = await generateSuisenbun(formData, narrative);
    results.push({ label: 'マンパワーグループ 提案メール', text: buildManpowerText(formData, sb) });
  }
  return results;
}

// ════════════════════════════════════════════
// API: ファイルダウンロード
// ════════════════════════════════════════════

app.get('/api/download/:sessionId/:filename', (req, res) => {
  const { sessionId, filename } = req.params;
  // Prevent directory traversal
  if (sessionId.includes('..') || filename.includes('..')) {
    return res.status(400).send('Invalid path');
  }
  const filePath = path.join(OUTPUT_DIR, sessionId, filename);
  if (!fs.existsSync(filePath)) {
    return res.status(404).send('File not found');
  }
  const ext  = path.extname(filename).toLowerCase();
  const mime = ext === '.docx'
    ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
  res.setHeader('Content-Type', mime);
  res.sendFile(filePath);
});

// ════════════════════════════════════════════
// Google OAuth
// ════════════════════════════════════════════

app.get('/api/auth/google', (req, res) => {
  if (!process.env.GOOGLE_CLIENT_ID) {
    return res.status(500).send('Google OAuth が .env に未設定です。');
  }
  const url = buildOAuth().generateAuthUrl({
    access_type: 'offline', prompt: 'consent', scope: GOOGLE_SCOPES,
  });
  res.redirect(url);
});

app.get('/api/auth/google/callback', async (req, res) => {
  try {
    const { code } = req.query;
    const oauth2Client = buildOAuth();
    const { tokens } = await oauth2Client.getToken(code);
    req.session.googleTokens = tokens;
    oauth2Client.setCredentials(tokens);
    try {
      const oauth2 = google.oauth2({ version: 'v2', auth: oauth2Client });
      req.session.googleEmail = (await oauth2.userinfo.get()).data.email || null;
    } catch (_) {}
    res.send(`<!doctype html><meta charset="utf-8"><title>連携完了</title>
<style>body{font-family:sans-serif;padding:40px;text-align:center}</style>
<h2>✅ Google連携が完了しました</h2><p>このタブを閉じてください。</p>
<script>if(window.opener){window.opener.postMessage('google-auth-success','*');window.close();}</script>`);
  } catch (err) {
    res.status(500).send('認証失敗: ' + err.message);
  }
});

app.get('/api/auth/status', (req, res) => {
  res.json({
    authenticated: !!(req.session.googleTokens?.access_token),
    email: req.session.googleEmail || null,
  });
});

app.post('/api/auth/logout', (req, res) => {
  req.session.destroy(() => res.json({ ok: true }));
});

// ════════════════════════════════════════════
// Google Drive 保存（生成済みファイルをアップロード）
// ════════════════════════════════════════════

app.post('/api/drive/save', async (req, res) => {
  if (!req.session.googleTokens) {
    return res.status(401).json({ error: 'Google認証が完了していません。' });
  }
  const { files } = req.body;
  if (!files?.length) return res.status(400).json({ error: 'ファイル情報がありません。' });

  const oauth2Client = buildOAuth();
  oauth2Client.setCredentials(req.session.googleTokens);
  const drive = google.drive({ version: 'v3', auth: oauth2Client });

  const parent = process.env.GOOGLE_DRIVE_FOLDER_ID || null;
  const results = [];

  for (const f of files) {
    if (f.error) continue;
    const filePath = path.join(OUTPUT_DIR, f.sessionId, f.filename);
    if (!fs.existsSync(filePath)) continue;

    const ext    = path.extname(f.filename).toLowerCase();
    const source = ext === '.docx'
      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const targetMime = 'application/vnd.google-apps.' + (ext === '.docx' ? 'document' : 'spreadsheet');

    const meta = { name: f.filename, mimeType: targetMime };
    if (parent) meta.parents = [parent];

    const created = await drive.files.create({
      requestBody: meta,
      media: { mimeType: source, body: fs.createReadStream(filePath) },
      fields: 'id,name,webViewLink',
    });
    results.push({ label: f.label, name: created.data.name, link: created.data.webViewLink });
  }

  res.json({ ok: true, files: results });
});

// ════════════════════════════════════════════
// 起動
// ════════════════════════════════════════════

app.listen(PORT, () => {
  console.log(`\n🚀 CA書類自動生成ツール v2 起動`);
  console.log(`   http://localhost:${PORT}\n`);
  if (USE_GEMINI) console.log('🤖 LLM: Gemini (無料枠)');
  else if (process.env.ANTHROPIC_API_KEY) console.log('🤖 LLM: Claude');
  else console.warn('⚠️  GEMINI_API_KEY / ANTHROPIC_API_KEY 未設定');
  if (!process.env.GOOGLE_CLIENT_ID)  console.warn('⚠️  GOOGLE_CLIENT_ID 未設定（Drive連携不可）');
});
