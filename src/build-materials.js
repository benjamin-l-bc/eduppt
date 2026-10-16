// 生成课堂打印材料（PDF）：教具卡片 + 贴纸，以及教案 PDF
// 用法：npm run materials
const path = require("path");
const fs = require("fs");
const { chromium } = require("playwright");
const { marked } = require("marked");

const ROOT = path.join(__dirname, "..");
const img = (f) => "file://" + path.join(ROOT, "assets", f);
const e = (n) => img(`e_${n}.png`);

const BASE_CSS = `
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Microsoft YaHei", "PingFang SC", "WenQuanYi Zen Hei", sans-serif; color: #2D2A6E; }
  .page { width: 297mm; height: 210mm; page-break-after: always; position: relative; overflow: hidden; }
  .page:last-child { page-break-after: auto; }
  .cut { position: absolute; bottom: 6mm; right: 10mm; font-size: 10pt; color: #9AA0C8; }
`;

// ---------- 1. 魔法公式大卡（贴黑板，每张一页） ----------
const formula = [
  ["谁", "一只小兔子", "#F25C66", "rabbit"],
  ["什么样", "白白的、毛茸茸的", "#FF9F1C", "sparkles"],
  ["在哪里", "在绿绿的草地上", "#1FAF6E", "herb"],
  ["做什么", "吃胡萝卜", "#3A86FF", "carrot"],
];
const formulaPages = formula
  .map(([k, ex, c, em], i) => `
  <section class="page" style="background:${c};color:#fff;display:flex;flex-direction:column;align-items:center;justify-content:center">
    <img src="${e(em)}" style="width:42mm;height:42mm;margin-bottom:6mm">
    <div style="font-size:150pt;font-weight:bold;line-height:1.1">${k}</div>
    <div style="font-size:36pt;margin-top:8mm;opacity:.95">例：${ex}</div>
    <div style="position:absolute;top:10mm;left:14mm;font-size:22pt;font-weight:bold;opacity:.85">魔法公式 ${i + 1} / 4</div>
  </section>`)
  .join("");

// ---------- 2. 英语指令卡（每页 2 张，沿虚线剪开） ----------
const cmds = [
  ["standing", "Stand up!", "站起来"], ["chair", "Sit down!", "坐下"],
  ["kangaroo", "Jump!", "跳一跳"], ["clap", "Clap your hands!", "拍拍手"],
  ["nose", "Touch your nose!", "摸鼻子"], ["turn", "Turn around!", "转个圈"],
];
let cmdPages = "";
for (let i = 0; i < cmds.length; i += 2) {
  cmdPages += `<section class="page" style="display:flex">${cmds
    .slice(i, i + 2)
    .map(([em, en, zh], j) => `
      <div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;${j === 0 ? "border-right:2px dashed #B8C4E8" : ""}">
        <img src="${e(em)}" style="width:70mm;height:70mm">
        <div style="font-family:Arial,sans-serif;font-weight:bold;font-size:${en.length > 12 ? 40 : 54}pt;color:#1FAF6E;margin-top:8mm;text-align:center">${en}</div>
        <div style="font-size:28pt;color:#5B5F8F;margin-top:3mm">${zh}</div>
      </div>`)
    .join("")}<div class="cut">✂ 沿虚线剪开</div></section>`;
}

// ---------- 3. ⭕ / ❌ 举牌（对折粘在一次性筷子上） ----------
const paddle = `
  <section class="page" style="display:flex">
    <div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;border-right:2px dashed #B8C4E8">
      <div style="width:120mm;height:120mm;border-radius:50%;border:22mm solid #1FAF6E"></div>
      <div style="font-size:48pt;font-weight:bold;color:#1FAF6E;margin-top:8mm">能 / 对</div>
    </div>
    <div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center">
      <div style="position:relative;width:130mm;height:130mm">
        <div style="position:absolute;left:54mm;top:0;width:22mm;height:130mm;background:#F25C66;border-radius:11mm;transform:rotate(45deg)"></div>
        <div style="position:absolute;left:54mm;top:0;width:22mm;height:130mm;background:#F25C66;border-radius:11mm;transform:rotate(-45deg)"></div>
      </div>
      <div style="font-size:48pt;font-weight:bold;color:#F25C66;margin-top:0mm">不能 / 错</div>
    </div>
    <div class="cut">（可选）沿虚线对折，背对背粘在一根筷子上，做成双面举牌</div>
  </section>`;

// ---------- 4. 机器人头套脸（贴在纸箱正面） ----------
const helmet = `
  <section class="page" style="display:flex;align-items:center;justify-content:center;gap:16mm;background:#EEF3FF">
    <div style="width:170mm;height:150mm;background:#2D2A6E;border-radius:40mm;display:flex;flex-direction:column;align-items:center;justify-content:center;border:5mm solid #fff;box-shadow:0 0 0 2mm #2D2A6E">
      <div style="display:flex;gap:46mm">
        <div style="width:30mm;height:38mm;border-radius:50%;background:#5CF2E5"></div>
        <div style="width:30mm;height:38mm;border-radius:50%;background:#5CF2E5"></div>
      </div>
      <div style="width:70mm;height:30mm;border-bottom:7mm solid #5CF2E5;border-radius:0 0 40mm 40mm;margin-top:6mm"></div>
    </div>
    <div style="display:flex;flex-direction:column;align-items:center">
      <div style="background:#FFC83D;border:3mm solid #2D2A6E;border-radius:10mm;padding:6mm 12mm;font-family:Arial;font-weight:bold;font-size:60pt">AI</div>
      <div style="font-size:30pt;font-weight:bold;margin-top:8mm">ROBOT 小智</div>
    </div>
    <div class="cut">机器人头套：剪下左边的“脸”贴在纸箱正面，右边的“AI”牌贴在胸前或箱子侧面（纸箱底部挖洞套头，正面挖两个小孔看路）</div>
  </section>`;

// ---------- 5. 「AI 小训练师」贴纸（每页 12 枚，班级约 45 人打印 4 页） ----------
const sticker = `
  <div style="width:62mm;height:62mm;border-radius:50%;background:#1E1B4B;color:#fff;display:flex;flex-direction:column;align-items:center;justify-content:center;border:2.5mm solid #FFC83D">
    <img src="${img("robot_happy.png")}" style="height:24mm">
    <div style="font-size:14pt;font-weight:bold;color:#FFC83D;margin-top:1.5mm">AI 小训练师</div>
    <div style="font-size:8.5pt;margin-top:1mm;letter-spacing:.5pt">语文 · 数学 · 英语</div>
    <div style="font-size:8pt;margin-top:.8mm;opacity:.8">建平实验小学 一（8）班</div>
  </div>`;
const stickerPage = `
  <section class="page" style="display:grid;grid-template-columns:repeat(4,62mm);grid-template-rows:repeat(3,62mm);gap:4mm 8mm;justify-content:center;align-content:center">
    ${sticker.repeat(12)}
    <div class="cut">建议用 A4 不干胶纸打印，全班约 45 人打印 4 页</div>
  </section>`;

const COVER = `
  <section class="page" style="display:flex;align-items:center;gap:14mm;padding:0 22mm;background:#1E1B4B;color:#fff">
    <img src="${img("robot_happy.png")}" style="height:150mm">
    <div>
      <div style="font-size:52pt;font-weight:bold;color:#FFC83D">《AI伴我成长》</div>
      <div style="font-size:30pt;margin-top:4mm">课堂教具打印包</div>
      <ol style="font-size:16pt;line-height:1.9;margin-top:10mm;padding-left:8mm;color:#E4DAFF">
        <li>魔法公式大卡 ×4（第 10–11 页课件，贴黑板）</li>
        <li>英语指令卡 ×6（第 19 页课件，Robot, please!）</li>
        <li>⭕ / ❌ 举牌（可选，代替手势）</li>
        <li>机器人头套贴纸（第 10、19 页“机器人老师”）</li>
        <li>「AI 小训练师」奖励贴纸 ×12/页</li>
      </ol>
      <div style="font-size:12pt;margin-top:8mm;opacity:.7">全部 A4 横向，彩色打印效果最好；黑白打印也能用。</div>
    </div>
  </section>`;

const packHTML = `<!doctype html><html><head><meta charset="utf-8"><style>${BASE_CSS}@page{size:A4 landscape;margin:0}</style></head>
<body>${COVER}${formulaPages}${cmdPages}${paddle}${helmet}${stickerPage}</body></html>`;

// ---------- 教案 PDF（由 教案.md 渲染） ----------
const md = fs.readFileSync(path.join(ROOT, "教案.md"), "utf8");
const planHTML = `<!doctype html><html><head><meta charset="utf-8"><style>
  @page { size: A4; margin: 16mm 16mm 18mm; }
  body { font-family: "Microsoft YaHei","PingFang SC","WenQuanYi Zen Hei",sans-serif; color:#222; font-size:10.5pt; line-height:1.7; }
  h1 { color:#2D2A6E; font-size:22pt; margin-bottom:4mm; }
  h2 { color:#fff; background:#2D2A6E; font-size:14pt; padding:1.5mm 4mm; border-radius:2mm; margin:7mm 0 3mm; }
  h3 { color:#2D2A6E; font-size:12pt; margin:5mm 0 2mm; }
  table { border-collapse:collapse; width:100%; margin:2mm 0 4mm; font-size:9.5pt; }
  th, td { border:1px solid #C9D2F5; padding:1.5mm 2.5mm; vertical-align:top; text-align:left; }
  th { background:#EEF3FF; }
  td:first-child, th:first-child { white-space: nowrap; }
  blockquote { background:#FFF6DC; border-radius:2mm; padding:2mm 4mm; margin:2mm 0; }
  code { background:#EEF3FF; padding:0 1.5mm; border-radius:1mm; }
  ul, ol { padding-left:6mm; }
  h2, h3 { page-break-after: avoid; } tr { page-break-inside: avoid; }
</style></head><body>${marked.parse(md)}</body></html>`;

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  for (const [html, out, landscape] of [
    [packHTML, "课堂教具打印包.pdf", true],
    [planHTML, "教案.pdf", false],
  ]) {
    const tmp = path.join(ROOT, "assets", "_tmp.html");
    fs.writeFileSync(tmp, html);
    await page.goto("file://" + tmp, { waitUntil: "load" });
    await page.pdf({ path: path.join(ROOT, out), format: "A4", landscape, printBackground: true, preferCSSPageSize: true });
    fs.unlinkSync(tmp);
    console.log("written", out);
  }
  await browser.close();
})();
