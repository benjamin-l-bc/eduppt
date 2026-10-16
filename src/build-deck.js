// 《AI伴我成长》一年级课件生成脚本
// 用法：npm run build   （先 npm run assets 生成图片素材）
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("./apply_theme.js");

const ROOT = path.join(__dirname, "..");
const A = (f) => path.join(ROOT, "assets", f);
const OUT = path.join(ROOT, "AI伴我成长.pptx");

const THEME = {
  name: "AI Grow Up",
  headFontFace: "Microsoft YaHei",
  bodyFontFace: "Microsoft YaHei",
  colors: {
    dk1: "2D2A6E", // 正文：靛蓝
    lt1: "FFFFFF",
    dk2: "1E1B4B", // 深色页背景
    lt2: "EEF3FF", // 浅色页背景
    accent1: "FFC83D", // 黄：高亮
    accent2: "F25C66", // 珊瑚红：语文
    accent3: "3A86FF", // 蓝：数学
    accent4: "1FAF6E", // 绿：英语
    accent5: "8B5CF6", // 紫：AI / 神秘
    accent6: "FF9F1C", // 橙
    hlink: "3A86FF",
    folHlink: "8B5CF6",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.title = "AI伴我成长";
pres.subject = "一年级 AI 启蒙课";
pres.author = "建平实验小学 一（8）班";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const W = 13.333;

// ---------- 版式 ----------
pres.defineSlideMaster({
  title: "COVER",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.9, w: 7.6, h: 1.6, fontSize: 72, bold: true, color: C.accent1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 3.6, w: 7.6, h: 0.9, fontSize: 28, color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "SECTION",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 6.0, y: 2.35, w: 6.8, h: 1.3, fontSize: 60, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 6.0, y: 3.75, w: 6.8, h: 1.2, fontSize: 30, color: C.accent1, align: "left", valign: "top", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.3, y: 6.95, w: 0.6, h: 0.35, fontSize: 12, color: "B8C4E8", align: "right" },
});
pres.defineSlideMaster({
  title: "DARK",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 1.0, fontSize: 44, bold: true, color: C.background1, align: "left", valign: "middle", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 12.3, y: 6.95, w: 0.6, h: 0.35, fontSize: 12, color: "B8C4E8", align: "right" },
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 12.1, h: 1.0, fontSize: 40, bold: true, color: C.text1, align: "left", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: "AI伴我成长 · 一（8）班", options: { x: 0.6, y: 6.95, w: 4, h: 0.35, fontSize: 12, color: "8A90B8", margin: 0 } } },
  ],
  slideNumber: { x: 12.3, y: 6.95, w: 0.6, h: 0.35, fontSize: 12, color: "8A90B8", align: "right" },
});

// ---------- 小工具 ----------
const shadow = () => ({ type: "outer", color: "2D2A6E", opacity: 0.14, blur: 8, offset: 3, angle: 90 });
let objN = 0;
const nm = (s) => `${s}-${++objN}`;

function card(slide, x, y, w, h, fill = C.background1, opts = {}) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: opts.r ?? 0.25, fill: { color: fill }, line: { type: "none" }, shadow: opts.noShadow ? undefined : shadow(), objectName: nm("card") });
}
function circle(slide, x, y, d, fill, transparency = 0) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill, transparency }, line: { type: "none" }, objectName: nm("circle") });
}
function emoji(slide, name, x, y, size) {
  slide.addImage({ path: A(`e_${name}.png`), x, y, w: size, h: size, altText: name, objectName: nm(`emoji-${name}`) });
}
// 圆圈里放一个 emoji —— 全课统一的视觉母题
function badge(slide, name, x, y, d, fill, transparency = 0) {
  circle(slide, x, y, d, fill, transparency);
  const s = d * 0.62;
  emoji(slide, name, x + (d - s) / 2, y + (d - s) / 2, s);
}
function robot(slide, mood, x, y, h) {
  slide.addImage({ path: A(`robot_${mood}.png`), x, y, w: h * (380 / 520), h, altText: "机器人小智", objectName: nm("robot") });
}
function text(slide, t, x, y, w, h, o = {}) {
  slide.addText(t, { x, y, w, h, isTextBox: true, margin: 0, fontSize: 24, color: C.text1, valign: "middle", objectName: nm("text"), ...o });
}
// 小智的对话气泡（tail 指向机器人所在的一侧）
function bubble(slide, t, x, y, w, h, o = {}) {
  const fill = o.fill ?? C.background1;
  const tail = o.tail ?? "left";
  const ts = 0.36;
  const tri = {
    left: { x: x - ts + 0.06, y: y + h / 2 - ts / 2, rotate: 270 },
    right: { x: x + w - 0.06, y: y + h / 2 - ts / 2, rotate: 90 },
    down: { x: x + w * 0.2, y: y + h - 0.06, rotate: 180 },
  }[tail];
  slide.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: tri.x, y: tri.y, w: ts, h: ts, rotate: tri.rotate, fill: { color: fill }, line: { type: "none" }, objectName: nm("bubble-tail") });
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.3, fill: { color: fill }, line: { type: "none" }, objectName: nm("bubble") });
  text(slide, t, x + 0.25, y, w - 0.5, h, { fontSize: o.fontSize ?? 24, bold: o.bold ?? true, color: o.color ?? C.text1, align: o.align ?? "left" });
}
function stamp(slide, ok, label, x, y, w, h) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h / 2, fill: { color: ok ? C.accent4 : C.accent2 }, line: { type: "none" }, objectName: nm("stamp") });
  text(slide, label, x, y, w, h, { fontSize: 22, bold: true, color: C.background1, align: "center" });
}

// ======================================================================
// 第一部分：认识 AI
// ======================================================================
pres.addSection({ title: "开场：认识 AI" });

// 1 封面
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "开场：认识 AI" });
  s.addText("AI伴我成长", { placeholder: "title" });
  s.addText("建平实验小学 一（8）班", { placeholder: "body" });
  [["books", C.accent2, "语文"], ["numbers", C.accent3, "数学"], ["abc", C.accent4, "英语"]].forEach(([e, c, l], i) => {
    const x = 0.8 + i * 2.15;
    slide_pill(s, e, c, l, x, 4.85);
  });
  circle(s, 7.5, 0.9, 5.4, C.accent5, 55);
  robot(s, "happy", 8.25, 1.2, 5.3);
  emoji(s, "sparkles", 7.4, 0.7, 1.0);
  emoji(s, "star", 12.0, 5.6, 0.7);
  s.addNotes(`【0:00–1:00 开场】
（上课前：戴好纸箱机器人头套，或把头套藏在讲台下，准备第 10 页用。）

老师：小朋友们好！今天老师带来了一节特别的课，名字叫——（指屏幕）《AI伴我成长》！
看，屏幕上这个圆圆的小家伙是谁呀？先别说，等一下我们来猜一猜。
今天我们还要找到三件“法宝”（指三个小标签）：语文、数学、英语。找到法宝的小朋友，下课可以拿到一张「AI 小训练师」贴纸哦！`);
}
function slide_pill(s, e, c, l, x, y) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.95, h: 0.75, rectRadius: 0.375, fill: { color: c }, line: { type: "none" }, objectName: nm("pill") });
  emoji(s, e, x + 0.15, y + 0.12, 0.5);
  text(s, l, x + 0.75, y, 1.1, 0.75, { fontSize: 24, bold: true, color: C.background1 });
}

// 2 猜猜我是谁
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "开场：认识 AI" });
  s.addText("猜猜我是谁？", { placeholder: "title" });
  const clues = [
    ["speak", C.accent2, "我没有嘴巴，却会说话"],
    ["palette", C.accent3, "我没有小手，却会画画"],
    ["books", C.accent4, "我不用上学，却认识好多好多字"],
  ];
  clues.forEach(([e, c, t], i) => {
    const y = 1.65 + i * 1.6;
    card(s, 0.6, y, 7.9, 1.3);
    badge(s, e, 0.8, y + 0.15, 1.0, c, 75);
    text(s, t, 2.05, y, 6.3, 1.3, { fontSize: 30, bold: true });
  });
  card(s, 9.0, 1.65, 3.7, 4.5, C.accent5);
  emoji(s, "question", 9.85, 2.0, 2.0);
  text(s, "举手猜一猜！", 9.0, 4.35, 3.7, 0.8, { fontSize: 30, bold: true, color: C.background1, align: "center" });
  text(s, "它就藏在下一页", 9.0, 5.1, 3.7, 0.6, { fontSize: 20, color: "E4DAFF", align: "center" });
  s.addNotes(`【1:00–3:00 猜谜】
老师用神秘的语气，一条一条读：
“我没有嘴巴，却会说话；我没有小手，却会画画；我不用上学，却认识好多好多字。我是谁？”

请 2–3 位小朋友猜。可能的答案：机器人、手机、电脑、小爱同学……都夸一夸：“很接近了！”
如果有小朋友答出 AI / 人工智能，马上给他一个大拇指：“哇，你是小专家！”`);
}

// 3 揭晓：我是 AI
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "开场：认识 AI" });
  s.addText("揭晓答案！", { placeholder: "title" });
  text(s, "AI", 0.6, 1.5, 5.2, 2.6, { fontSize: 180, bold: true, color: C.accent1, fontFace: "Arial" });
  text(s, "人工智能", 0.6, 4.1, 5.2, 0.9, { fontSize: 44, bold: true, color: C.background1 });
  text(s, "一个会学习的“电脑大脑”", 0.6, 5.0, 6.0, 0.6, { fontSize: 24, color: "C9D2F5" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.85, w: 5.6, h: 0.85, rectRadius: 0.42, fill: { color: C.accent5 }, line: { type: "none" }, objectName: nm("pill") });
  text(s, "跟我读：A～ I～ （做机器人动作）", 0.85, 5.85, 5.2, 0.85, { fontSize: 22, bold: true, color: C.background1 });
  robot(s, "happy", 9.2, 1.6, 5.2);
  bubble(s, "大家好！\n我叫“小智”！", 6.3, 2.0, 3.3, 1.5, { tail: "right", fontSize: 22 });
  s.addNotes(`【3:00–4:00 揭晓】
老师：答案就是——AI！（声音放大）
AI 的中文名字叫“人工智能”，就是人类做出来的、会学习的“电脑大脑”。

★ 互动：全班一起做“机器人动作”（两只手臂弯成直角，一顿一顿地动），用机器人的声音读：“A～ I～”，读 2 遍。
老师用机器人声音：“大家好，我是小智！今天我要和一（8）班的小朋友交朋友！”`);
}

// 4 AI 就在我们身边
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "开场：认识 AI" });
  s.addText("AI 就在我们身边", { placeholder: "title" });
  const items = [
    ["speaker", C.accent5, "智能音箱", "“小爱同学，\n讲个故事！”"],
    ["broom", C.accent6, "扫地机器人", "自己把地\n扫干净"],
    ["camera", C.accent4, "拍照识花", "拍一拍，就知道\n花叫什么名字"],
    ["map", C.accent3, "地图导航", "“前方路口，\n请右转！”"],
  ];
  items.forEach(([e, c, t, d], i) => {
    const x = 0.6 + i * 3.1;
    card(s, x, 1.6, 2.8, 3.7);
    badge(s, e, x + 0.6, 1.85, 1.6, c, 78);
    text(s, t, x, 3.6, 2.8, 0.6, { fontSize: 28, bold: true, align: "center" });
    text(s, d, x + 0.15, 4.2, 2.5, 0.95, { fontSize: 18, color: "5B5F8F", align: "center", valign: "top" });
  });
  robot(s, "happy", 0.7, 5.45, 1.4);
  bubble(s, "你家里有哪一个？有的小朋友请举手！", 2.1, 5.65, 7.2, 0.95, { tail: "left" });
  s.addNotes(`【4:00–6:00 AI 在身边】
老师：其实 AI 早就住在我们身边了！
逐个指：智能音箱——你跟它说话它会回答；扫地机器人——自己认路扫地；拍照识花——拍一下就知道这是什么花；地图导航——告诉爸爸妈妈往哪儿开。

★ 互动：“家里有智能音箱的举手！有扫地机器人的举手！”
请一位小朋友说说：“你平时让它帮你做什么？”
（可选彩蛋：如果带了手机，现场对智能音箱 / AI 助手说一句“给一年级小朋友讲个 10 秒钟的笑话”。）`);
}

// 5 游戏：AI 能不能？（提问）
const canQs = [
  ["dino", C.accent4, "AI 能画一只\n穿西装的恐龙吗？", true],
  ["herb", C.accent2, "AI 能帮你把\n青菜吃掉吗？", false],
  ["story", C.accent3, "AI 能给你\n讲故事吗？", true],
  ["ruler", C.accent6, "AI 能帮你\n长高吗？", false],
];
function canGrid(s, reveal) {
  canQs.forEach(([e, c, t, ok], i) => {
    const x = 5.05 + (i % 2) * 3.95;
    const y = 1.6 + Math.floor(i / 2) * 2.65;
    card(s, x, y, 3.65, 2.35);
    badge(s, e, x + 0.2, y + 0.3, 0.95, c, 75);
    text(s, `${i + 1}`, x + 0.2, y + 1.55, 0.95, 0.55, { fontSize: 26, bold: true, color: "8A90B8", align: "center" });
    text(s, t, x + 1.3, y + 0.2, 2.3, 1.5, { fontSize: 20, bold: true });
    if (reveal) stamp(s, ok, ok ? "⭕ 能！" : "❌ 不能！", x + 1.3, y + 1.6, 1.9, 0.55);
  });
}
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "开场：认识 AI" });
  s.addText("游戏：AI 能不能？", { placeholder: "title" });
  card(s, 0.6, 1.6, 4.1, 5.0, C.text1);
  text(s, "听题，用身体回答！", 0.9, 1.8, 3.6, 0.7, { fontSize: 24, bold: true, color: C.accent1 });
  text(s, "⭕", 0.9, 2.65, 1.2, 1.2, { fontSize: 66, color: C.accent4, align: "center", fontFace: "Arial" });
  text(s, "能\n双手举过头顶\n比个大圆圈", 2.2, 2.55, 2.4, 1.5, { fontSize: 20, color: C.background1 });
  text(s, "❌", 0.9, 4.55, 1.2, 1.2, { fontSize: 66, color: C.accent2, align: "center", fontFace: "Arial" });
  text(s, "不能\n双手在胸前\n打个大叉叉", 2.2, 4.45, 2.4, 1.5, { fontSize: 20, color: C.background1 });
  canGrid(s, false);
  s.addNotes(`【6:00–8:30 游戏：AI 能不能？】
规则：老师读题，小朋友用身体回答——“能”就双手举过头顶比大圆圈 ⭕；“不能”就双手在胸前打叉 ❌。先练一次！

老师一题一题读，读的时候要夸张：
1. “AI 能画一只——穿西装、打领带的——恐龙吗？”（孩子会笑）
2. “AI 能帮你把不爱吃的青菜……吃掉吗？”（会有孩子大喊“能！”——正好制造反差）
3. “AI 能给你讲故事吗？”
4. “AI 能帮你……长高吗？”（老师可以踮脚、伸长脖子做“长高”动作）

每题先不说答案，数一数举 ⭕ 和 ❌ 的人，下一页统一揭晓。`);
}
// 6 揭晓
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "开场：认识 AI" });
  s.addText("答案揭晓！", { placeholder: "title" });
  card(s, 0.6, 1.6, 4.1, 5.0, C.text1);
  robot(s, "happy", 1.75, 1.8, 2.5);
  text(s, "画画、讲故事\n我很在行！", 0.8, 4.4, 3.7, 0.95, { fontSize: 22, bold: true, color: C.accent1, align: "center" });
  text(s, "吃饭、长高……\n要靠你自己哦！", 0.8, 5.45, 3.7, 0.95, { fontSize: 22, bold: true, color: C.background1, align: "center" });
  canGrid(s, true);
  s.addNotes(`【8:30–10:00 揭晓】
逐题揭晓：
1. 能！AI 真的可以画出穿西装的恐龙。（可选：课前用 AI 画图工具生成一张“穿西装打领带的恐龙”，此时展示，效果最好。提示词见教案。）
2. 不能！青菜要你自己吃，AI 没有肚子。（“谁今天会把青菜吃光？”）
3. 能！AI 会讲故事，还能编新故事。
4. 不能！长高要靠好好吃饭、好好睡觉、多运动。

小结：AI 很厉害，可它不是万能的。`);
}

// 7 转折：三件法宝
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "开场：认识 AI" });
  s.addText("可是……AI 也会闹笑话！", { placeholder: "title" });
  robot(s, "oops", 0.8, 1.8, 4.7);
  text(s, "想让 AI 又听话、又不出错，\n你需要三件法宝：", 4.9, 1.75, 7.8, 1.3, { fontSize: 28, bold: true, color: C.background1 });
  [["books", C.accent2, "语文"], ["numbers", C.accent3, "数学"], ["abc", C.accent4, "英语"]].forEach(([e, c, l], i) => {
    const x = 4.9 + i * 2.7;
    badge(s, e, x, 3.35, 2.1, c);
    text(s, l, x, 5.6, 2.1, 0.7, { fontSize: 32, bold: true, color: C.background1, align: "center" });
  });
  s.addNotes(`【10:00–11:00 转折】
老师（换成垂头丧气的机器人声音）：“呜呜……小朋友，其实我有时候会闹笑话……”
老师（恢复正常）：AI 很聪明，但是它要靠我们“指挥”。想让 AI 又听话、又不出错，我们需要三件法宝——
（一个一个指）语文！数学！英语！
我们一件一件去找吧！`);
}

// ======================================================================
// 法宝一：语文
// ======================================================================
function sectionSlide(sec, title, sub, e, c, notes) {
  const s = pres.addSlide({ masterName: "SECTION", sectionTitle: sec });
  circle(s, 0.9, 1.25, 4.6, c);
  emoji(s, e, 1.95, 2.3, 2.5);
  s.addText(title, { placeholder: "title" });
  s.addText(sub, { placeholder: "body" });
  s.addNotes(notes);
  return s;
}
pres.addSection({ title: "法宝一：语文" });
sectionSlide("法宝一：语文", "法宝一：语文", "把话说清楚，AI 才听得懂", "books", C.accent2,
  `【11:00 法宝一】
老师：第一件法宝——语文！语文不只是认字写字，还要会“把话说清楚”。我们来看看，话没说清楚，会发生什么好笑的事……`);

// 9 给我画一只小动物
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "法宝一：语文" });
  s.addText("小明对 AI 说：“给我画一只小动物！”", { placeholder: "title" });
  card(s, 0.6, 1.6, 5.4, 4.0);
  text(s, "小明心里想的", 0.6, 1.75, 5.4, 0.6, { fontSize: 24, bold: true, color: "5B5F8F", align: "center" });
  s.addShape(pres.shapes.CLOUD, { x: 1.1, y: 2.4, w: 4.4, h: 3.0, fill: { color: "E3F6EA" }, line: { color: C.accent4, width: 2 }, objectName: nm("cloud") });
  emoji(s, "rabbit", 1.95, 2.95, 1.7);
  emoji(s, "carrot", 3.55, 3.35, 1.2);
  s.addShape(pres.shapes.RIGHT_ARROW, { x: 6.2, y: 3.15, w: 0.9, h: 0.9, fill: { color: C.accent1 }, line: { type: "none" }, objectName: nm("arrow") });
  card(s, 7.3, 1.6, 5.4, 4.0);
  text(s, "AI 画出来的", 7.3, 1.75, 5.4, 0.6, { fontSize: 24, bold: true, color: "5B5F8F", align: "center" });
  emoji(s, "octopus", 8.6, 2.45, 2.8);
  emoji(s, "lol", 11.5, 4.5, 0.9);
  robot(s, "oops", 0.7, 5.75, 1.15);
  bubble(s, "咦？为什么 AI 画错了？", 1.95, 5.9, 5.6, 0.85, { tail: "left" });
  s.addNotes(`【11:00–13:00 好笑的错误】
老师：小明想要一只小兔子，他对 AI 说：“给我画一只小动物！”
结果 AI 画出来——（停顿，指右边）——一只章鱼！（夸张地做章鱼八只脚乱舞的动作）

★ 提问：“章鱼是不是小动物？” ——是！“那 AI 错了吗？” ——AI 没错，是小明没说清楚！
“那小明应该怎么说？”请 2 位小朋友试着帮小明重新说一说。
（引导出：要说清楚是“兔子”，什么颜色，在干什么。）

过渡：说清楚真的很难吗？我们来玩个游戏试试！`);
}

// 10 游戏：指挥机器人老师
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "法宝一：语文" });
  s.addText("游戏：指挥机器人老师", { placeholder: "title" });
  card(s, 0.6, 1.6, 4.1, 5.0, C.accent5);
  robot(s, "happy", 1.45, 1.95, 3.6);
  emoji(s, "hat", 2.1, 1.6, 0.95);
  text(s, "机器人老师 已上线！", 0.6, 5.7, 4.1, 0.7, { fontSize: 24, bold: true, color: C.background1, align: "center" });
  const rules = [
    ["1", "机器人老师只听“字面意思”", "你怎么说，它就怎么做"],
    ["2", "一次只说一步", "说一句，做一步"],
    ["3", "任务：让老师戴好帽子", "再在黑板上画一张笑脸 😀"],
  ];
  rules.forEach(([n, t, d], i) => {
    const y = 1.6 + i * 1.72;
    card(s, 5.0, y, 7.7, 1.45);
    circle(s, 5.25, y + 0.3, 0.85, C.accent2);
    text(s, n, 5.25, y + 0.3, 0.85, 0.85, { fontSize: 30, bold: true, color: C.background1, align: "center", fontFace: "Arial" });
    text(s, t, 6.35, y + 0.15, 6.2, 0.65, { fontSize: 26, bold: true });
    text(s, d, 6.35, y + 0.78, 6.2, 0.5, { fontSize: 20, color: "5B5F8F" });
  });
  s.addNotes(`【13:00–17:00 ★核心互动：指挥机器人老师】
道具：纸箱头套、一顶帽子（或一本书当帽子）、白板笔。
老师戴上头套，用机器人声音：“滴——机器人老师已上线！我只听字面意思，请下达指令！”

任务一：戴帽子（请举手的小朋友一句一句说）
· 孩子说“把帽子戴上” → 机器人：“滴！戴在哪里？”然后把帽子戴到膝盖上 / 脚上 / 套在手上。
· 孩子说“戴在头上” → 把帽子倒过来扣在头上，或者帽檐朝后歪戴。
· 孩子说“帽子口朝下，戴在头顶上” → 终于戴好了！全班鼓掌。

任务二：画笑脸（请 3–4 位小朋友轮流指挥）
· “画一个圆” → 在黑板角落画一个米粒大的圆（或巨大到出黑板的圆）。
· “画眼睛” → 画 5 只眼睛，或画在圆圈外面。
· “在圆里面上面画两个小圆当眼睛” → 照做，夸奖：“说得真清楚！”
· “画嘴巴” → 画一张哭脸的嘴巴；直到有人说“画一个弯弯的、像香蕉一样往上笑的嘴巴”。

★ 关键：一定要“很认真地”做错，越认真孩子越笑。
收尾（摘下头套）：刚才机器人为什么老做错？——因为我们的话没说清楚！AI 也是这样。`);
}

// 11 说清楚的魔法公式
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "法宝一：语文" });
  s.addText("说清楚的魔法公式", { placeholder: "title" });
  const blocks = [
    ["谁", "一只小兔子", C.accent2],
    ["什么样", "白白的\n毛茸茸的", C.accent6],
    ["在哪里", "在绿绿的\n草地上", C.accent4],
    ["做什么", "吃胡萝卜", C.accent3],
  ];
  blocks.forEach(([k, v, c], i) => {
    const x = 0.6 + i * 3.15;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.6, w: 2.65, h: 2.2, rectRadius: 0.25, fill: { color: c }, line: { type: "none" }, shadow: shadow(), objectName: nm("block") });
    text(s, k, x, 1.75, 2.65, 0.8, { fontSize: 34, bold: true, color: C.background1, align: "center" });
    text(s, v, x + 0.1, 2.55, 2.45, 1.1, { fontSize: 22, color: C.background1, align: "center" });
    if (i < 3) text(s, "+", x + 2.65, 2.25, 0.5, 0.8, { fontSize: 40, bold: true, color: C.text1, align: "center", fontFace: "Arial" });
  });
  card(s, 0.6, 4.15, 5.7, 2.5, "E3F6EA");
  emoji(s, "herb", 0.9, 5.2, 1.1);
  emoji(s, "rabbit", 1.9, 4.35, 1.9);
  emoji(s, "carrot", 3.75, 4.95, 1.2);
  emoji(s, "sparkles", 5.15, 4.3, 0.85);
  robot(s, "happy", 6.7, 4.15, 2.5);
  bubble(s, "这次画对啦！\n你想让 AI 画什么？", 8.75, 4.55, 4.0, 1.6, { tail: "left", fontSize: 22 });
  s.addNotes(`【17:00–19:30 魔法公式】
老师：说清楚有一个魔法公式：谁 + 什么样 + 在哪里 + 做什么。
带读：“一只——白白的、毛茸茸的——小兔子，在绿绿的草地上——吃胡萝卜。”这次 AI 画对啦！

★ 互动：“你想让 AI 画什么？用魔法公式说一说！”请 2–3 位小朋友说，其他同学检查：“他说了‘谁’吗？说了‘在哪里’吗？”
（可选：教室能上网的话，老师把孩子说的话现场输入 AI 画图工具，几秒就出图，全班会非常兴奋。只由老师操作。）

小结：语文学得好，话说得清楚，AI 才能听懂你！——恭喜大家找到第一件法宝！`);
}

// ======================================================================
// 法宝二：数学
// ======================================================================
pres.addSection({ title: "法宝二：数学" });
sectionSlide("法宝二：数学", "法宝二：数学", "会数会算，帮 AI 找错误", "numbers", C.accent3,
  `【19:30 法宝二】
老师：第二件法宝——数学！你们知道吗？AI 有时候……连手指都数不清！不信？看下一页！`);

// 13 数一数：AI 画的手
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "法宝二：数学" });
  s.addText("AI 画了一只手，数一数！", { placeholder: "title" });
  card(s, 0.6, 1.6, 5.6, 5.0);
  s.addImage({ path: A("hand6.png"), x: 1.55, y: 1.85, w: 4.5 * (420 / 460), h: 4.5, altText: "AI 画的手，有 6 根手指", objectName: nm("hand") });
  badge(s, "eyes", 6.7, 1.75, 1.2, C.accent3, 75);
  text(s, "这只手有几根手指？", 8.1, 1.75, 4.6, 1.2, { fontSize: 32, bold: true });
  badge(s, "wave", 6.7, 3.35, 1.2, C.accent6, 75);
  text(s, "再看看你自己的手，\n有几根？", 8.1, 3.35, 4.6, 1.2, { fontSize: 28, bold: true });
  robot(s, "oops", 6.8, 4.85, 1.75);
  bubble(s, "哎呀，我多画了一根！", 8.3, 5.2, 4.4, 0.95, { tail: "left", fontSize: 22 });
  s.addNotes(`【20:00–21:00 数手指】
老师：这是 AI 画的一只手。我们一起数——（用教鞭一根一根指）1、2、3、4、5、6！
“6 根！”再请大家伸出自己的手数一数：“5 根！”
老师夸张地看自己的手：“我也是 5 根！AI 画错啦！”

这是真的：AI 画画时经常把手指画多或画少。（可选：课前用 AI 画图工具生成一张真实的“六指手”图片替换。）
小结：会数数，才能发现 AI 的错误。`);
}

// 14/15 AI 小侦探
const detective = [
  { e: ["bird", 5], c: C.accent3, t: "树上有 5 只小鸟，\n飞走了 2 只，\n还剩 7 只。", ok: false, ans: "❌ 错！还剩 3 只" },
  { e: ["apple", 2, "+"], c: C.accent2, t: "1 个苹果\n加 1 个苹果，\n等于 11 个苹果！", ok: false, ans: "❌ 错！是 2 个" },
  { e: ["cookie", 4, "+", 3], c: C.accent4, t: "盘子里有 4 块饼干，\n妈妈又拿来 3 块，\n一共 7 块。", ok: true, ans: "⭕ 对！一共 7 块" },
];
function detectiveSlide(reveal) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "法宝二：数学" });
  s.addText(reveal ? "侦探报告：AI 算对了吗？" : "当 AI 小侦探：它算对了吗？", { placeholder: "title" });
  detective.forEach((d, i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.55, 3.8, 4.15);
    // 用 emoji 摆出题目里的数量，孩子可以直接数
    const [name, n1, plus, n2] = d.e;
    if (name === "apple") {
      emoji(s, "apple", x + 0.75, 1.8, 0.75);
      text(s, "+", x + 1.5, 1.8, 0.8, 0.75, { fontSize: 36, bold: true, align: "center", fontFace: "Arial" });
      emoji(s, "apple", x + 2.3, 1.8, 0.75);
    } else if (plus) {
      const sz = 0.42;
      for (let k = 0; k < n1; k++) emoji(s, name, x + 0.2 + k * 0.44, 1.98, sz);
      text(s, "+", x + 1.95, 1.85, 0.35, 0.65, { fontSize: 30, bold: true, align: "center", fontFace: "Arial" });
      for (let k = 0; k < n2; k++) emoji(s, name, x + 2.3 + k * 0.44, 1.98, sz);
    } else {
      for (let k = 0; k < n1; k++) emoji(s, name, x + 0.3 + k * 0.65, 1.88, 0.6);
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.25, y: 2.8, w: 1.15, h: 0.45, rectRadius: 0.22, fill: { color: d.c }, line: { type: "none" }, objectName: nm("ai-says") });
    text(s, "AI 说", x + 0.25, 2.8, 1.15, 0.45, { fontSize: 16, bold: true, color: C.background1, align: "center" });
    text(s, d.t, x + 0.3, 3.3, 3.25, 1.5, { fontSize: 22, bold: true, valign: "top" });
    if (reveal) stamp(s, d.ok, d.ans, x + 0.3, 4.95, 3.2, 0.55);
    else text(s, "⭕ 对 还是 ❌ 错？", x + 0.3, 4.95, 3.2, 0.55, { fontSize: 20, color: "8A90B8" });
  });
  robot(s, reveal ? "happy" : "oops", 0.7, 5.8, 1.1);
  bubble(s, reveal ? "AI 也会粗心，会数学的你，就是 AI 的小老师！" : "伸出手指，帮 AI 算一算！", 1.9, 5.92, 8.2, 0.82, { tail: "left", fontSize: 22 });
  return s;
}
detectiveSlide(false).addNotes(`【21:00–23:00 ★互动：AI 小侦探】
老师：现在你们都是“AI 小侦探”！AI 做了三道题，请你检查它算对了没有。还是用 ⭕ 和 ❌ 回答。

第 1 题：树上有 5 只小鸟，飞走了 2 只，AI 说还剩 7 只。
——请大家伸出 5 根手指，弯下去 2 根，还剩几根？3 根！AI 错了，它把“飞走”当成了“飞来”。

第 2 题：1 个苹果加 1 个苹果，AI 说等于 11 个苹果！（老师读得一本正经，孩子会大笑）
——两个 1 挨在一起变成了 11，这是真实存在的电脑“笑话”！应该是 2 个。

第 3 题：4 块饼干，又拿来 3 块，AI 说一共 7 块。——这次对啦！所以 AI 也不是总错，要自己检查。`);
detectiveSlide(true).addNotes(`【23:00–24:30 揭晓 + 小结】
逐张揭晓，每揭晓一个就让全班一起喊答案：“3 只！”“2 个！”“7 块，对啦！”

小结：AI 说的话不一定都对。数学学得好，你就能当 AI 的小老师，帮它找错误！——恭喜找到第二件法宝！`);

// ======================================================================
// 法宝三：英语
// ======================================================================
pres.addSection({ title: "法宝三：英语" });
sectionSlide("法宝三：英语", "法宝三：英语", "会说英语，和全世界的 AI 交朋友", "abc", C.accent4,
  `【24:30 法宝三】
老师：第三件法宝——英语！你们知道吗，“AI”这两个字母本身就是英文。很多 AI 最先学会的就是英语。
不过，AI 翻译英语的时候，也闹过大笑话……`);

// 17 hot dog 猜一猜
const hotdogOpts = [
  [["hot", "dog"], "一只热得\n冒汗的小狗"],
  [["hotdog"], "夹着香肠\n的面包"],
  [["chili", "dog"], "一只辣辣的\n小狗"],
];
function hotdogSlide(reveal) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "法宝三：英语" });
  s.addText(reveal ? "答案是②：hot dog 就是热狗面包！" : "猜一猜：hot dog 是什么？", { placeholder: "title" });
  // 算式：hot + dog = ?
  card(s, 0.6, 1.55, 12.1, 1.35);
  emoji(s, "fire", 0.9, 1.7, 1.0);
  text(s, "hot = 热", 1.95, 1.55, 2.3, 1.35, { fontSize: 32, bold: true, color: C.accent2 });
  text(s, "+", 4.25, 1.55, 0.6, 1.35, { fontSize: 40, bold: true, align: "center", fontFace: "Arial" });
  emoji(s, "dog", 4.9, 1.7, 1.0);
  text(s, "dog = 狗", 5.95, 1.55, 2.3, 1.35, { fontSize: 32, bold: true, color: C.accent3 });
  text(s, "→", 8.25, 1.55, 0.7, 1.35, { fontSize: 40, bold: true, align: "center", fontFace: "Arial" });
  text(s, "hot dog = ？", 9.0, 1.55, 3.6, 1.35, { fontSize: 34, bold: true, color: C.accent4 });
  hotdogOpts.forEach(([es, label], i) => {
    const x = 0.6 + i * 4.1;
    const right = i === 1;
    const dim = reveal && !right;
    card(s, x, 3.15, 3.8, 3.55, dim ? "E2E6F3" : C.background1);
    circle(s, x + 0.2, 3.3, 0.75, dim ? "B8C4E8" : C.accent5);
    text(s, `${i + 1}`, x + 0.2, 3.3, 0.75, 0.75, { fontSize: 30, bold: true, color: C.background1, align: "center", fontFace: "Arial" });
    if (es.length === 1) emoji(s, es[0], x + 1.15, 3.4, 1.5);
    else { emoji(s, es[0], x + 0.95, 3.5, 1.15); emoji(s, es[1], x + 2.05, 3.5, 1.15); }
    text(s, label, x + 0.2, 5.05, 3.4, 1.0, { fontSize: 24, bold: true, align: "center", color: dim ? "8A90B8" : C.text1 });
    if (reveal && right) stamp(s, true, "⭕ 就是它！", x + 0.8, 6.05, 2.2, 0.5);
    if (reveal && i === 0) stamp(s, false, "AI 闹的笑话", x + 0.8, 6.05, 2.2, 0.5);
  });
  return s;
}
hotdogSlide(false).addNotes(`【25:00–26:30 ★猜一猜 hot dog】
老师：hot 是“热”，dog 是“狗”。那 hot dog 是什么？
（读选项时加动作：①老师扇风、吐舌头、喘气装热狗狗；②做咬面包的动作；③做被辣到的样子“嘶哈嘶哈”。）
★ 互动：伸出手指投票，觉得是 1 的伸 1 根手指，是 2 的伸 2 根……`);
hotdogSlide(true).addNotes(`【26:30–27:30 揭晓】
答案是 2：hot dog 是夹着香肠的面包，中文叫“热狗”。
可是，如果一个 AI 只会一个字一个字地翻译，它就会翻成“一只很热的狗”——①号这只满头大汗的小狗！（孩子会笑）
小结：学好英语，AI 翻错了你也能一眼发现！`);

// 19 Robot, please!
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "法宝三：英语" });
  s.addText("游戏：Robot, please!", { placeholder: "title" });
  const cmds = [
    ["standing", "Stand up!", "站起来"],
    ["chair", "Sit down!", "坐下"],
    ["kangaroo", "Jump!", "跳一跳"],
    ["clap", "Clap your hands!", "拍拍手"],
    ["nose", "Touch your nose!", "摸鼻子"],
    ["turn", "Turn around!", "转个圈"],
  ];
  cmds.forEach(([e, en, zh], i) => {
    const x = 0.6 + (i % 3) * 2.95;
    const y = 1.55 + Math.floor(i / 3) * 2.6;
    card(s, x, y, 2.7, 2.35);
    emoji(s, e, x + 0.85, y + 0.2, 1.0);
    text(s, en, x + 0.1, y + 1.25, 2.5, 0.55, { fontSize: en.length > 12 ? 19 : 22, bold: true, color: C.accent4, align: "center", fontFace: "Arial" });
    text(s, zh, x + 0.1, y + 1.78, 2.5, 0.45, { fontSize: 18, color: "5B5F8F", align: "center" });
  });
  card(s, 9.55, 1.55, 3.15, 4.95, C.text1);
  robot(s, "happy", 10.15, 1.75, 2.6);
  text(s, "说英语：\n机器人照做！", 9.7, 4.4, 2.85, 0.95, { fontSize: 20, bold: true, color: C.accent1, align: "center" });
  text(s, "说中文：\n“Beep! I don’t understand!”", 9.7, 5.35, 2.85, 1.05, { fontSize: 16, color: C.background1, align: "center" });
  s.addNotes(`【27:30–30:30 ★互动：Robot, please!】
第一轮：老师戴上机器人头套。小朋友举手用英语下指令，老师照做（指令要加 please 更有礼貌）。
· 有人用中文说“跳一下” → 机器人愣住：“Beep! Beep! I don't understand!”（原地卡顿抖动）——孩子会马上改成英语。
· 说 “Jump!” → 老师跳；说 “Turn around!” → 老师转圈，可以故意转 3 圈转晕。
· 搞笑彩蛋：有人说 “Touch your nose!” → 机器人先摸了旁边小朋友的鼻子 / 摸了自己的耳朵，“Oops! Sorry!”

第二轮（反转）：全班当机器人，老师下指令：Stand up! → Clap your hands! → Touch your nose! → Turn around! → Sit down!
最后一个一定是 Sit down，顺利收回课堂秩序。

小结：会英语，就能和更多 AI、更多世界上的小朋友交朋友！——恭喜找到第三件法宝！`);
}

// ======================================================================
// 总结
// ======================================================================
pres.addSection({ title: "总结" });
{
  const s = pres.addSlide({ masterName: "DARK", sectionTitle: "总结" });
  s.addText("我找到了三件法宝！", { placeholder: "title" });
  const tre = [
    ["books", C.accent2, "语文", "说清楚\nAI 听得懂"],
    ["numbers", C.accent3, "数学", "会检查\nAI 错了我知道"],
    ["abc", C.accent4, "英语", "会英语\n世界朋友多"],
  ];
  tre.forEach(([e, c, l, d], i) => {
    const x = 0.6 + i * 3.35;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.6, w: 3.05, h: 5.0, rectRadius: 0.3, fill: { color: c }, line: { type: "none" }, objectName: nm("treasure") });
    circle(s, x + 0.65, 1.9, 1.75, C.background1);
    emoji(s, e, x + 0.97, 2.22, 1.1);
    text(s, l, x, 3.85, 3.05, 0.8, { fontSize: 38, bold: true, color: C.background1, align: "center" });
    text(s, d, x + 0.15, 4.7, 2.75, 1.4, { fontSize: 24, bold: true, color: C.background1, align: "center" });
  });
  robot(s, "happy", 10.6, 2.35, 3.6);
  emoji(s, "trophy", 10.75, 1.45, 0.9);
  emoji(s, "sparkles", 12.0, 1.6, 0.7);
  s.addNotes(`【30:30–31:30 总结三件法宝】
老师：今天我们找到了哪三件法宝？（全班一起喊）语文！数学！英语！
· 语文——把话说清楚，AI 才听得懂。（回忆：章鱼、机器人老师戴帽子）
· 数学——会数会算，AI 错了我知道。（回忆：六根手指、11 个苹果）
· 英语——会说英语，世界朋友多。（回忆：很热的狗）
所以，想要用好 AI，最重要的是——先把自己的本领学扎实！`);
}

// 21 四个约定
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "总结" });
  s.addText("和 AI 做朋友的四个约定", { placeholder: "title" });
  const rules = [
    ["brain", C.accent5, "作业自己做", "AI 是好帮手，不是“替手”"],
    ["family", C.accent6, "大人陪着用", "用 AI 时，爸爸妈妈在身边"],
    ["shush", C.accent2, "秘密不告诉", "姓名、电话、住址不告诉 AI"],
    ["think", C.accent3, "动脑想一想", "AI 说的话，不一定都对"],
  ];
  rules.forEach(([e, c, t, d], i) => {
    const x = 0.6 + (i % 2) * 6.15;
    const y = 1.6 + Math.floor(i / 2) * 2.55;
    card(s, x, y, 5.95, 2.25);
    badge(s, e, x + 0.3, y + 0.42, 1.4, c, 75);
    text(s, t, x + 1.95, y + 0.35, 3.8, 0.8, { fontSize: 30, bold: true });
    text(s, d, x + 1.95, y + 1.15, 3.85, 0.85, { fontSize: 20, color: "5B5F8F", valign: "top" });
  });
  s.addNotes(`【31:30–33:00 四个约定】
老师：AI 是我们的好朋友，和好朋友相处也有约定。我们伸出小拇指，一起来拉钩：
1. 作业自己做——AI 是帮手，不能替我们学习。（“AI 帮你写作业，那长本领的是谁呀？”——是 AI！）
2. 大人陪着用——用手机、用 AI，要有爸爸妈妈在身边。
3. 秘密不告诉——自己的名字、电话、家住在哪里，不告诉 AI，也不告诉陌生人。
4. 动脑想一想——AI 说的话不一定都对，就像刚才的“11 个苹果”。`);
}

// 22 拍手口诀
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "总结" });
  s.addText("一起拍手念口诀", { placeholder: "title" });
  card(s, 0.6, 1.6, 8.6, 5.0);
  const lines = [
    ["AI 好帮手，", "学习靠自己；", C.accent5],
    ["说话要清楚，", "算数要仔细；", C.accent2],
    ["英语学起来，", "朋友遍天下！", C.accent4],
  ];
  lines.forEach(([a, b, c], i) => {
    const y = 1.95 + i * 1.5;
    emoji(s, "clap", 1.0, y + 0.15, 0.85);
    text(s, [
      { text: a, options: { color: c, bold: true } },
      { text: b, options: { color: C.text1, bold: true } },
    ], 2.1, y, 6.9, 1.15, { fontSize: 40 });
  });
  robot(s, "happy", 9.7, 2.1, 3.9);
  emoji(s, "star", 9.6, 1.6, 0.75);
  emoji(s, "star", 12.15, 5.6, 0.6);
  s.addNotes(`【33:00–34:00 拍手口诀】
节奏：每三个字拍一下手（“AI 好 | 帮手，学习 | 靠自己”），老师先念一遍，再全班一起念两遍，最后一遍越念越快！

AI 好帮手，学习靠自己；
说话要清楚，算数要仔细；
英语学起来，朋友遍天下！`);
}

// 23 结束
{
  const s = pres.addSlide({ masterName: "COVER", sectionTitle: "总结" });
  s.addText("AI伴我成长", { placeholder: "title" });
  s.addText("我们一起长大！Bye-bye！", { placeholder: "body" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 4.85, w: 6.4, h: 0.9, rectRadius: 0.45, fill: { color: C.accent1 }, line: { type: "none" }, objectName: nm("pill") });
  emoji(s, "star", 1.0, 4.97, 0.66);
  text(s, "颁发「AI 小训练师」贴纸", 1.8, 4.85, 5.2, 0.9, { fontSize: 26, bold: true, color: C.text1 });
  circle(s, 7.5, 0.9, 5.4, C.accent5, 55);
  robot(s, "happy", 8.25, 1.2, 5.3);
  emoji(s, "wave", 7.5, 1.0, 1.2);
  s.addNotes(`【34:00–35:00 结束】
老师：今天大家都找到了三件法宝，你们都是合格的「AI 小训练师」！
（发贴纸：可以请小组长帮忙，或者下课时站在门口一个一个贴。）
机器人声音：“Bye-bye，一（8）班的小朋友！记得好好学语文、数学、英语，下次我们再见！”
全班挥手：Bye-bye！`);
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
