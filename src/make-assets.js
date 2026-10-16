// 生成课件用到的图片素材：emoji 图标（用 Chromium + Noto Color Emoji 渲染）和手绘 SVG（机器人小智、六指手）
const path = require("path");
const fs = require("fs");
const sharp = require("sharp");
const { chromium } = require("playwright");

const OUT = path.join(__dirname, "..", "assets");
fs.mkdirSync(OUT, { recursive: true });

const EMOJI = {
  speak: "🗣️", palette: "🎨", books: "📚", speaker: "🔊", broom: "🧹", camera: "📷", map: "🗺️",
  flower: "🌻", dino: "🦖", rice: "🍚", story: "📖", ruler: "📏", rabbit: "🐰", carrot: "🥕",
  octopus: "🐙", hat: "🎩", smile: "😀", numbers: "🔢", abc: "🔤", bird: "🐦", apple: "🍎",
  cookie: "🍪", fire: "🔥", dog: "🐶", hotdog: "🌭", hot: "🥵", chili: "🌶️", standing: "🧍",
  chair: "🪑", kangaroo: "🦘", clap: "👏", nose: "👃", turn: "🔄", brain: "🧠", family: "🤝",
  shush: "🤫", think: "🤔", globe: "🌏", key: "🗝️", gift: "🎁", star: "⭐", trophy: "🏆",
  phone: "📱", herb: "🌿", question: "❓", sparkles: "✨", wave: "👋", rocket: "🚀", house: "🏠",
  pencil: "✏️", magnifier: "🔍", eyes: "👀", lol: "🤣", teacher: "🧑‍🏫", leaf: "🌱", sun: "☀️",
};

async function renderEmoji() {
  const browser = await chromium.launch({ executablePath: "/opt/pw-browsers/chromium-1194/chrome-linux/chrome" }).catch(() => chromium.launch());
  const page = await browser.newPage({ deviceScaleFactor: 1 });
  for (const [name, ch] of Object.entries(EMOJI)) {
    await page.setContent(
      `<html><body style="margin:0;background:transparent"><div id="e" style="display:inline-block;font-family:'Noto Color Emoji';font-size:220px;line-height:1;padding:16px">${ch}</div></body></html>`
    );
    const buf = await page.locator("#e").screenshot({ omitBackground: true });
    // 裁掉透明边，统一成 300x300
    await sharp(buf).trim().resize(300, 300, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toFile(path.join(OUT, `e_${name}.png`));
  }
  await browser.close();
}

function robotSVG(mood) {
  const eyes =
    mood === "oops"
      ? `<path d="M118 150 l22 14 M118 178 l22 -14" stroke="#5CF2E5" stroke-width="9" stroke-linecap="round" fill="none"/>
         <path d="M262 150 l-22 14 M262 178 l-22 -14" stroke="#5CF2E5" stroke-width="9" stroke-linecap="round" fill="none"/>
         <path d="M150 215 q20 -14 40 0 q20 14 40 0" stroke="#5CF2E5" stroke-width="8" stroke-linecap="round" fill="none"/>
         <path d="M318 92 q14 22 0 34 q-14 -12 0 -34z" fill="#7CC8FF" stroke="#2D2A6E" stroke-width="4"/>`
      : `<ellipse cx="135" cy="162" rx="17" ry="22" fill="#5CF2E5"/><ellipse cx="245" cy="162" rx="17" ry="22" fill="#5CF2E5"/>
         <circle cx="141" cy="154" r="6" fill="#fff"/><circle cx="251" cy="154" r="6" fill="#fff"/>
         <path d="M150 205 q40 34 80 0" stroke="#5CF2E5" stroke-width="9" stroke-linecap="round" fill="none"/>`;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="380" height="520" viewBox="0 0 380 520">
  <line x1="190" y1="60" x2="190" y2="24" stroke="#2D2A6E" stroke-width="8"/>
  <circle cx="190" cy="22" r="17" fill="#FFC83D" stroke="#2D2A6E" stroke-width="6"/>
  <rect x="22" y="128" width="30" height="70" rx="14" fill="#FF8FA3" stroke="#2D2A6E" stroke-width="6"/>
  <rect x="328" y="128" width="30" height="70" rx="14" fill="#FF8FA3" stroke="#2D2A6E" stroke-width="6"/>
  <rect x="44" y="58" width="292" height="230" rx="70" fill="#F2F6FF" stroke="#2D2A6E" stroke-width="8"/>
  <rect x="82" y="100" width="216" height="148" rx="48" fill="#2D2A6E"/>
  ${eyes}
  <circle cx="98" cy="262" r="14" fill="#FFB3C1"/><circle cx="282" cy="262" r="14" fill="#FFB3C1"/>
  <rect x="160" y="286" width="60" height="22" fill="#B8C4E8" stroke="#2D2A6E" stroke-width="6"/>
  <rect x="34" y="330" width="44" height="120" rx="22" fill="#F2F6FF" stroke="#2D2A6E" stroke-width="7" transform="rotate(${mood === "oops" ? -150 : 18} 56 340)"/>
  <rect x="302" y="330" width="44" height="120" rx="22" fill="#F2F6FF" stroke="#2D2A6E" stroke-width="7" transform="rotate(-18 324 340)"/>
  <rect x="76" y="304" width="228" height="176" rx="56" fill="#F2F6FF" stroke="#2D2A6E" stroke-width="8"/>
  <rect x="130" y="342" width="120" height="78" rx="24" fill="#FFC83D" stroke="#2D2A6E" stroke-width="6"/>
  <text x="190" y="397" font-family="Arial" font-weight="bold" font-size="46" text-anchor="middle" fill="#2D2A6E">AI</text>
  <rect x="110" y="476" width="56" height="34" rx="14" fill="#2D2A6E"/><rect x="214" y="476" width="56" height="34" rx="14" fill="#2D2A6E"/>
</svg>`;
}

// AI 画错的手：一眼能数出 6 根手指（5 根竖指 + 1 根拇指）
function handSVG() {
  const fingers = [70, 112, 154, 196, 238]
    .map((x, i) => {
      const top = [96, 60, 44, 58, 86][i];
      return `<rect x="${x}" y="${top}" width="40" height="${300 - top}" rx="20" fill="#FFD3B0" stroke="#B5651D" stroke-width="7"/>`;
    })
    .join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="420" height="460" viewBox="0 0 420 460">
  ${fingers}
  <rect x="252" y="230" width="42" height="160" rx="21" fill="#FFD3B0" stroke="#B5651D" stroke-width="7" transform="rotate(40 273 390)"/>
  <rect x="64" y="210" width="220" height="210" rx="80" fill="#FFD3B0" stroke="#B5651D" stroke-width="7"/>
  <rect x="78" y="200" width="196" height="60" rx="20" fill="#FFD3B0"/>
</svg>`;
}

(async () => {
  await renderEmoji();
  await sharp(Buffer.from(robotSVG("happy"))).resize({ height: 800 }).png().toFile(path.join(OUT, "robot_happy.png"));
  await sharp(Buffer.from(robotSVG("oops"))).resize({ height: 800 }).png().toFile(path.join(OUT, "robot_oops.png"));
  await sharp(Buffer.from(handSVG())).resize({ height: 800 }).png().toFile(path.join(OUT, "hand6.png"));
  console.log("assets done:", fs.readdirSync(OUT).length);
})();
