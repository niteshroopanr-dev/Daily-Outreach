const pptxgen = require("pptxgenjs");

// Brand colours (Rule 3), no hash prefix
const BLACK = "000000";
const TEAL = "01A296";
const AMBER_BRIGHT = "F8C806";
const AMBER_DEEP = "F6A102";
const GOLD = "E3A712";
const WHITE = "FFFFFF";
const OFF_WHITE = "E6E5DE";

const HEADER_SERIF = "Cambria";
const BODY_SANS = "Calibri";

const DATE_STAMP = "31 Jul 2026";
const COMPANY = "Orbitkey";

function newDeck() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  return pres;
}

function addChrome(slide, eyebrow, rightLabel) {
  // page background
  slide.background = { color: WHITE };
  // left accent stripe
  slide.addShape("rect", { x: 0, y: 0, w: 0.1, h: 7.5, fill: { color: AMBER_BRIGHT }, line: { type: "none" } });
  // header band
  slide.addShape("rect", { x: 0, y: 0, w: 13.333, h: 1.0, fill: { color: BLACK }, line: { type: "none" } });
  slide.addText(eyebrow, {
    x: 0.3, y: 0.32, w: 7.5, h: 0.4, fontFace: BODY_SANS, fontSize: 12, bold: true,
    color: OFF_WHITE, align: "left", margin: 0,
  });
  slide.addText(rightLabel, {
    x: 9.5, y: 0.32, w: 3.5, h: 0.4, fontFace: BODY_SANS, fontSize: 12, bold: true,
    color: TEAL, align: "right", margin: 0,
  });
  // footer line
  slide.addText("Prepared by Nitesh Roopa CA, Managing Partner and Founder, ProfitPulse, Profit-Pulse.com.au", {
    x: 0.3, y: 7.05, w: 9.0, h: 0.3, fontFace: BODY_SANS, fontSize: 9, color: BLACK, align: "left", margin: 0,
  });
  slide.addText(DATE_STAMP, {
    x: 10.0, y: 7.05, w: 3.0, h: 0.3, fontFace: BODY_SANS, fontSize: 9, color: BLACK, align: "right", margin: 0,
  });
}

const pres = newDeck();

// ─────────────────────────────────────────────────────────────────────────
// SLIDE 1: COMMERCIAL INTELLIGENCE BRIEF
// ─────────────────────────────────────────────────────────────────────────
const s1 = pres.addSlide();
addChrome(s1, "COMMERCIAL INTELLIGENCE BRIEF", "PROFITPULSE");

s1.addText(COMPANY, {
  x: 0.35, y: 1.12, w: 9.0, h: 0.75, fontFace: HEADER_SERIF, fontSize: 42, bold: true,
  color: BLACK, align: "left", margin: 0,
});
s1.addText("Design led everyday carry accessories, Richmond, Melbourne VIC", {
  x: 0.35, y: 1.85, w: 10.5, h: 0.35, fontFace: BODY_SANS, fontSize: 13, italic: true,
  color: BLACK, align: "left", margin: 0,
});

// Stat cards row: 5 cards, 2.3" wide, 0.15" gap, top y 2.4, height 1.6
const cardY = 2.4, cardW = 2.3, cardH = 1.6, cardGap = 0.15, cardStartX = 0.35;
const cards = [
  { num: "$24.6M", lbl: "Annual revenue", src: "SmartCompany Smart50 2025" },
  { num: "20%", lbl: "Average 3 year growth", src: "SmartCompany Smart50 2025" },
  { num: "44", lbl: "People on the team", src: "SmartCompany Smart50 2025" },
  { num: "83.7", lbl: "B Corp impact score", src: "B Lab Global directory" },
  { num: "$8M+", lbl: "Raised via Kickstarter", src: "SmartCompany, 11 campaigns" },
];
cards.forEach((c, i) => {
  const x = cardStartX + i * (cardW + cardGap);
  s1.addShape("rect", { x, y: cardY, w: cardW, h: cardH, fill: { color: BLACK }, line: { type: "none" } });
  s1.addShape("rect", { x, y: cardY, w: cardW, h: 0.06, fill: { color: TEAL }, line: { type: "none" } });
  s1.addText(c.num, {
    x: x + 0.15, y: cardY + 0.16, w: cardW - 0.3, h: 0.5, fontFace: HEADER_SERIF, fontSize: 27, bold: true,
    color: AMBER_BRIGHT, align: "left", margin: 0,
  });
  s1.addText(c.lbl, {
    x: x + 0.15, y: cardY + 0.68, w: cardW - 0.3, h: 0.5, fontFace: BODY_SANS, fontSize: 11,
    color: OFF_WHITE, align: "left", margin: 0, valign: "top",
  });
  s1.addText(c.src, {
    x: x + 0.15, y: cardY + cardH - 0.32, w: cardW - 0.3, h: 0.26, fontFace: BODY_SANS, fontSize: 7.5,
    color: TEAL, align: "left", margin: 0,
  });
});

// Key Commercial Signals
s1.addText("KEY COMMERCIAL SIGNALS", {
  x: 0.35, y: 4.22, w: 6.0, h: 0.3, fontFace: BODY_SANS, fontSize: 12, bold: true,
  color: TEAL, align: "left", margin: 0,
});
const signals = [
  "B Corp Certified, impact score 83.7 vs global median 50.9 (B Lab Global).",
  "2025 Australian Export Award winner, Creative Industries (exportawards.gov.au).",
  "New Richmond HQ built beside the Kensington design studio (company sources).",
  "Latest Kickstarter campaign raised over $393,500, no external investors (SmartCompany).",
  "South Korea flagged as a significant emerging export market (exportawards.gov.au).",
  "Smart50 2025 ranking, independent confirmation of continued fast growth (SmartCompany).",
];
s1.addText(
  signals.map((t, i) => ({ text: t, options: { bullet: { code: "2022" }, breakLine: i < signals.length - 1, color: BLACK, fontSize: 11.5, fontFace: BODY_SANS, paraSpaceAfter: 6 } })),
  { x: 0.35, y: 4.56, w: 12.4, h: 2.3, margin: 0, valign: "top" }
);

// ─────────────────────────────────────────────────────────────────────────
// SLIDE 2: THE OPPORTUNITY
// ─────────────────────────────────────────────────────────────────────────
const s2 = pres.addSlide();
addChrome(s2, "THE OPPORTUNITY", "PROFITPULSE");

s2.addText("Orbitkey. Three commercial observations from ProfitPulse.", {
  x: 0.35, y: 1.14, w: 12.4, h: 0.4, fontFace: HEADER_SERIF, fontSize: 19, bold: true,
  color: BLACK, align: "left", margin: 0,
});

const colY = 1.7, colH = 4.6, colW = 4.1, colGap = 0.15, colStartX = 0.35;
const observations = [
  {
    index: "01", header: "Stock now funds Orbitkey's growth", fill: TEAL, textColor: BLACK,
    para: "Twenty percent average growth with zero external capital means inventory must be bought before wholesale partners or Kickstarter backers pay Orbitkey back. With around 3,000 retail partners plus direct ecommerce, cash timing matters as much as sales. A working capital review shows what releasing cash could fund next.",
  },
  {
    index: "02", header: "The channel mix is getting harder to read", fill: BLACK, textColor: OFF_WHITE,
    para: "Orbitkey sells through direct ecommerce, around 3,000 wholesale partners, and periodic Kickstarter launches, three channels with different margins and cost to serve. As the wholesale network keeps growing internationally, not every channel is contributing equally once freight and 3PL terms are allocated properly.",
  },
  {
    index: "03", header: "Recognition is outrunning reporting", fill: GOLD, textColor: BLACK,
    para: "A B Corp certification, a national Export Award, and a Smart50 ranking in one year put Orbitkey in front of partners and future capital sources expecting board grade numbers on request. Staying self funded is a real strength, but it still benefits from the same monthly rhythm investor backed rivals must run.",
  },
];
observations.forEach((o, i) => {
  const x = colStartX + i * (colW + colGap);
  s2.addShape("rect", { x, y: colY, w: colW, h: colH, fill: { color: o.fill }, line: { type: "none" } });
  s2.addText(o.index, {
    x: x + 0.25, y: colY + 0.2, w: colW - 0.5, h: 0.55, fontFace: HEADER_SERIF, fontSize: 30, bold: true,
    color: o.textColor, align: "left", margin: 0,
  });
  s2.addText(o.header, {
    x: x + 0.25, y: colY + 0.85, w: colW - 0.5, h: 0.75, fontFace: HEADER_SERIF, fontSize: 15, bold: true,
    color: o.textColor, align: "left", margin: 0, valign: "top",
  });
  s2.addText(o.para, {
    x: x + 0.25, y: colY + 1.65, w: colW - 0.5, h: colH - 1.85, fontFace: BODY_SANS, fontSize: 10.5,
    color: o.textColor, align: "left", margin: 0, valign: "top", lineSpacingMultiple: 1.15,
  });
});

s2.addText(
  "These observations are offered in good faith. Orbitkey has built something impressive without taking a dollar of outside capital. The question is simply whether the financial architecture keeps pace with the ambition.",
  { x: 0.35, y: colY + colH + 0.12, w: 12.4, h: 0.5, fontFace: BODY_SANS, fontSize: 10.5, italic: true, color: BLACK, align: "left", margin: 0 }
);

// ─────────────────────────────────────────────────────────────────────────
// SLIDE 3: THE RECOMMENDATION
// ─────────────────────────────────────────────────────────────────────────
const s3 = pres.addSlide();
addChrome(s3, "THE RECOMMENDATION", "PROFITPULSE");

const leftX = 0.35, leftW = 7.65;
const rightX = 8.35, rightW = 4.6;

s3.addText("Working Capital Unlock", {
  x: leftX, y: 1.15, w: leftW, h: 0.5, fontFace: HEADER_SERIF, fontSize: 26, bold: true,
  color: BLACK, align: "left", margin: 0,
});
s3.addText("$6,000 one off", {
  x: leftX, y: 1.65, w: leftW, h: 0.4, fontFace: BODY_SANS, fontSize: 18, bold: true,
  color: AMBER_DEEP, align: "left", margin: 0,
});
s3.addText(
  "A four week project mapping cash trapped in stock, debtors, and supplier terms across five global warehouses, with a prioritised, quantified action list to release it.",
  { x: leftX, y: 2.12, w: leftW, h: 0.55, fontFace: BODY_SANS, fontSize: 12, color: BLACK, align: "left", margin: 0, valign: "top" }
);

// Step one block
s3.addShape("rect", { x: leftX, y: 2.85, w: leftW, h: 1.35, fill: { color: OFF_WHITE }, line: { type: "none" } });
s3.addText("Step one, answer a few quick questions", {
  x: leftX + 0.2, y: 2.98, w: leftW - 0.4, h: 0.35, fontFace: BODY_SANS, fontSize: 13, bold: true,
  color: BLACK, align: "left", margin: 0,
});
s3.addText("See the solutions matched to your size and industry.", {
  x: leftX + 0.2, y: 3.32, w: leftW - 0.4, h: 0.32, fontFace: BODY_SANS, fontSize: 11.5,
  color: BLACK, align: "left", margin: 0,
});
s3.addText("profit-pulse.com.au/services/find-your-fit", {
  x: leftX + 0.2, y: 3.66, w: leftW - 0.4, h: 0.4, fontFace: BODY_SANS, fontSize: 13, bold: true,
  color: TEAL, align: "left", margin: 0,
  hyperlink: { url: "https://profit-pulse.com.au/services/find-your-fit/" },
});

// Direct CTA
s3.addShape("roundRect", {
  x: leftX, y: 4.4, w: leftW, h: 0.55, rectRadius: 0.1,
  fill: { color: AMBER_DEEP }, line: { type: "none" },
});
s3.addText("Purchase the suggested product now to get started", {
  x: leftX, y: 4.4, w: leftW, h: 0.55, fontFace: BODY_SANS, fontSize: 13, bold: true,
  color: BLACK, align: "center", valign: "middle", margin: 0, wrap: false,
  hyperlink: { url: "https://buy.stripe.com/00w4gA7OUfGM5TadMH3ks0z" },
});

// Booking
s3.addText("Prefer a conversation first?", {
  x: leftX, y: 5.2, w: leftW, h: 0.32, fontFace: BODY_SANS, fontSize: 12, color: BLACK, align: "left", margin: 0,
});
s3.addText("Book a complimentary discovery call", {
  x: leftX, y: 5.52, w: leftW, h: 0.35, fontFace: BODY_SANS, fontSize: 12.5, bold: true,
  color: TEAL, align: "left", margin: 0,
  hyperlink: { url: "https://bookings.cloud.microsoft/book/ProfitPulse1@profit-pulse.com.au/?ismsaljsauthenabled=true" },
});

// Right column: credibility panel
s3.addShape("rect", { x: rightX, y: 1.15, w: rightW, h: 5.6, fill: { color: BLACK }, line: { type: "none" } });
s3.addText("NITESH ROOPA", {
  x: rightX + 0.25, y: 1.32, w: rightW - 0.5, h: 0.42, fontFace: HEADER_SERIF, fontSize: 17, bold: true,
  color: AMBER_BRIGHT, align: "left", margin: 0,
});
s3.addText("CA, Managing Partner, ProfitPulse", {
  x: rightX + 0.25, y: 1.75, w: rightW - 0.5, h: 0.35, fontFace: BODY_SANS, fontSize: 12,
  color: WHITE, align: "left", margin: 0,
});

const credLines = [
  "16 years across 4 countries",
  "52 deals executed and managed",
  "Largest single deal USD 1.3 billion, Cahora Bassa",
  "Total GRBT project value over AUD 10 billion",
];
s3.addText(
  credLines.map((t, i) => ({ text: t, options: { bullet: { code: "2022" }, breakLine: i < credLines.length - 1, color: OFF_WHITE, fontSize: 10.5, fontFace: BODY_SANS, paraSpaceAfter: 8 } })),
  { x: rightX + 0.25, y: 2.25, w: rightW - 0.5, h: 1.7, margin: 0, valign: "top" }
);

s3.addShape("line", { x: rightX + 0.25, y: 4.05, w: rightW - 0.5, h: 0, line: { color: TEAL, width: 1 } });

const contactLines = [
  "Profit-Pulse.com.au",
  "Nitesh@Profit-Pulse.com.au",
  "+61 411 876 267",
  "linkedin.com/in/nitesh-roopa-77594163",
];
s3.addText(
  contactLines.map((t, i) => ({ text: t, options: { breakLine: i < contactLines.length - 1, color: OFF_WHITE, fontSize: 10.5, fontFace: BODY_SANS, paraSpaceAfter: 6 } })),
  { x: rightX + 0.25, y: 4.25, w: rightW - 0.5, h: 1.4, margin: 0, valign: "top" }
);

pres.writeFile({ fileName: "/home/user/Daily-Outreach/Out-reach efforts/Brief_Orbitkey_31Jul2026.pptx" })
  .then(() => console.log("PPTX written"))
  .catch((e) => { console.error(e); process.exit(1); });
