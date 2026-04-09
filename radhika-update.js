/**
 * KindKart – Sponsor Update for Radhika
 * April 2026 · PptxGenJS
 * 17 slides — designed for a 5-8 minute read
 */
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10" × 5.625"
pres.title = "KindKart Sponsor Update – Radhika – April 2026";
pres.author = "Sri G Gopalakrishnan";

// ─── Brand Palette ─────────────────────────────────────────────
const C = {
  bgDark:   "1A1A1C",
  bgCard:   "28282C",
  bgCard2:  "383840",
  orange:   "E87028",
  green:    "2DB352",
  gold:     "F0C142",
  teal:     "14B8A6",
  white:    "FFFFFF",
  offWhite: "E6E6EA",
  gray:     "8A8A94",
  grayMid:  "AEAEB8",
  blue:     "4A90D9",
  purple:   "9B72CF",
  pink:     "E8507A",
};

// Fresh shadow objects (never reuse — pptxgenjs mutates in-place)
const mkShadow  = () => ({ type: "outer", blur: 10, offset: 4, angle: 135, color: "000000", opacity: 0.28 });
const mkShadowS = () => ({ type: "outer", blur:  6, offset: 2, angle: 135, color: "000000", opacity: 0.20 });

// Standard footer added to every slide
function addFooter(slide) {
  slide.addShape("rect", {
    x: 0, y: 5.25, w: 10, h: 0.375,
    fill: { color: C.bgCard }, line: { color: C.bgCard },
  });
  slide.addText("kindkart  ·  Sponsor Update  ·  April 2026", {
    x: 0.35, y: 5.27, w: 9.3, h: 0.33,
    fontSize: 10, fontFace: "Calibri", color: C.gray,
    align: "right", margin: 0,
  });
}

// Orange section label
function addLabel(slide, text, x, y) {
  slide.addText(text, {
    x, y, w: 6, h: 0.32,
    fontSize: 10, fontFace: "Calibri", bold: true,
    color: C.orange, charSpacing: 4, margin: 0,
  });
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 1 · Cover
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  // Left orange accent bar
  s.addShape("rect", {
    x: 0, y: 0, w: 0.1, h: 5.625,
    fill: { color: C.orange }, line: { color: C.orange },
  });

  // Logo wordmark
  s.addText([
    { text: "kind", options: { color: C.white,  bold: true } },
    { text: "kart", options: { color: C.green,  bold: true } },
  ], { x: 0.35, y: 0.28, w: 4.5, h: 0.65, fontSize: 32, fontFace: "Arial Black", margin: 0 });

  // Tag line
  s.addText("Transparency Infrastructure for Charitable Giving", {
    x: 0.35, y: 0.97, w: 6.5, h: 0.32,
    fontSize: 12, fontFace: "Calibri", color: C.gray, italic: true, margin: 0,
  });

  // Main headline
  s.addText("A Personal Update", {
    x: 0.35, y: 1.55, w: 8, h: 0.9,
    fontSize: 50, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });
  s.addText("for Radhika", {
    x: 0.35, y: 2.45, w: 8, h: 0.9,
    fontSize: 50, fontFace: "Georgia", bold: true, color: C.orange, margin: 0,
  });

  // Date range
  s.addText("July 2025  →  April 2026", {
    x: 0.35, y: 3.5, w: 5, h: 0.38,
    fontSize: 15, fontFace: "Calibri", color: C.grayMid, margin: 0,
  });

  // Sub-note
  s.addText("Prepared by Sri G Gopalakrishnan · Founder, KindKart", {
    x: 0.35, y: 4.0, w: 6, h: 0.32,
    fontSize: 11, fontFace: "Calibri", color: C.gray, italic: true, margin: 0,
  });

  // Hero stat card (right side)
  s.addShape("rect", {
    x: 7.15, y: 0.95, w: 2.55, h: 2.75,
    fill: { color: C.bgCard }, line: { color: C.bgCard2, width: 1 },
    shadow: mkShadow(),
  });
  s.addShape("rect", {
    x: 7.15, y: 0.95, w: 2.55, h: 0.08,
    fill: { color: C.orange }, line: { color: C.orange },
  });
  s.addText("$1M+", {
    x: 7.15, y: 1.1, w: 2.55, h: 1.1,
    fontSize: 58, fontFace: "Georgia", bold: true,
    color: C.gold, align: "center", margin: 0,
  });
  s.addText("Total Donations\nSince Launch", {
    x: 7.15, y: 2.2, w: 2.55, h: 0.58,
    fontSize: 11, fontFace: "Calibri", color: C.grayMid,
    align: "center", margin: 0,
  });
  s.addText("First Horizon  ✓", {
    x: 7.15, y: 2.82, w: 2.55, h: 0.45,
    fontSize: 12, fontFace: "Calibri", bold: true,
    color: C.green, align: "center", margin: 0,
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 2 · Personal Note
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "PERSONAL NOTE", 0.5, 0.28);

  // Large decorative quote mark
  s.addText("\u201C", {
    x: 0.28, y: 0.5, w: 1, h: 1.2,
    fontSize: 110, fontFace: "Georgia", color: C.orange, margin: 0,
  });

  s.addText("Radhika,", {
    x: 0.7, y: 1.05, w: 9, h: 0.5,
    fontSize: 26, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });

  s.addText(
    "Thank you for your early and generous support of KindKart. Your $100K contribution gave us the freedom to build and scale — without compromise. This is a transparent look at what we have accomplished in the nine months since your support, how we have evolved, and where we are headed.",
    {
      x: 0.7, y: 1.6, w: 8.6, h: 1.0,
      fontSize: 15, fontFace: "Calibri", color: C.offWhite, margin: 0,
    }
  );

  // Three highlight pillars
  const pillars = [
    { label: "MILESTONE",  val: "$1M+",      sub: "Donations Crossed",     color: C.gold   },
    { label: "SCALE",      val: "86 Orgs",   sub: "across 14 States",       color: C.green  },
    { label: "MOMENTUM",   val: "$250K",      sub: "raised at our Gala",     color: C.orange },
  ];

  pillars.forEach((p, i) => {
    const x = 0.45 + i * 3.1;
    s.addShape("rect", {
      x, y: 2.88, w: 2.9, h: 1.95,
      fill: { color: C.bgCard }, line: { color: C.bgCard2, width: 1 },
      shadow: mkShadowS(),
    });
    s.addShape("rect", {
      x, y: 2.88, w: 2.9, h: 0.07,
      fill: { color: p.color }, line: { color: p.color },
    });
    s.addText(p.label, {
      x: x + 0.18, y: 3.03, w: 2.55, h: 0.3,
      fontSize: 9, fontFace: "Calibri", bold: true, color: p.color,
      charSpacing: 3, margin: 0,
    });
    s.addText(p.val, {
      x: x + 0.18, y: 3.37, w: 2.55, h: 0.72,
      fontSize: 34, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
    });
    s.addText(p.sub, {
      x: x + 0.18, y: 4.12, w: 2.55, h: 0.38,
      fontSize: 12, fontFace: "Calibri", color: C.grayMid, italic: true, margin: 0,
    });
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 3 · The Journey: July 2025 → April 2026
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "PROGRESS SINCE YOUR SUPPORT", 0.5, 0.28);
  s.addText("The Journey: July 2025 → April 2026", {
    x: 0.5, y: 0.62, w: 9, h: 0.68,
    fontSize: 30, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });

  // THEN panel
  s.addShape("rect", {
    x: 0.38, y: 1.52, w: 4.28, h: 3.48,
    fill: { color: C.bgCard }, line: { color: C.bgCard2, width: 1 },
    shadow: mkShadowS(),
  });
  s.addShape("rect", {
    x: 0.38, y: 1.52, w: 4.28, h: 0.07,
    fill: { color: C.gray }, line: { color: C.gray },
  });
  s.addText("WHEN YOU LAST HEARD FROM US", {
    x: 0.58, y: 1.67, w: 3.9, h: 0.3,
    fontSize: 9, fontFace: "Calibri", bold: true, color: C.grayMid,
    charSpacing: 2, margin: 0,
  });
  s.addText("July 2025", {
    x: 0.58, y: 2.0, w: 3.9, h: 0.45,
    fontSize: 18, fontFace: "Georgia", bold: true, color: C.gray, margin: 0,
  });

  const thenStats = [
    ["$550K",   "Total Donation Volume"],
    ["64",      "Charities Onboarded"],
    ["6,839",   "People Served Daily"],
    ["31",      "Cities & Villages"],
  ];
  thenStats.forEach(([val, lbl], i) => {
    s.addText(val, {
      x: 0.58, y: 2.6 + i * 0.58, w: 1.6, h: 0.45,
      fontSize: 24, fontFace: "Georgia", bold: true, color: C.grayMid, margin: 0,
    });
    s.addText(lbl, {
      x: 2.22, y: 2.65 + i * 0.58, w: 2.25, h: 0.4,
      fontSize: 12, fontFace: "Calibri", color: C.gray, valign: "middle", margin: 0,
    });
  });

  // Arrow
  s.addText("→", {
    x: 4.73, y: 2.95, w: 0.54, h: 0.65,
    fontSize: 32, fontFace: "Arial", color: C.orange, align: "center", margin: 0,
  });

  // NOW panel
  s.addShape("rect", {
    x: 5.34, y: 1.52, w: 4.28, h: 3.48,
    fill: { color: C.bgCard }, line: { color: C.orange, width: 1 },
    shadow: mkShadow(),
  });
  s.addShape("rect", {
    x: 5.34, y: 1.52, w: 4.28, h: 0.07,
    fill: { color: C.orange }, line: { color: C.orange },
  });
  s.addText("WHERE WE ARE TODAY", {
    x: 5.54, y: 1.67, w: 3.9, h: 0.3,
    fontSize: 9, fontFace: "Calibri", bold: true, color: C.orange,
    charSpacing: 2, margin: 0,
  });
  s.addText("April 2026", {
    x: 5.54, y: 2.0, w: 3.9, h: 0.45,
    fontSize: 18, fontFace: "Georgia", bold: true, color: C.offWhite, margin: 0,
  });

  const nowStats = [
    ["$1M+",    "Total Donation Volume"],
    ["86",      "Grassroots Orgs Onboarded"],
    ["10,000+", "People Served Daily"],
    ["34",      "Locations Across India"],
  ];
  nowStats.forEach(([val, lbl], i) => {
    s.addText(val, {
      x: 5.54, y: 2.6 + i * 0.58, w: 1.85, h: 0.45,
      fontSize: 24, fontFace: "Georgia", bold: true, color: C.gold, margin: 0,
    });
    s.addText(lbl, {
      x: 7.42, y: 2.65 + i * 0.58, w: 2.0, h: 0.4,
      fontSize: 12, fontFace: "Calibri", color: C.grayMid, valign: "middle", margin: 0,
    });
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 4 · $1M Milestone
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  // Subtle decorative ring
  s.addShape("ellipse", {
    x: 3.8, y: -0.5, w: 6.0, h: 6.0,
    fill: { color: C.gold, transparency: 94 },
    line: { color: C.bgDark },
  });

  addLabel(s, "MILESTONE", 0.5, 0.28);

  s.addText("We Crossed", {
    x: 0.5, y: 0.72, w: 9, h: 0.72,
    fontSize: 36, fontFace: "Georgia", color: C.offWhite, margin: 0,
  });
  s.addText("$1,000,000", {
    x: 0.5, y: 1.4, w: 9.2, h: 1.5,
    fontSize: 100, fontFace: "Georgia", bold: true, color: C.gold, margin: 0,
  });
  s.addText("in Total Donation Volume  —  In Just Over Two Years", {
    x: 0.5, y: 2.92, w: 9, h: 0.55,
    fontSize: 22, fontFace: "Georgia", color: C.offWhite, margin: 0,
  });

  const proofs = [
    { icon: "✓", text: "Validates the KindKart model",         color: C.green  },
    { icon: "✓", text: "Demonstrates strong donor trust",       color: C.green  },
    { icon: "✓", text: "Increasing scale and momentum",         color: C.green  },
  ];
  proofs.forEach((p, i) => {
    s.addShape("ellipse", {
      x: 0.5 + i * 2.9, y: 3.72, w: 0.42, h: 0.42,
      fill: { color: p.color }, line: { color: p.color },
    });
    s.addText(p.icon, {
      x: 0.5 + i * 2.9, y: 3.72, w: 0.42, h: 0.42,
      fontSize: 14, fontFace: "Calibri", bold: true,
      color: C.white, align: "center", valign: "middle", margin: 0,
    });
    s.addText(p.text, {
      x: 1.05 + i * 2.9, y: 3.76, w: 2.55, h: 0.36,
      fontSize: 14, fontFace: "Calibri", color: C.offWhite, valign: "middle", margin: 0,
    });
  });

  s.addText("This is our First Horizon — achieved from launch in August 2023 to April 2026.", {
    x: 0.5, y: 4.3, w: 9, h: 0.35,
    fontSize: 13, fontFace: "Calibri", italic: true, color: C.gray, margin: 0,
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 5 · Scale Across India
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "SCALE", 0.5, 0.28);
  s.addText("Reaching Across India", {
    x: 0.5, y: 0.62, w: 9, h: 0.62,
    fontSize: 32, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });

  const stats = [
    { num: "86",    label: "Grassroots Orgs\nOnboarded",   sub: "+34% since July 2025",     color: C.orange },
    { num: "14",    label: "States Across\nIndia",          sub: "Pan-India presence",        color: C.green  },
    { num: "34",    label: "Cities &\nLocations",           sub: "Deep community reach",      color: C.gold   },
    { num: "10K+",  label: "People Served\nEvery Day",      sub: "Daily essential support",   color: C.teal   },
  ];

  const positions = [
    { x: 0.4,  y: 1.55 },
    { x: 5.1,  y: 1.55 },
    { x: 0.4,  y: 3.3  },
    { x: 5.1,  y: 3.3  },
  ];

  stats.forEach((st, i) => {
    const { x, y } = positions[i];
    s.addShape("rect", {
      x, y, w: 4.38, h: 1.55,
      fill: { color: C.bgCard }, line: { color: C.bgCard2, width: 1 },
      shadow: mkShadowS(),
    });
    s.addShape("rect", {
      x, y, w: 0.08, h: 1.55,
      fill: { color: st.color }, line: { color: st.color },
    });
    s.addText(st.num, {
      x: x + 0.22, y: y + 0.14, w: 1.65, h: 0.88,
      fontSize: 54, fontFace: "Georgia", bold: true, color: st.color, margin: 0,
    });
    s.addText(st.label, {
      x: x + 1.92, y: y + 0.18, w: 2.25, h: 0.65,
      fontSize: 16, fontFace: "Calibri", bold: true, color: C.white, margin: 0,
    });
    s.addText(st.sub, {
      x: x + 1.92, y: y + 0.9, w: 2.25, h: 0.35,
      fontSize: 11, fontFace: "Calibri", color: C.gray, margin: 0,
    });
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 6 · Depth of Impact
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "IMPACT", 0.5, 0.28);
  s.addText("Predictable Support, Real Change", {
    x: 0.5, y: 0.62, w: 9, h: 0.62,
    fontSize: 30, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });

  // Left hero stat
  s.addShape("rect", {
    x: 0.38, y: 1.48, w: 4.28, h: 3.5,
    fill: { color: C.bgCard }, line: { color: C.orange, width: 1 },
    shadow: mkShadow(),
  });
  s.addShape("rect", {
    x: 0.38, y: 1.48, w: 4.28, h: 0.07,
    fill: { color: C.orange }, line: { color: C.orange },
  });
  s.addText("40–50%", {
    x: 0.48, y: 1.62, w: 4.08, h: 1.3,
    fontSize: 76, fontFace: "Georgia", bold: true, color: C.gold,
    align: "center", margin: 0,
  });
  s.addText("of Daily Essential Needs", {
    x: 0.48, y: 2.95, w: 4.08, h: 0.5,
    fontSize: 18, fontFace: "Georgia", color: C.offWhite,
    align: "center", margin: 0,
  });
  s.addText("covered for most partner organizations", {
    x: 0.48, y: 3.47, w: 4.08, h: 0.38,
    fontSize: 13, fontFace: "Calibri", italic: true, color: C.gray,
    align: "center", margin: 0,
  });
  s.addText("Consistent  &  Predictable", {
    x: 0.48, y: 4.03, w: 4.08, h: 0.58,
    fontSize: 16, fontFace: "Calibri", bold: true, color: C.green,
    align: "center", margin: 0,
  });

  // Right: impact numbers
  const items = [
    { num: "2.15M",  label: "Acts of Kindness",               color: C.orange },
    { num: "182K",   label: "KGs of Groceries Delivered",      color: C.green  },
    { num: "787K",   label: "Meals Served",                    color: C.gold   },
    { num: "65K",    label: "Cups of Milk Delivered",          color: C.teal   },
  ];

  items.forEach((it, i) => {
    const y = 1.48 + i * 0.875;
    s.addShape("rect", {
      x: 5.05, y, w: 4.55, h: 0.75,
      fill: { color: C.bgCard }, line: { color: C.bgCard2, width: 1 },
      shadow: mkShadowS(),
    });
    s.addShape("rect", {
      x: 5.05, y, w: 0.07, h: 0.75,
      fill: { color: it.color }, line: { color: it.color },
    });
    s.addText(it.num, {
      x: 5.2, y: y + 0.04, w: 1.65, h: 0.62,
      fontSize: 30, fontFace: "Georgia", bold: true, color: it.color, margin: 0,
    });
    s.addText(it.label, {
      x: 6.88, y: y + 0.1, w: 2.52, h: 0.52,
      fontSize: 14, fontFace: "Calibri", color: C.offWhite, valign: "middle", margin: 0,
    });
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 7 · Gala Highlight
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "EVENT HIGHLIGHT", 0.5, 0.28);
  s.addText("The First KindKart Gala", {
    x: 0.5, y: 0.62, w: 9, h: 0.72,
    fontSize: 36, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });
  s.addText("September 14, 2025  ·  Computer History Museum  ·  Mountain View, CA", {
    x: 0.5, y: 1.35, w: 9, h: 0.38,
    fontSize: 14, fontFace: "Calibri", italic: true, color: C.grayMid, margin: 0,
  });

  const galaStats = [
    { num: "200+",   label: "Attendees",      sub: "An evening of inspiration",      color: C.orange },
    { num: "$250K",  label: "Raised",          sub: "Exceeded every expectation",     color: C.gold   },
    { num: "1st",    label: "Annual Gala",     sub: "A new KindKart tradition",       color: C.green  },
  ];

  galaStats.forEach((gs, i) => {
    const x = 0.42 + i * 3.1;
    s.addShape("rect", {
      x, y: 1.9, w: 2.92, h: 2.65,
      fill: { color: C.bgCard }, line: { color: C.bgCard2, width: 1 },
      shadow: mkShadow(),
    });
    s.addShape("rect", {
      x, y: 1.9, w: 2.92, h: 0.08,
      fill: { color: gs.color }, line: { color: gs.color },
    });
    s.addText(gs.num, {
      x: x + 0.14, y: 2.1, w: 2.64, h: 1.2,
      fontSize: 68, fontFace: "Georgia", bold: true,
      color: gs.color, align: "center", margin: 0,
    });
    s.addText(gs.label, {
      x: x + 0.14, y: 3.32, w: 2.64, h: 0.45,
      fontSize: 18, fontFace: "Calibri", bold: true,
      color: C.white, align: "center", margin: 0,
    });
    s.addText(gs.sub, {
      x: x + 0.14, y: 3.8, w: 2.64, h: 0.45,
      fontSize: 12, fontFace: "Calibri", italic: true,
      color: C.gray, align: "center", margin: 0,
    });
  });

  s.addText(
    "\u201CA night of light, where stories of kindness shine ever bright — one heart at a time.\u201D",
    {
      x: 0.5, y: 4.78, w: 9, h: 0.32,
      fontSize: 12, fontFace: "Georgia", italic: true, color: C.gray,
      align: "center", margin: 0,
    }
  );

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 8 · Vision Evolution — Transparency Infrastructure
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "VISION EVOLUTION", 0.5, 0.28);
  s.addText("The Transparency Infrastructure for Philanthropy", {
    x: 0.5, y: 0.62, w: 9, h: 0.62,
    fontSize: 26, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });

  s.addText("We started by delivering essentials. We are now building the transparency layer for all of social giving.", {
    x: 0.5, y: 1.28, w: 9, h: 0.38,
    fontSize: 13, fontFace: "Calibri", color: C.grayMid, margin: 0,
  });

  // 5-row segment table
  const segments = [
    { segment: "Donors, Foundations,\nCorporates", problem: "Lack of transparency in\nimpact and spending", value: "Show exactly where your donation\ngoes — and the impact it creates", color: C.orange },
    { segment: "Grassroot\nOrganizations", problem: "Predictable recurring expenses\nbut no steady cashflow", value: "Minimum guaranteed support\nin kind for predictable expenses", color: C.green },
    { segment: "Traditional\nNon-Profits", problem: "Operational inefficiency\nlimiting impact and speed", value: "Infrastructure for charitable ops\n(ServiceNow for nonprofits)", color: C.teal },
    { segment: "Non-Profits Funded\nby CSR", problem: "Inefficiency and manual\nreporting of impact", value: "Power corporate impact\ntransparency, eliminate reporting", color: C.gold },
    { segment: "Giving\nChannels", problem: "Lost opportunity communicating\nimpact back to donors", value: "Simplified integration for\ntransparent donor visibility", color: C.blue },
  ];

  // Column headers
  const headerY = 1.82;
  s.addShape("rect", {
    x: 0.38, y: headerY, w: 9.24, h: 0.38,
    fill: { color: C.orange }, line: { color: C.orange },
  });
  s.addText("Segment", {
    x: 0.5, y: headerY, w: 2.6, h: 0.38,
    fontSize: 11, fontFace: "Calibri", bold: true, color: C.white, valign: "middle", margin: 0,
  });
  s.addText("Problem", {
    x: 3.2, y: headerY, w: 3.0, h: 0.38,
    fontSize: 11, fontFace: "Calibri", bold: true, color: C.white, valign: "middle", margin: 0,
  });
  s.addText("Value Proposition", {
    x: 6.35, y: headerY, w: 3.2, h: 0.38,
    fontSize: 11, fontFace: "Calibri", bold: true, color: C.white, valign: "middle", margin: 0,
  });

  segments.forEach((seg, i) => {
    const y = 2.22 + i * 0.6;
    const bgColor = i % 2 === 0 ? C.bgCard : C.bgCard2;
    s.addShape("rect", {
      x: 0.38, y, w: 9.24, h: 0.58,
      fill: { color: bgColor }, line: { color: bgColor },
    });
    // Left accent
    s.addShape("rect", {
      x: 0.38, y, w: 0.06, h: 0.58,
      fill: { color: seg.color }, line: { color: seg.color },
    });
    s.addText(seg.segment, {
      x: 0.55, y, w: 2.55, h: 0.58,
      fontSize: 10, fontFace: "Calibri", bold: true, color: seg.color, valign: "middle", margin: 0,
    });
    s.addText(seg.problem, {
      x: 3.2, y, w: 3.0, h: 0.58,
      fontSize: 10, fontFace: "Calibri", color: C.offWhite, valign: "middle", margin: 0,
    });
    s.addText(seg.value, {
      x: 6.35, y, w: 3.2, h: 0.58,
      fontSize: 10, fontFace: "Calibri", color: C.grayMid, valign: "middle", margin: 0,
    });
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 9 · Functional Architecture — Layers of Social Giving
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "PLATFORM ARCHITECTURE", 0.5, 0.28);
  s.addText("Layers of Social Giving", {
    x: 0.5, y: 0.62, w: 9, h: 0.62,
    fontSize: 30, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });

  // Layer definitions (top to bottom)
  const layers = [
    {
      label: "LAYER 4",
      title: "DONORS",
      desc: "Individual Donors  ·  Corporates  ·  Private Foundations  ·  Public Foundations",
      color: C.orange,
      y: 1.42,
    },
    {
      label: "LAYER 3",
      title: "GIVING CHANNELS",
      desc: "Give India  ·  Global Giving  ·  Aid India  ·  Other Platforms",
      sub: "Aggregators that embed KindKart's transparency infrastructure",
      color: C.blue,
      y: 2.32,
    },
    {
      label: "LAYER 2",
      title: "PARTNER NON-PROFITS SaaS",
      desc: "Education ✓  ·  Women Empowerment  ·  Health  ·  Skill & Vocational Training",
      sub: "Operational efficiency platform organized by vertical",
      color: C.teal,
      y: 3.22,
    },
    {
      label: "LAYER 1",
      title: "CORE KINDKART MARKETPLACE",
      desc: "Recurring, in-kind support for grassroots organizations' daily essential needs",
      sub: "82 orgs  ·  15 states  ·  36 cities  ·  $1M+ donation volume",
      color: C.green,
      y: 4.12,
    },
  ];

  layers.forEach((layer) => {
    // Card background
    s.addShape("rect", {
      x: 0.38, y: layer.y, w: 9.24, h: 0.82,
      fill: { color: C.bgCard }, line: { color: layer.color, width: 1 },
      shadow: mkShadowS(),
    });
    // Left accent bar
    s.addShape("rect", {
      x: 0.38, y: layer.y, w: 0.08, h: 0.82,
      fill: { color: layer.color }, line: { color: layer.color },
    });
    // Layer label
    s.addText(layer.label, {
      x: 0.58, y: layer.y + 0.05, w: 1.0, h: 0.28,
      fontSize: 8, fontFace: "Calibri", bold: true, color: layer.color,
      charSpacing: 2, margin: 0,
    });
    // Title
    s.addText(layer.title, {
      x: 1.55, y: layer.y + 0.02, w: 4.2, h: 0.34,
      fontSize: 14, fontFace: "Calibri", bold: true, color: C.white, margin: 0,
    });
    // Description
    s.addText(layer.desc, {
      x: 1.55, y: layer.y + 0.34, w: 7.8, h: 0.22,
      fontSize: 10, fontFace: "Calibri", color: C.offWhite, margin: 0,
    });
    // Sub-description
    if (layer.sub) {
      s.addText(layer.sub, {
        x: 1.55, y: layer.y + 0.56, w: 7.8, h: 0.2,
        fontSize: 9, fontFace: "Calibri", italic: true, color: C.gray, margin: 0,
      });
    }
  });

  // Arrows between layers
  [1.42, 2.32, 3.22].forEach((y) => {
    s.addText("▼", {
      x: 4.75, y: y + 0.78, w: 0.5, h: 0.22,
      fontSize: 12, fontFace: "Arial", color: C.grayMid, align: "center", margin: 0,
    });
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 10 · Key Product Launches — Shipped
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "PRODUCT PORTFOLIO", 0.5, 0.28);
  s.addText("Products We Have Shipped", {
    x: 0.5, y: 0.62, w: 9, h: 0.62,
    fontSize: 30, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });

  const shipped = [
    { name: "KindKart Marketplace", audience: "Donors, Beneficiaries", year: "2023", color: C.orange },
    { name: "Charity App 2.0", audience: "Beneficiaries", year: "2024", color: C.orange },
    { name: "India Donor Marketplace", audience: "Donors (India)", year: "2024", color: C.orange },
    { name: "Scholarship Mgmt Platform", audience: "Partners", year: "2025", color: C.green },
    { name: "KindKash", audience: "Donors", year: "2025", color: C.green },
    { name: "Transparency Report", audience: "Donors", year: "2025", color: C.green },
  ];

  // Two rows of three cards
  shipped.forEach((p, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.4 + col * 3.1;
    const y = 1.45 + row * 1.95;

    s.addShape("rect", {
      x, y, w: 2.92, h: 1.75,
      fill: { color: C.bgCard }, line: { color: C.bgCard2, width: 1 },
      shadow: mkShadowS(),
    });
    s.addShape("rect", {
      x, y, w: 2.92, h: 0.07,
      fill: { color: p.color }, line: { color: p.color },
    });
    // Year badge
    s.addShape("rect", {
      x: x + 1.92, y: y + 0.18, w: 0.82, h: 0.28,
      fill: { color: p.color, transparency: 75 }, line: { color: p.color, width: 0.5 },
      rectRadius: 0.05,
    });
    s.addText(p.year, {
      x: x + 1.92, y: y + 0.18, w: 0.82, h: 0.28,
      fontSize: 10, fontFace: "Calibri", bold: true, color: p.color,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(p.name, {
      x: x + 0.15, y: y + 0.55, w: 2.62, h: 0.55,
      fontSize: 15, fontFace: "Calibri", bold: true, color: C.white, margin: 0,
    });
    s.addText(p.audience, {
      x: x + 0.15, y: y + 1.15, w: 2.62, h: 0.35,
      fontSize: 11, fontFace: "Calibri", color: C.gray, margin: 0,
    });
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 11 · Key Product Launches — 2026
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "2026 PRODUCT ROADMAP", 0.5, 0.28);
  s.addText("Building Next", {
    x: 0.5, y: 0.62, w: 9, h: 0.62,
    fontSize: 30, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });

  const upcoming = [
    { name: "Learning Management\nPlatform", audience: "Partners", color: C.teal },
    { name: "Volunteer Management\nPlatform", audience: "Partners", color: C.teal },
    { name: "Student Mentoring\nPlatform", audience: "Partners", color: C.teal },
    { name: "Logistics Management\nPlatform", audience: "Internal", color: C.gold },
    { name: "Digital Delivery\nVerification", audience: "Partners, Corporates", color: C.gold },
    { name: "AI Literacy\nCurriculum", audience: "Govt School Teachers", color: C.purple },
  ];

  upcoming.forEach((p, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.4 + col * 3.1;
    const y = 1.45 + row * 1.95;

    s.addShape("rect", {
      x, y, w: 2.92, h: 1.75,
      fill: { color: C.bgCard }, line: { color: p.color, width: 1 },
      shadow: mkShadowS(),
    });
    s.addShape("rect", {
      x, y, w: 2.92, h: 0.07,
      fill: { color: p.color }, line: { color: p.color },
    });
    // "NEW" badge
    s.addShape("rect", {
      x: x + 2.02, y: y + 0.18, w: 0.72, h: 0.28,
      fill: { color: p.color, transparency: 75 }, line: { color: p.color, width: 0.5 },
      rectRadius: 0.05,
    });
    s.addText("NEW", {
      x: x + 2.02, y: y + 0.18, w: 0.72, h: 0.28,
      fontSize: 9, fontFace: "Calibri", bold: true, color: p.color,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(p.name, {
      x: x + 0.15, y: y + 0.55, w: 2.62, h: 0.65,
      fontSize: 15, fontFace: "Calibri", bold: true, color: C.white, margin: 0,
    });
    s.addText(p.audience, {
      x: x + 0.15, y: y + 1.25, w: 2.62, h: 0.3,
      fontSize: 11, fontFace: "Calibri", color: C.gray, margin: 0,
    });
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 12 · AI-Native Development
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "TECHNOLOGY", 0.5, 0.28);
  s.addText("AI-Native Development", {
    x: 0.5, y: 0.62, w: 9, h: 0.62,
    fontSize: 30, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });

  s.addText("We are transitioning to an AI-driven software development model — enabling us to ship faster, at lower cost, and with a leaner team.", {
    x: 0.5, y: 1.3, w: 9, h: 0.45,
    fontSize: 14, fontFace: "Calibri", color: C.grayMid, margin: 0,
  });

  const aiItems = [
    { title: "Faster Iteration", desc: "AI-powered development accelerates feature delivery across all products", color: C.teal },
    { title: "Lower Cost", desc: "Reduced engineering overhead through AI code generation and agentic workflows", color: C.green },
    { title: "Agentic Workflows", desc: "AI agents in production for scholarship selection, impact reporting, and operations", color: C.orange },
  ];

  aiItems.forEach((item, i) => {
    const x = 0.4 + i * 3.17;
    s.addShape("rect", {
      x, y: 2.0, w: 3.0, h: 2.6,
      fill: { color: C.bgCard }, line: { color: C.bgCard2, width: 1 },
      shadow: mkShadowS(),
    });
    s.addShape("rect", {
      x, y: 2.0, w: 3.0, h: 0.07,
      fill: { color: item.color }, line: { color: item.color },
    });
    s.addText(item.title, {
      x: x + 0.18, y: 2.2, w: 2.65, h: 0.5,
      fontSize: 18, fontFace: "Georgia", bold: true, color: item.color, margin: 0,
    });
    s.addText(item.desc, {
      x: x + 0.18, y: 2.82, w: 2.65, h: 1.2,
      fontSize: 13, fontFace: "Calibri", color: C.offWhite, margin: 0,
    });
  });

  s.addText("This approach allows KindKart to build 6 new products in 2026 with a fraction of traditional engineering cost.", {
    x: 0.5, y: 4.75, w: 9, h: 0.32,
    fontSize: 12, fontFace: "Calibri", italic: true, color: C.gray, margin: 0,
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 13 · Leadership Team
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "LEADERSHIP", 0.5, 0.28);
  s.addText("The Team Behind the Mission", {
    x: 0.5, y: 0.62, w: 9, h: 0.62,
    fontSize: 30, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });

  const team = [
    { initials: "SG", name: "Sri G\nGopalakrishnan", role: "Founder",                    bg: "Ex-VP PayPal & eBay",           color: C.orange },
    { initials: "RM", name: "Ram Memuri",            role: "Head of Logistics",           bg: "Operations across India",       color: C.green  },
    { initials: "AN", name: "Anshuman",              role: "Head of Product & Tech",      bg: "Platform leadership",           color: C.teal   },
    { initials: "AR", name: "Abhishek Ranjan",       role: "Chief Impact Officer",        bg: "Grassroots partnerships",       color: C.gold   },
  ];

  team.forEach((t, i) => {
    const cardX = 0.4 + i * 2.37;

    // Avatar circle
    s.addShape("ellipse", {
      x: cardX + 0.64, y: 1.5, w: 1.08, h: 1.08,
      fill: { color: t.color, transparency: 25 },
      line: { color: t.color, width: 2 },
    });
    s.addText(t.initials, {
      x: cardX + 0.64, y: 1.5, w: 1.08, h: 1.08,
      fontSize: 28, fontFace: "Georgia", bold: true,
      color: C.white, align: "center", valign: "middle", margin: 0,
    });

    // Card
    s.addShape("rect", {
      x: cardX, y: 2.72, w: 2.2, h: 2.2,
      fill: { color: C.bgCard }, line: { color: C.bgCard2, width: 1 },
      shadow: mkShadowS(),
    });
    s.addShape("rect", {
      x: cardX, y: 2.72, w: 2.2, h: 0.07,
      fill: { color: t.color }, line: { color: t.color },
    });
    s.addText(t.name, {
      x: cardX + 0.12, y: 2.85, w: 1.98, h: 0.65,
      fontSize: 13, fontFace: "Calibri", bold: true, color: C.white, margin: 0,
    });
    s.addText(t.role, {
      x: cardX + 0.12, y: 3.52, w: 1.98, h: 0.5,
      fontSize: 12, fontFace: "Calibri", bold: true, color: t.color, margin: 0,
    });
    s.addText(t.bg, {
      x: cardX + 0.12, y: 4.05, w: 1.98, h: 0.5,
      fontSize: 11, fontFace: "Calibri", italic: true, color: C.gray, margin: 0,
    });
  });

  s.addText("+ Head of Charity & Volunteer Operations (former ServiceNow)", {
    x: 0.5, y: 4.96, w: 9, h: 0.26,
    fontSize: 11, fontFace: "Calibri", italic: true, color: C.gray, margin: 0,
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 14 · 2026 OKRs
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "2026 OBJECTIVES", 0.5, 0.28);
  s.addText("Where We Are Headed", {
    x: 0.5, y: 0.62, w: 9, h: 0.62,
    fontSize: 30, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });

  const okrs = [
    { obj: "Scale Donation Volume",        target: "$1.2M",   detail: "2x from 2025",                color: C.gold   },
    { obj: "Grow Grassroots Network",      target: "150",     detail: "+50 new orgs",                 color: C.green  },
    { obj: "Unlock Corporate Giving",      target: "$500K",   detail: "10x increase from 2025",       color: C.orange },
    { obj: "Compliance Standard",          target: "100%",    detail: "Monthly audit, all orgs",       color: C.teal   },
    { obj: "Launch Non-Profit SaaS",       target: "25",      detail: "Partners → $100K revenue",     color: C.blue   },
    { obj: "Improve Product Margins",      target: "25%",     detail: "Optimize top 10 products",     color: C.purple },
    { obj: "Fulfillment Excellence",       target: "95%",     detail: "Delivered within 7 days",       color: C.pink   },
  ];

  okrs.forEach((okr, i) => {
    const y = 1.42 + i * 0.52;
    const bgColor = i % 2 === 0 ? C.bgCard : C.bgCard2;

    s.addShape("rect", {
      x: 0.38, y, w: 9.24, h: 0.48,
      fill: { color: bgColor }, line: { color: bgColor },
    });
    // Left accent
    s.addShape("rect", {
      x: 0.38, y, w: 0.07, h: 0.48,
      fill: { color: okr.color }, line: { color: okr.color },
    });
    // Objective
    s.addText(okr.obj, {
      x: 0.6, y, w: 4.0, h: 0.48,
      fontSize: 13, fontFace: "Calibri", bold: true, color: C.white, valign: "middle", margin: 0,
    });
    // Target
    s.addText(okr.target, {
      x: 4.7, y, w: 1.6, h: 0.48,
      fontSize: 22, fontFace: "Georgia", bold: true, color: okr.color,
      align: "center", valign: "middle", margin: 0,
    });
    // Detail
    s.addText(okr.detail, {
      x: 6.4, y, w: 3.1, h: 0.48,
      fontSize: 11, fontFace: "Calibri", color: C.grayMid, valign: "middle", margin: 0,
    });
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 15 · Legal & Compliance — Entity Structure
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "LEGAL & COMPLIANCE", 0.5, 0.28);
  s.addText("Four-Entity Structure", {
    x: 0.5, y: 0.62, w: 9, h: 0.62,
    fontSize: 30, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });

  s.addText("Legal counsel: Oban and Mason (New Delhi) — referred by Priya Devaan, Chief Legal Counsel at Akasa Air", {
    x: 0.5, y: 1.25, w: 9, h: 0.32,
    fontSize: 12, fontFace: "Calibri", italic: true, color: C.grayMid, margin: 0,
  });

  // Singapore HoldCo (top center)
  s.addShape("rect", {
    x: 3.2, y: 1.72, w: 3.6, h: 0.95,
    fill: { color: C.bgCard }, line: { color: C.gold, width: 1 },
    shadow: mkShadowS(),
  });
  s.addShape("rect", {
    x: 3.2, y: 1.72, w: 3.6, h: 0.06,
    fill: { color: C.gold }, line: { color: C.gold },
  });
  s.addText("Singapore Holding Company", {
    x: 3.3, y: 1.82, w: 3.4, h: 0.35,
    fontSize: 13, fontFace: "Calibri", bold: true, color: C.gold, margin: 0,
  });
  s.addText("IP, trademarks, software assets", {
    x: 3.3, y: 2.18, w: 3.4, h: 0.25,
    fontSize: 10, fontFace: "Calibri", color: C.grayMid, margin: 0,
  });
  // Status badge
  s.addText("IN EXECUTION", {
    x: 5.55, y: 2.38, w: 1.15, h: 0.22,
    fontSize: 8, fontFace: "Calibri", bold: true, color: C.gold, margin: 0,
  });

  // Arrow down-left to KindKart US
  s.addText("↙", {
    x: 2.65, y: 2.62, w: 0.6, h: 0.4,
    fontSize: 22, fontFace: "Arial", color: C.grayMid, align: "center", margin: 0,
  });
  // Arrow down-right to India Pvt
  s.addText("↘", {
    x: 6.75, y: 2.62, w: 0.6, h: 0.4,
    fontSize: 22, fontFace: "Arial", color: C.grayMid, align: "center", margin: 0,
  });

  // KindKart Inc US (left)
  s.addShape("rect", {
    x: 0.38, y: 3.05, w: 4.2, h: 0.95,
    fill: { color: C.bgCard }, line: { color: C.orange, width: 1 },
    shadow: mkShadowS(),
  });
  s.addShape("rect", {
    x: 0.38, y: 3.05, w: 4.2, h: 0.06,
    fill: { color: C.orange }, line: { color: C.orange },
  });
  s.addText("KindKart Inc (US)", {
    x: 0.5, y: 3.15, w: 3.0, h: 0.35,
    fontSize: 13, fontFace: "Calibri", bold: true, color: C.orange, margin: 0,
  });
  s.addText("501(c)(3)  ·  US donor marketplace, fundraising, grants", {
    x: 0.5, y: 3.52, w: 3.95, h: 0.25,
    fontSize: 10, fontFace: "Calibri", color: C.grayMid, margin: 0,
  });
  s.addText("ESTABLISHED", {
    x: 3.38, y: 3.15, w: 1.1, h: 0.22,
    fontSize: 8, fontFace: "Calibri", bold: true, color: C.green, margin: 0,
  });

  // India Pvt Ltd (right)
  s.addShape("rect", {
    x: 5.42, y: 3.05, w: 4.2, h: 0.95,
    fill: { color: C.bgCard }, line: { color: C.teal, width: 1 },
    shadow: mkShadowS(),
  });
  s.addShape("rect", {
    x: 5.42, y: 3.05, w: 4.2, h: 0.06,
    fill: { color: C.teal }, line: { color: C.teal },
  });
  s.addText("India Private Limited", {
    x: 5.55, y: 3.15, w: 3.0, h: 0.35,
    fontSize: 13, fontFace: "Calibri", bold: true, color: C.teal, margin: 0,
  });
  s.addText("SaaS reseller/distributor + logistics operations", {
    x: 5.55, y: 3.52, w: 3.95, h: 0.25,
    fontSize: 10, fontFace: "Calibri", color: C.grayMid, margin: 0,
  });
  s.addText("IN EXECUTION", {
    x: 8.35, y: 3.15, w: 1.15, h: 0.22,
    fontSize: 8, fontFace: "Calibri", bold: true, color: C.gold, margin: 0,
  });

  // Arrows to India Foundation
  s.addText("↓", {
    x: 2.3, y: 3.98, w: 0.5, h: 0.3,
    fontSize: 18, fontFace: "Arial", color: C.grayMid, align: "center", margin: 0,
  });
  s.addText("↓", {
    x: 7.3, y: 3.98, w: 0.5, h: 0.3,
    fontSize: 18, fontFace: "Arial", color: C.grayMid, align: "center", margin: 0,
  });
  s.addText("funds", {
    x: 1.5, y: 3.98, w: 1.0, h: 0.28,
    fontSize: 9, fontFace: "Calibri", italic: true, color: C.gray, align: "center", margin: 0,
  });
  s.addText("revenue funds", {
    x: 7.8, y: 3.98, w: 1.5, h: 0.28,
    fontSize: 9, fontFace: "Calibri", italic: true, color: C.gray, margin: 0,
  });

  // India Foundation (bottom center)
  s.addShape("rect", {
    x: 2.2, y: 4.28, w: 5.6, h: 0.82,
    fill: { color: C.bgCard }, line: { color: C.green, width: 1 },
    shadow: mkShadow(),
  });
  s.addShape("rect", {
    x: 2.2, y: 4.28, w: 5.6, h: 0.06,
    fill: { color: C.green }, line: { color: C.green },
  });
  s.addText("KindKart India Foundation", {
    x: 2.35, y: 4.38, w: 3.5, h: 0.32,
    fontSize: 14, fontFace: "Calibri", bold: true, color: C.green, margin: 0,
  });
  s.addText("Section 8 Company  ·  Charitable operations in India  ·  Grassroots org support", {
    x: 2.35, y: 4.72, w: 5.3, h: 0.25,
    fontSize: 10, fontFace: "Calibri", color: C.grayMid, margin: 0,
  });
  s.addText("ESTABLISHED", {
    x: 6.55, y: 4.38, w: 1.1, h: 0.22,
    fontSize: 8, fontFace: "Calibri", bold: true, color: C.green, margin: 0,
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 16 · Strategic Intent — Funding Loop
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  addLabel(s, "STRATEGIC INTENT", 0.5, 0.28);
  s.addText("A Self-Reinforcing Funding Loop", {
    x: 0.5, y: 0.62, w: 9, h: 0.62,
    fontSize: 30, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });

  s.addText("Licensing revenue from the SaaS platform — combined with donations — funds grassroots charitable operations. Commercial scale directly amplifies charitable impact.", {
    x: 0.5, y: 1.35, w: 9, h: 0.55,
    fontSize: 14, fontFace: "Calibri", color: C.grayMid, margin: 0,
  });

  // Three-part loop visualization
  const loopItems = [
    {
      title: "SaaS Licensing\nRevenue",
      desc: "25 non-profit partners pay for operational efficiency platform",
      color: C.teal,
      x: 0.4, y: 2.2,
    },
    {
      title: "Donor\nContributions",
      desc: "Individual, corporate, and foundation donations via marketplace",
      color: C.orange,
      x: 3.5, y: 2.2,
    },
    {
      title: "Grassroots\nImpact",
      desc: "100% of combined revenue funds predictable essential support",
      color: C.green,
      x: 6.6, y: 2.2,
    },
  ];

  loopItems.forEach((item) => {
    s.addShape("rect", {
      x: item.x, y: item.y, w: 2.92, h: 2.4,
      fill: { color: C.bgCard }, line: { color: item.color, width: 1 },
      shadow: mkShadow(),
    });
    s.addShape("rect", {
      x: item.x, y: item.y, w: 2.92, h: 0.07,
      fill: { color: item.color }, line: { color: item.color },
    });
    s.addText(item.title, {
      x: item.x + 0.18, y: item.y + 0.2, w: 2.56, h: 0.72,
      fontSize: 20, fontFace: "Georgia", bold: true, color: item.color, margin: 0,
    });
    s.addText(item.desc, {
      x: item.x + 0.18, y: item.y + 1.05, w: 2.56, h: 1.0,
      fontSize: 12, fontFace: "Calibri", color: C.offWhite, margin: 0,
    });
  });

  // Arrows between
  s.addText("  +  ", {
    x: 3.15, y: 3.1, w: 0.5, h: 0.5,
    fontSize: 22, fontFace: "Georgia", bold: true, color: C.gold, align: "center", margin: 0,
  });
  s.addText("  →  ", {
    x: 6.25, y: 3.1, w: 0.5, h: 0.5,
    fontSize: 22, fontFace: "Georgia", bold: true, color: C.gold, align: "center", margin: 0,
  });

  s.addText("\"At scale, our impact grows faster than our costs — and our operating costs approach zero.\"", {
    x: 0.5, y: 4.78, w: 9, h: 0.32,
    fontSize: 12, fontFace: "Georgia", italic: true, color: C.gray, align: "center", margin: 0,
  });

  addFooter(s);
}

// ═══════════════════════════════════════════════════════════════
// SLIDE 17 · Closing — Next Horizon
// ═══════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDark };

  // Left green accent bar
  s.addShape("rect", {
    x: 0, y: 0, w: 0.1, h: 5.625,
    fill: { color: C.green }, line: { color: C.green },
  });

  addLabel(s, "NEXT HORIZON", 0.35, 0.28);
  s.addText([
    { text: "We Are Just\n", options: { color: C.white } },
    { text: "Getting Started.", options: { color: C.green } },
  ], {
    x: 0.35, y: 0.62, w: 5.8, h: 1.35,
    fontSize: 44, fontFace: "Georgia", bold: true, margin: 0,
  });

  s.addText(
    "We crossed our first horizon. Our goal is to scale to $2M in total donation volume and 150+ organizations — and your support is a direct lever for that growth.",
    {
      x: 0.35, y: 2.1, w: 5.7, h: 0.72,
      fontSize: 14, fontFace: "Calibri", color: C.grayMid, margin: 0,
    }
  );

  // Opex card
  s.addShape("rect", {
    x: 0.35, y: 2.98, w: 5.35, h: 1.95,
    fill: { color: C.bgCard }, line: { color: C.green, width: 1 },
    shadow: mkShadow(),
  });
  s.addShape("rect", {
    x: 0.35, y: 2.98, w: 5.35, h: 0.07,
    fill: { color: C.green }, line: { color: C.green },
  });
  s.addText("The Opex Gap to Scale", {
    x: 0.55, y: 3.12, w: 5.0, h: 0.38,
    fontSize: 15, fontFace: "Calibri", bold: true, color: C.green, margin: 0,
  });
  s.addText(
    "KindKart earns ~25% product margin from operations. To scale from $1M → $2M in donations, we need OpEx funding for engineering, fulfillment, and growth infrastructure.",
    {
      x: 0.55, y: 3.52, w: 5.0, h: 0.75,
      fontSize: 12, fontFace: "Calibri", color: C.offWhite, margin: 0,
    }
  );
  s.addText("Seeking $100K in OpEx support for FY 2026", {
    x: 0.55, y: 4.32, w: 5.0, h: 0.38,
    fontSize: 14, fontFace: "Calibri", bold: true, color: C.gold, margin: 0,
  });

  // Right: Your Impact card
  s.addShape("rect", {
    x: 6.05, y: 1.3, w: 3.62, h: 3.72,
    fill: { color: C.bgCard }, line: { color: C.orange, width: 1 },
    shadow: mkShadow(),
  });
  s.addShape("rect", {
    x: 6.05, y: 1.3, w: 3.62, h: 0.08,
    fill: { color: C.orange }, line: { color: C.orange },
  });
  s.addText("Your Impact\nLast Year", {
    x: 6.22, y: 1.45, w: 3.28, h: 0.82,
    fontSize: 20, fontFace: "Georgia", bold: true, color: C.white, margin: 0,
  });
  s.addText("$100K → $1M+\nin donation volume reached", {
    x: 6.22, y: 2.32, w: 3.28, h: 0.72,
    fontSize: 15, fontFace: "Calibri", color: C.offWhite, margin: 0,
  });

  // Divider
  s.addShape("rect", {
    x: 6.22, y: 3.15, w: 3.08, h: 0.03,
    fill: { color: C.bgCard2 }, line: { color: C.bgCard2 },
  });

  s.addText("Join Us Again", {
    x: 6.22, y: 3.28, w: 3.28, h: 0.52,
    fontSize: 20, fontFace: "Georgia", bold: true, color: C.orange, margin: 0,
  });
  s.addText(
    "Help us reach our Second Horizon. Your support directly funds the infrastructure enabling transparent giving for thousands of grassroots organizations across India.",
    {
      x: 6.22, y: 3.84, w: 3.28, h: 0.9,
      fontSize: 12, fontFace: "Calibri", color: C.offWhite, margin: 0,
    }
  );
  s.addText("sri@kindkart.com", {
    x: 6.22, y: 4.68, w: 3.28, h: 0.28,
    fontSize: 12, fontFace: "Calibri", bold: true, color: C.orange, margin: 0,
  });

  // Footer (custom for last slide)
  s.addShape("rect", {
    x: 0, y: 5.25, w: 10, h: 0.375,
    fill: { color: C.bgCard }, line: { color: C.bgCard },
  });
  s.addText("kindkart  ·  sri@kindkart.com  ·  April 2026", {
    x: 0.35, y: 5.27, w: 9.3, h: 0.33,
    fontSize: 10, fontFace: "Calibri", color: C.gray, margin: 0,
  });
  s.addText("kindkart  ·  Sponsor Update  ·  April 2026", {
    x: 0.35, y: 5.27, w: 9.3, h: 0.33,
    fontSize: 10, fontFace: "Calibri", color: C.gray, align: "right", margin: 0,
  });
}

// ─── Save ──────────────────────────────────────────────────────
pres
  .writeFile({ fileName: "KindKart-Radhika-Update-April2026.pptx" })
  .then(() => console.log("✓  Saved: KindKart-Radhika-Update-April2026.pptx"))
  .catch((err) => { console.error("Error:", err); process.exit(1); });
