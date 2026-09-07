const pptxgen = require("pptxgenjs");
const fs = require("fs");

// ---------- palette (lifted from the Mattress Warehouse rebuild) ----------
const NAVY = "000946";        // header band, tiles, table header, callout
const PERI = "9999FF";        // titles, section labels, emphasis
const BODY = "1A1A2E";        // body copy
const TILE_LABEL = "C9CBE8";  // small label under a tile value
const CARD = "F3F4FB";        // light card fill
const CARD_DEEP = "E9EBF7";   // light callout band
const RULE = "EBEBEB";        // table borders
const WHITE = "FFFFFF";

const FONT = "Arial";

const pres = new pptxgen();
pres.defineLayout({ name: "PREMION_WIDE", width: 13.333, height: 7.5 });
pres.layout = "PREMION_WIDE";

const W = 13.333, H = 7.5;
const M = 0.5;                   // outer margin (MW uses 0.5)
const CONTENT_W = W - 2 * M;
const BAND_H = 1.9;              // MW header band is 25.4% of slide height
const CONTENT_TOP = BAND_H + 0.2;

// ---------- assets ----------
const img = (f) => "image/png;base64," + fs.readFileSync(`/home/claude/assets/${f}`).toString("base64");
const HATCH_TRI = img("hatch_tri.png");     // 748x1506 diagonal-hatch triangle, right edge of band
const HATCH_L = img("hatch_L.png");         // 686x686 periwinkle hatch "L", cover motif
const PREMION_WHITE = img("premion_white.png"); // 526x88
const PREMION_BLACK = img("premion_black.png");

const cardShadow = () => ({ type: "outer", blur: 6, offset: 2, angle: 90, color: "000946", opacity: 0.14 });

// ---------- shared pieces ----------
// Navy header band + periwinkle title + white subtitle + hatch triangle at the right edge.
function frame(slide, title, subtitle) {
  slide.background = { color: WHITE };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: BAND_H, fill: { color: NAVY }, line: { color: NAVY } });
  const triW = BAND_H * (748 / 1506);
  slide.addImage({ data: HATCH_TRI, x: W - triW - 0.03, y: 0.02, w: triW, h: BAND_H - 0.02 });
  slide.addText(title, {
    x: M, y: 0.38, w: CONTENT_W - 1.3, h: 0.62,
    fontFace: FONT, fontSize: 28, bold: true, color: PERI,
    isTextBox: true, margin: 0, valign: "middle",
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: M, y: 1.02, w: CONTENT_W - 1.3, h: 0.42,
      fontFace: FONT, fontSize: 15, color: WHITE,
      isTextBox: true, margin: 0, valign: "top",
    });
  }
  footer(slide);
}

// Bottom-left logo lockup: TEGNA wordmark (custGeom, injected by inject_logos.py) + PREMION png.
function footer(slide) {
  slide.addImage({ data: PREMION_BLACK, x: M + 1.05, y: H - 0.43, w: 0.177 * (526 / 88), h: 0.177 });
}

function kpiTile(slide, x, y, w, h, valueToken, label, big) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.08,
  });
  slide.addText(valueToken, {
    x: x + 0.15, y: y + 0.1, w: w - 0.3, h: h * 0.58,
    fontFace: FONT, fontSize: big ? 26 : 20, bold: true, color: WHITE, align: "center",
    isTextBox: true, margin: 0, valign: "middle", fit: "shrink",
  });
  slide.addText(label, {
    x: x + 0.15, y: y + h * 0.66, w: w - 0.3, h: h * 0.28,
    fontFace: FONT, fontSize: 10, color: TILE_LABEL, align: "center",
    isTextBox: true, margin: 0, valign: "top",
  });
}

function card(slide, x, y, w, h, color) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: color || CARD }, line: { color: color || CARD }, shadow: cardShadow(),
  });
}

function region(slide, name, x, y, w, h, label) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: CARD }, line: { color: PERI, dashType: "dash", width: 1 },
    objectName: name,
  });
  slide.addText(label, {
    x, y, w, h,
    fontFace: FONT, fontSize: 11, color: PERI, align: "center", valign: "middle",
    isTextBox: true, margin: 0, objectName: name + "Label",
  });
}

// Small periwinkle all-caps label (MW: "TOP 10 PUBLISHERS BY DELIVERY")
function sectionHeader(slide, text, x, y, w, color) {
  slide.addText(text.toUpperCase(), {
    x, y, w, h: 0.26,
    fontFace: FONT, fontSize: 11, bold: true, color: color || PERI, charSpacing: 1,
    isTextBox: true, margin: 0, valign: "middle",
  });
}

function tokenTable(slide, name, x, y, w, colHeads, colW, tokenRow, leftCols) {
  leftCols = leftCols || 1;
  const head = colHeads.map((t, i) => ({
    text: t,
    options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 10, fontFace: FONT, valign: "middle", align: i < leftCols ? "left" : "center" },
  }));
  const row = tokenRow.map((t, i) => ({
    text: t,
    options: { color: BODY, fontSize: 10, fontFace: FONT, valign: "middle", align: i < leftCols ? "left" : "center", fill: { color: WHITE } },
  }));
  slide.addTable([head, row], {
    x, y, w, colW, rowH: 0.34,
    border: { type: "solid", color: RULE, pt: 0.75 },
    margin: [0.02, 0.06, 0.02, 0.06],
    objectName: name,
  });
}

// One paragraph, two runs: bold HEAD token + plain DETAIL token.
function twoRunBullets(slide, prefix, x, y, w, h, count, fontSize) {
  const runs = [];
  for (let i = 1; i <= count; i++) {
    runs.push({
      text: `{{${prefix}_HEAD_${i}}}`,
      options: { bold: true, color: NAVY, bullet: { indent: 18 }, paraSpaceAfter: 12 },
    });
    runs.push({
      text: ` {{${prefix}_DETAIL_${i}}}`,
      options: { bold: false, color: BODY, breakLine: i < count },
    });
  }
  slide.addText(runs, {
    x, y, w, h, fontFace: FONT, fontSize: fontSize || 13, valign: "top",
    isTextBox: true, margin: 0, objectName: prefix + "Bullets",
  });
}

function bulletToken(slide, token, x, y, w, h, name, color) {
  slide.addText([{ text: token, options: { bullet: { indent: 18 }, color: color || BODY } }], {
    x, y, w, h, fontFace: FONT, fontSize: 13, valign: "top",
    isTextBox: true, margin: 0, objectName: name,
  });
}

function narrative(slide, token, x, y, w, h, name, fontSize) {
  slide.addText(token, {
    x, y, w, h,
    fontFace: FONT, fontSize: fontSize || 12, color: BODY, valign: "top",
    isTextBox: true, margin: 0, objectName: name,
  });
}

// =====================================================================
// 1. Campaign Recap — report:recap   (cover treatment from the MW deck)
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  const coverH = 3.1;
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: coverH, fill: { color: NAVY }, line: { color: NAVY } });
  s.addImage({ data: HATCH_L, x: M, y: 0.4, w: 1.7, h: 1.7 });
  // white logo lockup, top right (TEGNA custGeom injected by inject_logos.py)
  s.addImage({ data: PREMION_WHITE, x: W - M - 1.26, y: 0.5, w: 0.21 * (526 / 88), h: 0.21 });

  s.addText("{{CLIENT_NAME}}", {
    x: M + 0.75, y: 0.95, w: CONTENT_W - 3.0, h: 0.9,
    fontFace: FONT, fontSize: 36, bold: true, color: WHITE,
    isTextBox: true, margin: 0, valign: "middle", fit: "shrink",
  });
  s.addText("{{REPORT_TITLE}}", {
    x: M + 0.75, y: 1.9, w: CONTENT_W - 3.0, h: 0.5,
    fontFace: FONT, fontSize: 18, color: PERI,
    isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText("CAMPAIGN RECAP", {
    x: M + 0.75, y: 2.45, w: 6, h: 0.3,
    fontFace: FONT, fontSize: 11, bold: true, color: TILE_LABEL, charSpacing: 2,
    isTextBox: true, margin: 0, valign: "middle",
  });

  // three info tiles
  const tileY = coverH + 0.3, tileH = 1.0, gap = 0.3;
  const tileW = (CONTENT_W - 2 * gap) / 3;
  kpiTile(s, M, tileY, tileW, tileH, "{{REPORT_PERIOD_LABEL}}", "Report period", false);
  kpiTile(s, M + tileW + gap, tileY, tileW, tileH, "{{FLIGHT_LABEL}}", "Campaign flight", false);
  kpiTile(s, M + 2 * (tileW + gap), tileY, tileW, tileH, "{{GEOGRAPHY_LABEL}}", "Geography", false);

  // two columns on light cards
  const colY = tileY + tileH + 0.35, colGap = 0.3, colW = (CONTENT_W - colGap) / 2;
  const cardH = H - 0.6 - colY;
  card(s, M, colY, colW, cardH);
  sectionHeader(s, "Campaign goals", M + 0.3, colY + 0.2, colW - 0.6);
  bulletToken(s, "{{GOALS_BULLETS}}", M + 0.3, colY + 0.6, colW - 0.6, cardH - 0.8, "GoalsBullets");
  card(s, M + colW + colGap, colY, colW, cardH);
  sectionHeader(s, "Audience", M + colW + colGap + 0.3, colY + 0.2, colW - 0.6);
  bulletToken(s, "{{AUDIENCE_BULLETS}}", M + colW + colGap + 0.3, colY + 0.6, colW - 0.6, cardH - 0.8, "AudienceBullets");

  s.addNotes("key: report:recap");
}

// =====================================================================
// 2. Top Highlights — report:highlights
// =====================================================================
{
  const s = pres.addSlide();
  frame(s, "Top Highlights");
  const tileY = CONTENT_TOP, tileH = 1.25, gap = 0.3;
  const tileW = (CONTENT_W - 2 * gap) / 3;
  kpiTile(s, M, tileY, tileW, tileH, "{{HEADLINE_IMPRESSIONS}}", "Impressions delivered", true);
  kpiTile(s, M + tileW + gap, tileY, tileW, tileH, "{{HEADLINE_UNIQUE_VISITORS}}", "Attributed unique website visitors", true);
  kpiTile(s, M + 2 * (tileW + gap), tileY, tileW, tileH, "{{HEADLINE_ATTRIBUTED_RATE}}", "Attributed rate", true);

  const cardY = tileY + tileH + 0.35, cardH = H - 0.6 - cardY;
  card(s, M, cardY, CONTENT_W, cardH);
  sectionHeader(s, "What stood out", M + 0.3, cardY + 0.2, CONTENT_W - 0.6);
  twoRunBullets(s, "HIGHLIGHT", M + 0.3, cardY + 0.6, CONTENT_W - 0.6, cardH - 0.8, 4);
  s.addNotes("key: report:highlights");
}

// =====================================================================
// 3. Delivery Recap — report:delivery_recap  (DELIVERY SET)
// =====================================================================
{
  const s = pres.addSlide();
  frame(s, "Delivery Recap", "Streaming delivery at a glance");
  const tileY = CONTENT_TOP, tileH = 0.95, gap = 0.25;
  const tileW = (CONTENT_W - 3 * gap) / 4;
  kpiTile(s, M, tileY, tileW, tileH, "{{DELIVERED_IMPRESSIONS}}", "Impressions delivered", false);
  kpiTile(s, M + (tileW + gap), tileY, tileW, tileH, "{{VCR}}", "Video completion rate", false);
  kpiTile(s, M + 2 * (tileW + gap), tileY, tileW, tileH, "{{FREQUENCY}}", "Average frequency", false);
  kpiTile(s, M + 3 * (tileW + gap), tileY, tileW, tileH, "{{UNIQUES}}", "Unique households reached", false);

  const colY = tileY + tileH + 0.35, colGap = 0.4, colW = (CONTENT_W - colGap) / 2;
  const rx = M + colW + colGap;

  sectionHeader(s, "Top publishers", M, colY, colW);
  tokenTable(s, "TopPublishersTable", M, colY + 0.35, colW,
    ["Channel", "Impressions", "VCR"], [colW * 0.5, colW * 0.28, colW * 0.22],
    ["{{TOP_PUBLISHERS_ROWS}}", "", ""]);

  sectionHeader(s, "By creative", M, colY + 1.65, colW);
  tokenTable(s, "CreativeTable", M, colY + 2.0, colW,
    ["Creative", "Impressions", "VCR"], [colW * 0.5, colW * 0.28, colW * 0.22],
    ["{{CREATIVE_ROWS}}", "", ""]);

  sectionHeader(s, "Delivery by publisher", rx, colY, colW);
  region(s, "ChartRegion", rx, colY + 0.35, colW, 2.55, "ChartRegion — publisher delivery");
  narrative(s, "{{DELIVERY_NARRATIVE}}", rx, colY + 3.1, colW, 0.75, "DeliveryNarrative");
  s.addNotes("key: report:delivery_recap\ndelivery_set: true");
}

// =====================================================================
// 4. Website Attribution Breakdown — report:attribution_breakdown
// =====================================================================
{
  const s = pres.addSlide();
  frame(s, "Website Attribution", "{{ATTRIBUTION_HEADLINE_NOTE}}");
  const colY = CONTENT_TOP, colGap = 0.4, colW = (CONTENT_W - colGap) / 2;
  const rx = M + colW + colGap;

  sectionHeader(s, "Attributed rate by dimension", M, colY, colW);
  tokenTable(s, "BreakdownTable", M, colY + 0.35, colW,
    ["{{BREAKDOWN_DIMENSION_LABEL}}", "Delivered", "Attributed", "Rate"],
    [colW * 0.4, colW * 0.22, colW * 0.22, colW * 0.16],
    ["{{BREAKDOWN_ROWS}}", "", "", ""]);

  sectionHeader(s, "How the dimensions compare", rx, colY, colW);
  region(s, "ChartRegion", rx, colY + 0.35, colW, 3.0, "ChartRegion — attributed rate by dimension");

  const bandY = colY + 3.65, bandH = H - 0.6 - bandY;
  card(s, M, bandY, CONTENT_W, bandH, CARD_DEEP);
  narrative(s, "{{ATTRIBUTION_NARRATIVE}}", M + 0.3, bandY + 0.2, CONTENT_W - 0.6, bandH - 0.4, "AttributionNarrative", 13);
  s.addNotes("key: report:attribution_breakdown");
}

// =====================================================================
// 5. URL Report — report:url_report
// =====================================================================
{
  const s = pres.addSlide();
  frame(s, "Where Visitors Went", "{{URL_HEADLINE_NOTE}}");
  const colY = CONTENT_TOP, colGap = 0.4;
  const leftW = 7.2, rightW = CONTENT_W - leftW - colGap;
  const rx = M + leftW + colGap;

  sectionHeader(s, "Top pages by attributed visits", M, colY, leftW);
  tokenTable(s, "TopUrlTable", M, colY + 0.35, leftW,
    ["Page / section", "Visits", "Share"], [leftW * 0.6, leftW * 0.2, leftW * 0.2],
    ["{{TOP_URL_ROWS}}", "", ""]);

  const cardH = H - 0.6 - colY;
  card(s, rx, colY, rightW, cardH);
  sectionHeader(s, "What this signals", rx + 0.3, colY + 0.2, rightW - 0.6);
  narrative(s, "{{URL_INTENT_NARRATIVE}}", rx + 0.3, colY + 0.6, rightW - 0.6, cardH - 0.8, "UrlIntentNarrative", 13);
  s.addNotes("key: report:url_report");
}

// =====================================================================
// 6. Zip-code analysis — report:zip_analysis
// =====================================================================
{
  const s = pres.addSlide();
  frame(s, "Where Visitors Came From", "{{ZIP_HEADLINE_NOTE}}");
  const colY = CONTENT_TOP, colGap = 0.4;
  const mapW = 6.6, rightW = CONTENT_W - mapW - colGap;
  const rx = M + mapW + colGap;

  sectionHeader(s, "Attributed impressions — heat map", M, colY, mapW);
  region(s, "MapRegion", M, colY + 0.35, mapW, 3.45, "MapRegion — visitor zips (+ targeted zips when a proposal is linked)");

  sectionHeader(s, "Top zip codes", rx, colY, rightW);
  tokenTable(s, "TopZipTable", rx, colY + 0.35, rightW,
    ["Zip", "Area", "Rate vs. avg"], [rightW * 0.22, rightW * 0.46, rightW * 0.32],
    ["{{TOP_ZIP_ROWS}}", "", ""], 2);

  narrative(s, "{{ZIP_NARRATIVE}}", M, colY + 4.0, CONTENT_W, 0.8, "ZipNarrative");
  s.addNotes("key: report:zip_analysis");
}

// =====================================================================
// 7. Takeaways / What's next — report:takeaways
// =====================================================================
{
  const s = pres.addSlide();
  frame(s, "Takeaways & What's Next");
  const colY = CONTENT_TOP, colGap = 0.4;
  const leftW = 7.0, rightW = CONTENT_W - leftW - colGap;
  const rx = M + leftW + colGap;
  const panelH = H - 0.6 - colY;

  card(s, M, colY, leftW, panelH);
  sectionHeader(s, "Key takeaways", M + 0.3, colY + 0.2, leftW - 0.6);
  twoRunBullets(s, "TAKEAWAY", M + 0.3, colY + 0.6, leftW - 0.6, panelH - 0.8, 4);

  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: rx, y: colY, w: rightW, h: panelH, fill: { color: NAVY }, line: { color: NAVY }, rectRadius: 0.08,
  });
  sectionHeader(s, "What's next", rx + 0.3, colY + 0.2, rightW - 0.6, PERI);
  bulletToken(s, "{{WHATS_NEXT_BULLETS}}", rx + 0.3, colY + 0.6, rightW - 0.6, panelH - 0.8, "WhatsNextBullets", WHITE);
  s.addNotes("key: report:takeaways");
}

pres.writeFile({ fileName: "/home/claude/REPORT_MASTER_v0_2.pptx" }).then((f) => console.log("wrote", f));
