// Uplift Modeling deck: S/T/X-learner comparison + full X-learner walkthrough.
const pptxgen = require("pptxgenjs");

const NAVY = "1E2761";
const NAVY_DK = "141B4D";
const ICE = "CADCFC";
const ICE_LT = "EAF0FC";
const WHITE = "FFFFFF";
const BODY = "33344A";
const MUTED = "6B7280";
const GREEN = "1B7A5A";
const GREEN_BG = "E4F3EC";
const RED = "B03A2E";
const RED_BG = "FBEAE8";
const GOLD = "B8892B";
const GOLD_BG = "FBF3E1";

const HEAD_FONT = "Cambria";
const BODY_FONT = "Calibri";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
const PW = 13.333, PH = 7.5;

function newSlide(bg) {
  const s = pres.addSlide();
  s.background = { color: bg || WHITE };
  return s;
}

// Section header for a light-background content slide.
function header(s, kicker, title, opts = {}) {
  s.addText(kicker.toUpperCase(), {
    x: 0.6, y: 0.35, w: 10, h: 0.3, fontFace: BODY_FONT, fontSize: 12,
    color: GOLD, bold: true, charSpacing: 2, isTextBox: true, margin: 0,
  });
  s.addText(title, {
    x: 0.6, y: 0.62, w: opts.titleW || 11.8, h: opts.titleH || 0.7,
    fontFace: HEAD_FONT, fontSize: opts.size || 30, bold: true, color: NAVY,
    isTextBox: true, margin: 0,
  });
}

function pageNum(s, n) {
  s.addText(String(n), {
    x: PW - 0.9, y: PH - 0.5, w: 0.5, h: 0.3, fontFace: BODY_FONT, fontSize: 10,
    color: MUTED, align: "right", isTextBox: true, margin: 0,
  });
  s.addText("Uplift Modeling — X-Learner Deep Dive", {
    x: 0.6, y: PH - 0.5, w: 6, h: 0.3, fontFace: BODY_FONT, fontSize: 9,
    color: MUTED, isTextBox: true, margin: 0,
  });
}

function pill(s, text, x, y, w, h, fill, textColor) {
  s.addShape("roundRect", { x, y, w, h, rectRadius: h / 2, fill: { color: fill }, line: { type: "none" } });
  s.addText(text, {
    x, y, w, h, align: "center", valign: "middle", fontFace: BODY_FONT, fontSize: 12,
    bold: true, color: textColor, isTextBox: true, margin: 0,
  });
}

function formulaBox(s, text, x, y, w, h, opts = {}) {
  s.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.08,
    fill: { color: opts.fill || NAVY_DK }, line: { type: "none" },
  });
  s.addText(text, {
    x: x + 0.25, y, w: w - 0.5, h, valign: "middle", fontFace: "Courier New", fontSize: opts.size || 16,
    bold: true, color: opts.color || WHITE, isTextBox: true, margin: 0, align: opts.align || "left",
  });
}

// ---------------------------------------------------------------------------
// Slide 1 — Title
// ---------------------------------------------------------------------------
{
  const s = newSlide(NAVY);
  s.addShape("rect", { x: 0, y: 0, w: PW, h: PH, fill: { color: NAVY }, line: { type: "none" } });
  // decorative concentric arcs, bottom-right, subtle
  s.addShape("ellipse", { x: 9.6, y: 3.6, w: 6, h: 6, fill: { type: "none" }, line: { color: "34408A", width: 1.5 } });
  s.addShape("ellipse", { x: 10.3, y: 4.3, w: 4.6, h: 4.6, fill: { type: "none" }, line: { color: "34408A", width: 1.5 } });
  s.addShape("ellipse", { x: 11, y: 5, w: 3.2, h: 3.2, fill: { color: GOLD }, line: { type: "none" }, shadow: undefined });

  s.addText("UPLIFT MODELING", {
    x: 0.9, y: 2.15, w: 10, h: 0.4, fontFace: BODY_FONT, fontSize: 14, color: ICE,
    bold: true, charSpacing: 3, isTextBox: true, margin: 0,
  });
  s.addText("Targeting the Persuadables:\nThe X-Learner Approach", {
    x: 0.85, y: 2.55, w: 9.6, h: 2.0, fontFace: HEAD_FONT, fontSize: 44, bold: true, color: WHITE,
    isTextBox: true, margin: 0, lineSpacingMultiple: 1.05,
  });
  s.addText("From two outcome models to a single ranked uplift score — how it's built, why X-Learner over S/T-Learner, and how it's validated.", {
    x: 0.9, y: 4.65, w: 8.8, h: 0.8, fontFace: BODY_FONT, fontSize: 15, color: ICE, isTextBox: true, margin: 0,
  });
  s.addText("Credit Card Prescreen Targeting  ·  Methodology Walkthrough  ·  Synthetic-Data Demo", {
    x: 0.9, y: 6.7, w: 9, h: 0.4, fontFace: BODY_FONT, fontSize: 11.5, color: "8E9BD6", isTextBox: true, margin: 0,
  });
}

// ---------------------------------------------------------------------------
// Slide 2 — Agenda
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "Roadmap", "What We'll Walk Through");
  const items = [
    ["01", "Why an uplift model", "Response models waste budget on people who'd act anyway — uplift targets who the offer actually changes."],
    ["02", "Three ways to estimate it", "S-Learner, T-Learner, X-Learner — what each one does, and why X-Learner wins here."],
    ["03", "X-Learner, layer by layer", "Two outcome models → pseudo effects → two effect models → weighted fusion, with a worked numeric example."],
    ["04", "Validation & business impact", "Qini curve, a ground-truth check only synthetic data allows, and what it means for targeting decisions."],
  ];
  const top = 1.75, rh = 1.22, gap = 0.18;
  items.forEach((it, i) => {
    const y = top + i * (rh + gap);
    s.addShape("roundRect", { x: 0.6, y, w: 12.1, h: rh, rectRadius: 0.09, fill: { color: i % 2 === 0 ? ICE_LT : WHITE }, line: { color: ICE, width: 1 } });
    s.addText(it[0], {
      x: 0.9, y, w: 1.1, h: rh, valign: "middle", fontFace: HEAD_FONT, fontSize: 30, bold: true, color: ICE.length ? "9FB4E8" : NAVY,
      isTextBox: true, margin: 0,
    });
    s.addText(it[1], {
      x: 2.15, y: y + 0.14, w: 4.6, h: 0.5, fontFace: HEAD_FONT, fontSize: 17, bold: true, color: NAVY, isTextBox: true, margin: 0,
    });
    s.addText(it[2], {
      x: 2.15, y: y + 0.58, w: 9.9, h: 0.55, fontFace: BODY_FONT, fontSize: 12.5, color: BODY, isTextBox: true, margin: 0,
    });
  });
  pageNum(s, 2);
}

// ---------------------------------------------------------------------------
// Slide 3 — Why uplift, not response (four quadrants)
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "The Core Idea", "Not “Who Responds” — “Who Responds Because of Us”");
  s.addText(
    "A response model predicts who's likely to respond — but that includes people who'd apply on the existing offer anyway. An uplift model predicts the incremental effect of switching someone to the new offer, so targeting can focus on the one segment where the new offer actually changes the outcome.",
    { x: 0.6, y: 1.5, w: 5.7, h: 1.7, fontFace: BODY_FONT, fontSize: 13.5, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 }
  );
  s.addText([
    { text: "τ(x) = P(Y=1 | X=x, T=1) − P(Y=1 | X=x, T=0)", options: {} },
  ], {
    x: 0.6, y: 3.35, w: 5.7, h: 0.6, fontFace: "Courier New", fontSize: 14, bold: true, color: NAVY,
    isTextBox: true, margin: 0,
  });
  s.addShape("roundRect", { x: 0.6, y: 4.0, w: 5.7, h: 0.85, rectRadius: 0.06, fill: { color: ICE_LT }, line: { type: "none" } });
  s.addText([
    { text: "In our prescreen example:  ", options: { bold: true, color: NAVY } },
    { text: "Y (our Applied column) = 1 if the customer applies for the credit card, 0 if not.  T (our new_offer column) = 1 if they received the new offer variant (treatment), 0 if they received the existing offer (control) — everyone gets an offer, this is old vs. new.", options: { color: BODY } },
  ], { x: 0.85, y: 4.0, w: 5.35, h: 0.85, valign: "middle", fontFace: BODY_FONT, fontSize: 11, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  s.addText("τ(x) — our uplift score — is the treatment effect we're trying to estimate, not the raw response probability.", {
    x: 0.6, y: 4.95, w: 5.7, h: 0.45, fontFace: BODY_FONT, fontSize: 11, italic: true, color: MUTED, isTextBox: true, margin: 0,
  });

  // 2x2 quadrant grid
  const gx = 6.7, gy = 1.5, gw = 5.9, gh = 5.35;
  const cw = gw / 2, ch = gh / 2, pad = 0.12;
  const quads = [
    { t: "Persuadables", d: "Apply only with the new offer (Y=1 under T=1, Y=0 under T=0) — switching the offer changes their decision. This is the target segment.", fill: GREEN_BG, tc: GREEN },
    { t: "Sure Things", d: "Apply on either offer (Y=1 regardless of T). Giving them the new offer is wasted — the old one would've worked too.", fill: ICE_LT, tc: NAVY },
    { t: "Lost Causes", d: "Never apply, old or new offer (Y=0 regardless of T). Nothing to gain.", fill: "F2F2F5", tc: MUTED },
    { t: "Sleeping Dogs", d: "Apply LESS often with the new offer than they would have with the old one — the new offer backfires.", fill: RED_BG, tc: RED },
  ];
  quads.forEach((q, i) => {
    const cx = gx + (i % 2) * (cw + pad);
    const cy = gy + Math.floor(i / 2) * (ch + pad);
    s.addShape("roundRect", { x: cx, y: cy, w: cw - pad, h: ch - pad, rectRadius: 0.08, fill: { color: q.fill }, line: { type: "none" } });
    s.addText(q.t, { x: cx + 0.22, y: cy + 0.18, w: cw - pad - 0.4, h: 0.4, fontFace: HEAD_FONT, fontSize: 15.5, bold: true, color: q.tc, isTextBox: true, margin: 0 });
    s.addText(q.d, { x: cx + 0.22, y: cy + 0.62, w: cw - pad - 0.44, h: ch - pad - 0.8, fontFace: BODY_FONT, fontSize: 11, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  });
  pageNum(s, 3);
}

// ---------------------------------------------------------------------------
// Slide 4 — Three estimators overview
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "Estimation Strategies", "Three Ways to Estimate an Uplift Score");
  const cards = [
    { t: "S-Learner", sub: "Single model", d: "One model, new_offer (old vs. new offer, 1/0) as just another input feature. τ(x) = f(x,1) − f(x,0) — our uplift score — where 1/0 flips which offer they got.", fill: ICE_LT, badge: null },
    { t: "T-Learner", sub: "Two models", d: "Two fully separate outcome models, one per arm. τ(x) = μ̂₁(x) − μ̂₀(x), where μ̂₁ is the new-offer outcome model and μ̂₀ is the old-offer (control) outcome model.", fill: ICE_LT, badge: null },
    { t: "X-Learner", sub: "Cross-learning", d: "Two outcome models feed two effect models, combined by a weight g(x). Same Applied and new_offer columns as above.", fill: GOLD_BG, badge: "Used here" },
  ];
  const cw = 3.85, gap = 0.3, top = 1.85, ch = 4.9;
  const startX = (PW - (cw * 3 + gap * 2)) / 2;
  cards.forEach((c, i) => {
    const x = startX + i * (cw + gap);
    s.addShape("roundRect", { x, y: top, w: cw, h: ch, rectRadius: 0.1, fill: { color: c.fill }, line: { color: c.badge ? GOLD : ICE, width: c.badge ? 1.5 : 1 } });
    if (c.badge) {
      pill(s, c.badge, x + cw - 1.7, top + 0.22, 1.45, 0.34, GOLD, WHITE);
    }
    s.addText(String(i + 1), { x: x + 0.3, y: top + 0.3, w: 1, h: 0.6, fontFace: HEAD_FONT, fontSize: 26, bold: true, color: NAVY, isTextBox: true, margin: 0 });
    s.addText(c.t, { x: x + 0.3, y: top + 0.95, w: cw - 0.6, h: 0.5, fontFace: HEAD_FONT, fontSize: 21, bold: true, color: NAVY, isTextBox: true, margin: 0 });
    s.addText(c.sub.toUpperCase(), { x: x + 0.3, y: top + 1.42, w: cw - 0.6, h: 0.3, fontFace: BODY_FONT, fontSize: 10.5, bold: true, color: MUTED, charSpacing: 1.5, isTextBox: true, margin: 0 });
    s.addShape("line", { x: x + 0.3, y: top + 1.85, w: cw - 0.6, h: 0, line: { color: c.badge ? GOLD : ICE, width: 1 } });
    s.addText(c.d, { x: x + 0.3, y: top + 2.05, w: cw - 0.6, h: 1.6, fontFace: BODY_FONT, fontSize: 13, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25 });
  });
  pageNum(s, 4);
}

// ---------------------------------------------------------------------------
// Slide 5 — S-Learner detail: train / score / pros-cons / why not
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "Estimator 1 of 3", "S-Learner, In Our Prescreen Example");

  const cw = 5.85, x1 = 0.6, x2 = 6.85, top = 1.5, ch = 2.3;

  // TRAIN — one pooled dataset, one fit() call
  s.addShape("roundRect", { x: x1, y: top, w: cw, h: ch, rectRadius: 0.1, fill: { color: ICE_LT }, line: { color: NAVY, width: 1 } });
  s.addText([
    { text: "TRAIN   ", options: { bold: true, color: NAVY } },
    { text: "pooled — ONE model, ONE fit() call", options: { color: MUTED, italic: true } },
  ], { x: x1 + 0.25, y: top + 0.14, w: cw - 0.5, h: 0.28, fontFace: BODY_FONT, fontSize: 11, isTextBox: true, margin: 0 });
  {
    const tw = cw - 0.5, ty = top + 0.5, th = ch - 0.65;
    const head = ["Customer", "new_offer", "Applied"].map(t => ({ text: t, options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "center", valign: "middle" } }));
    const data = [["A", "0", "0"], ["B", "0", "1"], ["C", "0", "0"], ["D", "1", "1"], ["E", "1", "0"], ["F", "1", "1"]];
    const rows = [head, ...data.map((r, i) => r.map((v, j) => ({
      text: v,
      options: { align: "center", valign: "middle", bold: j === 0, color: j === 1 ? (v === "0" ? NAVY : GOLD) : BODY, fill: { color: i % 2 === 0 ? WHITE : "F4F6FB" } },
    })))];
    s.addTable(rows, {
      x: x1 + 0.25, y: ty, w: tw, h: th, colW: [tw * 0.34, tw * 0.34, tw * 0.32],
      fontFace: BODY_FONT, fontSize: 10.5, border: { type: "solid", color: "E4E7F0", pt: 1 },
      autoPage: false, margin: [0.02, 0.05, 0.02, 0.05], rowH: th / 7,
    });
  }

  // SCORE — same model, called twice for one customer
  s.addShape("roundRect", { x: x2, y: top, w: cw, h: ch, rectRadius: 0.1, fill: { color: GOLD_BG }, line: { color: GOLD, width: 1 } });
  s.addText([
    { text: "SCORE   ", options: { bold: true, color: GOLD } },
    { text: "one model, run for BOTH offer scenarios on customer G", options: { color: MUTED, italic: true } },
  ], { x: x2 + 0.25, y: top + 0.14, w: cw - 0.5, h: 0.28, fontFace: BODY_FONT, fontSize: 11, isTextBox: true, margin: 0 });
  {
    const ty = top + 0.5;
    const inputW = 2.5, arrowW = 0.4, outputW = 1.7;
    const inX = x2 + 0.25, arrowX = inX + inputW, outX = arrowX + arrowW;
    const pairH = 0.46, gapPair = 0.13, gapFinal = 0.12;
    const row2Y = ty + pairH + gapPair;
    const finalY = row2Y + pairH + gapFinal;
    const finalH = (top + ch - 0.15) - finalY;

    function scenarioPair(y, offerVal, offerColor, rowBg, output) {
      const inHead = ["Customer", "new_offer"].map(t => ({ text: t, options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "center", valign: "middle", fontSize: 8.5 } }));
      const inRow = [
        { text: "G", options: { align: "center", valign: "middle", bold: true, color: NAVY, fill: { color: rowBg } } },
        { text: String(offerVal), options: { align: "center", valign: "middle", bold: true, color: offerColor, fill: { color: rowBg } } },
      ];
      s.addTable([inHead, inRow], {
        x: inX, y, w: inputW, h: pairH, colW: [inputW * 0.45, inputW * 0.55],
        fontFace: BODY_FONT, fontSize: 11, border: { type: "solid", color: "F0DFB0", pt: 1 },
        autoPage: false, margin: [0.02, 0.04, 0.02, 0.04], rowH: pairH / 2,
      });
      s.addText("→", { x: arrowX, y, w: arrowW, h: pairH, align: "center", valign: "middle", fontFace: BODY_FONT, fontSize: 16, bold: true, color: MUTED, isTextBox: true, margin: 0 });
      const outHead = [{ text: "Applied", options: { fill: { color: GOLD }, color: WHITE, bold: true, align: "center", valign: "middle", fontSize: 8.5 } }];
      const outRow = [{ text: output, options: { align: "center", valign: "middle", bold: true, color: NAVY_DK, fill: { color: rowBg } } }];
      s.addTable([outHead, outRow], {
        x: outX, y, w: outputW, h: pairH, colW: [outputW],
        fontFace: BODY_FONT, fontSize: 12.5, border: { type: "solid", color: "F0DFB0", pt: 1 },
        autoPage: false, margin: [0.02, 0.04, 0.02, 0.04], rowH: pairH / 2,
      });
    }
    scenarioPair(ty, "1", GOLD, WHITE, "0.42");
    scenarioPair(row2Y, "0", NAVY, "FBF3E1", "0.31");

    s.addShape("roundRect", { x: x2 + 0.25, y: finalY, w: cw - 0.5, h: finalH, rectRadius: 0.06, fill: { color: NAVY_DK }, line: { type: "none" } });
    s.addText([
      { text: "τ(G) = f(G,1) − f(G,0) = 0.42 − 0.31 = ", options: { color: WHITE } },
      { text: "0.11", options: { bold: true, color: GOLD, fontSize: 14 } },
    ], { x: x2 + 0.45, y: finalY, w: cw - 0.9, h: finalH, valign: "middle", fontFace: BODY_FONT, fontSize: 11.5, bold: true, isTextBox: true, margin: 0 });
  }

  const pc_top = top + ch + 0.15, pc_h = 1.35;
  s.addShape("roundRect", { x: x1, y: pc_top, w: cw, h: pc_h, rectRadius: 0.1, fill: { color: GREEN_BG }, line: { type: "none" } });
  s.addText("+  Pros", { x: x1 + 0.3, y: pc_top + 0.14, w: cw - 0.6, h: 0.32, fontFace: HEAD_FONT, fontSize: 14.5, bold: true, color: GREEN, isTextBox: true, margin: 0 });
  s.addText("One model to build and monitor. Pools all customers, so it's data-efficient even with a small treatment group.", {
    x: x1 + 0.3, y: pc_top + 0.5, w: cw - 0.6, h: pc_h - 0.65, fontFace: BODY_FONT, fontSize: 11.5, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });

  s.addShape("roundRect", { x: x2, y: pc_top, w: cw, h: pc_h, rectRadius: 0.1, fill: { color: RED_BG }, line: { type: "none" } });
  s.addText("−  Cons", { x: x2 + 0.3, y: pc_top + 0.14, w: cw - 0.6, h: 0.32, fontFace: HEAD_FONT, fontSize: 14.5, bold: true, color: RED, isTextBox: true, margin: 0 });
  s.addText("Regularization has no reason to favor one new_offer feature over dozens of stronger predictors — the offer signal can get nearly ignored.", {
    x: x2 + 0.3, y: pc_top + 0.5, w: cw - 0.6, h: pc_h - 0.65, fontFace: BODY_FONT, fontSize: 11.5, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });

  const why_top = pc_top + pc_h + 0.15, why_h = 1.15;
  s.addShape("roundRect", { x: 0.6, y: why_top, w: 12.1, h: why_h, rectRadius: 0.08, fill: { color: NAVY_DK }, line: { type: "none" } });
  s.addText([
    { text: "Why not chosen:  ", options: { bold: true, color: GOLD } },
    { text: "baseline behavioral features — utilization, tenure, balances — predict response far better than which offer they got. A pooled model leans on those and treats new_offer as noise, exactly the risk we can't take when isolating the new offer's effect is the whole point.", options: { color: WHITE } },
  ], { x: 0.95, y: why_top, w: 11.4, h: why_h, valign: "middle", fontFace: BODY_FONT, fontSize: 12.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  pageNum(s, 5);
}

// ---------------------------------------------------------------------------
// Slide 6 — T-Learner detail: train / score / pros-cons / why not
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "Estimator 2 of 3", "T-Learner, In Our Prescreen Example");

  const cw = 5.85, x1 = 0.6, x2 = 6.85, top = 1.5, ch = 2.3;

  // TRAIN — two disjoint datasets, two fit() calls
  s.addShape("roundRect", { x: x1, y: top, w: cw, h: ch, rectRadius: 0.1, fill: { color: ICE_LT }, line: { color: NAVY, width: 1 } });
  s.addText([
    { text: "TRAIN   ", options: { bold: true, color: NAVY } },
    { text: "two separate models, TWO fit() calls", options: { color: MUTED, italic: true } },
  ], { x: x1 + 0.25, y: top + 0.14, w: cw - 0.5, h: 0.28, fontFace: BODY_FONT, fontSize: 11, isTextBox: true, margin: 0 });
  {
    const innerW = cw - 0.5, gap = 0.2, miniW = (innerW - gap) / 2;
    const mx1 = x1 + 0.25, mx2 = mx1 + miniW + gap, labelY = top + 0.5, tableY = top + 0.74, tableH = ch - 0.89;
    s.addText("Control (A,B,C) → μ̂₀ model", { x: mx1, y: labelY, w: miniW, h: 0.24, fontFace: BODY_FONT, fontSize: 10, bold: true, color: NAVY, isTextBox: true, margin: 0 });
    const ctrlHead = ["Customer", "Applied"].map(t => ({ text: t, options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "center", valign: "middle" } }));
    const ctrlRows = [ctrlHead, ...[["A", "0"], ["B", "1"], ["C", "0"]].map((r, i) => r.map((v, j) => ({
      text: v, options: { align: "center", valign: "middle", bold: j === 0, color: BODY, fill: { color: i % 2 === 0 ? WHITE : "F4F6FB" } },
    })))];
    s.addTable(ctrlRows, { x: mx1, y: tableY, w: miniW, h: tableH, colW: [miniW * 0.5, miniW * 0.5], fontFace: BODY_FONT, fontSize: 10.5, border: { type: "solid", color: "E4E7F0", pt: 1 }, autoPage: false, margin: [0.02, 0.04, 0.02, 0.04], rowH: tableH / 4 });

    s.addText("Treatment (D,E,F) → μ̂₁ model", { x: mx2, y: labelY, w: miniW, h: 0.24, fontFace: BODY_FONT, fontSize: 10, bold: true, color: GOLD, isTextBox: true, margin: 0 });
    const trtHead = ["Customer", "Applied"].map(t => ({ text: t, options: { fill: { color: GOLD }, color: WHITE, bold: true, align: "center", valign: "middle" } }));
    const trtRows = [trtHead, ...[["D", "1"], ["E", "0"], ["F", "1"]].map((r, i) => r.map((v, j) => ({
      text: v, options: { align: "center", valign: "middle", bold: j === 0, color: BODY, fill: { color: i % 2 === 0 ? WHITE : "F4F6FB" } },
    })))];
    s.addTable(trtRows, { x: mx2, y: tableY, w: miniW, h: tableH, colW: [miniW * 0.5, miniW * 0.5], fontFace: BODY_FONT, fontSize: 10.5, border: { type: "solid", color: "E4E7F0", pt: 1 }, autoPage: false, margin: [0.02, 0.04, 0.02, 0.04], rowH: tableH / 4 });
  }

  // SCORE — both models, one predict() call each, for the SAME customer
  s.addShape("roundRect", { x: x2, y: top, w: cw, h: ch, rectRadius: 0.1, fill: { color: GOLD_BG }, line: { color: GOLD, width: 1 } });
  s.addText([
    { text: "SCORE   ", options: { bold: true, color: GOLD } },
    { text: "each model scores customer G under ITS OWN scenario", options: { color: MUTED, italic: true } },
  ], { x: x2 + 0.25, y: top + 0.14, w: cw - 0.5, h: 0.28, fontFace: BODY_FONT, fontSize: 11, isTextBox: true, margin: 0 });
  {
    const ty = top + 0.5;
    const inputW = 2.7, arrowW = 0.4, outputW = 1.7;
    const inX = x2 + 0.25, arrowX = inX + inputW, outX = arrowX + arrowW;
    const pairH = 0.46, gapPair = 0.13, gapFinal = 0.12;
    const row2Y = ty + pairH + gapPair;
    const finalY = row2Y + pairH + gapFinal;
    const finalH = (top + ch - 0.15) - finalY;

    function scenarioPair(y, modelVal, modelColor, rowBg, output) {
      const inHead = ["Customer", "model"].map(t => ({ text: t, options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "center", valign: "middle", fontSize: 8.5 } }));
      const inRow = [
        { text: "G", options: { align: "center", valign: "middle", bold: true, color: NAVY, fill: { color: rowBg } } },
        { text: modelVal, options: { align: "center", valign: "middle", bold: true, color: modelColor, fill: { color: rowBg } } },
      ];
      s.addTable([inHead, inRow], {
        x: inX, y, w: inputW, h: pairH, colW: [inputW * 0.34, inputW * 0.66],
        fontFace: BODY_FONT, fontSize: 10.5, border: { type: "solid", color: "F0DFB0", pt: 1 },
        autoPage: false, margin: [0.02, 0.04, 0.02, 0.04], rowH: pairH / 2,
      });
      s.addText("→", { x: arrowX, y, w: arrowW, h: pairH, align: "center", valign: "middle", fontFace: BODY_FONT, fontSize: 16, bold: true, color: MUTED, isTextBox: true, margin: 0 });
      const outHead = [{ text: "Applied", options: { fill: { color: GOLD }, color: WHITE, bold: true, align: "center", valign: "middle", fontSize: 8.5 } }];
      const outRow = [{ text: output, options: { align: "center", valign: "middle", bold: true, color: NAVY_DK, fill: { color: rowBg } } }];
      s.addTable([outHead, outRow], {
        x: outX, y, w: outputW, h: pairH, colW: [outputW],
        fontFace: BODY_FONT, fontSize: 12.5, border: { type: "solid", color: "F0DFB0", pt: 1 },
        autoPage: false, margin: [0.02, 0.04, 0.02, 0.04], rowH: pairH / 2,
      });
    }
    scenarioPair(ty, "μ̂₀", NAVY, WHITE, "0.19");
    scenarioPair(row2Y, "μ̂₁", GOLD, "FBF3E1", "0.35");

    s.addShape("roundRect", { x: x2 + 0.25, y: finalY, w: cw - 0.5, h: finalH, rectRadius: 0.06, fill: { color: NAVY_DK }, line: { type: "none" } });
    s.addText([
      { text: "τ(G) = μ̂₁(G) − μ̂₀(G) = 0.35 − 0.19 = ", options: { color: WHITE } },
      { text: "0.16", options: { bold: true, color: GOLD, fontSize: 14 } },
    ], { x: x2 + 0.45, y: finalY, w: cw - 0.9, h: finalH, valign: "middle", fontFace: BODY_FONT, fontSize: 11.5, bold: true, isTextBox: true, margin: 0 });
  }

  const pc_top = top + ch + 0.15, pc_h = 1.35;
  s.addShape("roundRect", { x: x1, y: pc_top, w: cw, h: pc_h, rectRadius: 0.1, fill: { color: GREEN_BG }, line: { type: "none" } });
  s.addText("+  Pros", { x: x1 + 0.3, y: pc_top + 0.14, w: cw - 0.6, h: 0.32, fontFace: HEAD_FONT, fontSize: 14.5, bold: true, color: GREEN, isTextBox: true, margin: 0 });
  s.addText("Each model specializes freely — no shared regularization fighting over the treatment signal. Simplest to reason about: “model B minus model A.”", {
    x: x1 + 0.3, y: pc_top + 0.5, w: cw - 0.6, h: pc_h - 0.65, fontFace: BODY_FONT, fontSize: 11.5, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });

  s.addShape("roundRect", { x: x2, y: pc_top, w: cw, h: pc_h, rectRadius: 0.1, fill: { color: RED_BG }, line: { type: "none" } });
  s.addText("−  Cons", { x: x2 + 0.3, y: pc_top + 0.14, w: cw - 0.6, h: 0.32, fontFace: HEAD_FONT, fontSize: 14.5, bold: true, color: RED, isTextBox: true, margin: 0 });
  s.addText("Quality depends on how well each arm is estimated alone. A small arm's model is noisy, and subtracting two noisy predictions adds their variances.", {
    x: x2 + 0.3, y: pc_top + 0.5, w: cw - 0.6, h: pc_h - 0.65, fontFace: BODY_FONT, fontSize: 11.5, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });

  const why_top = pc_top + pc_h + 0.15, why_h = 1.3;
  s.addShape("roundRect", { x: 0.6, y: why_top, w: 12.1, h: why_h, rectRadius: 0.08, fill: { color: NAVY_DK }, line: { type: "none" } });
  s.addText([
    { text: "Why not chosen:  ", options: { bold: true, color: GOLD } },
    { text: "T-Learner isn't broken by design — if new_offer were roughly 50/50, μ̂₀ and μ̂₁ would train on comparable sample sizes and subtracting two similarly-precise models wouldn't cost much. ", options: { color: WHITE } },
    { text: "The problem is our specific setup: ", options: { bold: true, color: "C9D2F0" } },
    { text: "about 70% of customers stay on the existing offer (control) and only ~30% are rolled out to the new offer (treatment) — a deliberately smaller test slice to limit risk before scaling a new offer design. That imbalance starves μ̂₁, the new-offer-side model, exactly where T-Learner needs it most: τ(x)=μ̂₁(x)−μ̂₀(x) inherits nearly all of μ̂₁'s extra noise. Closing that gap without needing a bigger treatment group is the reason X-Learner exists.", options: { color: WHITE } },
  ], { x: 0.95, y: why_top, w: 11.4, h: why_h, valign: "middle", fontFace: BODY_FONT, fontSize: 12, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });

  pageNum(s, 6);
}

// ---------------------------------------------------------------------------
// Slide 7 — T-Learner deep dive: WHY imbalance breaks the subtraction
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "Estimator 2 of 3 · Mechanism", "Why Imbalance Breaks the Subtraction");

  // Notation legend
  s.addShape("roundRect", { x: 0.6, y: 1.45, w: 12.1, h: 0.82, rectRadius: 0.06, fill: { color: ICE_LT }, line: { type: "none" } });
  s.addText([
    { text: "Notation:  ", options: { bold: true, color: NAVY } },
    { text: "μ̂₀(x), μ̂₁(x) = the two outcome models — predicted P(Applied) under the old offer vs. the new offer.   τ(x) = μ̂₁(x) − μ̂₀(x) = the uplift score we're estimating.   n₀, n₁ = how many customers trained each model (sample size).   Var(·) = how noisy a model's predictions are — it shrinks as n grows.", options: { color: BODY } },
  ], { x: 0.85, y: 1.45, w: 11.6, h: 0.82, valign: "middle", fontFace: BODY_FONT, fontSize: 11, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });

  // Formula: variance adds under subtraction
  formulaBox(s, "Var(τ̂) = Var(μ̂₀) + Var(μ̂₁)", 0.6, 2.47, 6.35, 0.68, { fill: NAVY_DK, size: 16 });
  s.addText("Two independently-fit models ⇒ subtracting them doesn't cancel their errors, it adds them. And a model's variance shrinks as its training sample n grows — fewer rows means a noisier model.", {
    x: 7.15, y: 2.47, w: 5.55, h: 0.68, valign: "middle", fontFace: BODY_FONT, fontSize: 11.5, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.18,
  });

  // Two comparison cards: balanced vs. our real imbalance
  const cw = 5.85, x1 = 0.6, x2 = 6.85, top = 3.4, ch = 2.0;
  s.addShape("roundRect", { x: x1, y: top, w: cw, h: ch, rectRadius: 0.1, fill: { color: GREEN_BG }, line: { color: GREEN, width: 1.3 } });
  s.addText("✓  If new_offer Were 50/50", { x: x1 + 0.3, y: top + 0.2, w: cw - 0.6, h: 0.4, fontFace: HEAD_FONT, fontSize: 15.5, bold: true, color: GREEN, isTextBox: true, margin: 0 });
  s.addText([
    { text: "n₀ ≈ n₁ ", options: { bold: true, color: NAVY } }, { text: "— both models see plenty of data.\n", options: { color: BODY } },
    { text: "Var(μ̂₀) ≈ Var(μ̂₁) ", options: { bold: true, color: NAVY } }, { text: "— both reasonably precise.\n", options: { color: BODY } },
    { text: "Var(τ̂) = small + small ", options: { bold: true, color: NAVY } }, { text: "— still small.\n", options: { color: BODY } },
    { text: "→ T-Learner's subtraction is safe.", options: { bold: true, color: GREEN } },
  ], { x: x1 + 0.3, y: top + 0.68, w: cw - 0.6, h: ch - 0.85, fontFace: BODY_FONT, fontSize: 12.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.35 });

  s.addShape("roundRect", { x: x2, y: top, w: cw, h: ch, rectRadius: 0.1, fill: { color: RED_BG }, line: { color: RED, width: 1.3 } });
  s.addText("✗  Our Real Setup — Old Offer ≈70% / New Offer ≈30%", { x: x2 + 0.3, y: top + 0.2, w: cw - 0.6, h: 0.4, fontFace: HEAD_FONT, fontSize: 14, bold: true, color: RED, isTextBox: true, margin: 0 });
  s.addText([
    { text: "n₁ (treatment) is small ", options: { bold: true, color: NAVY } }, { text: "— μ̂₁ trains on few rows.\n", options: { color: BODY } },
    { text: "Var(μ̂₁) is large ", options: { bold: true, color: NAVY } }, { text: "— μ̂₁ is noisy.\n", options: { color: BODY } },
    { text: "Var(τ̂) = small + LARGE ", options: { bold: true, color: NAVY } }, { text: "— dominated by μ̂₁.\n", options: { color: BODY } },
    { text: "→ every τ(x) inherits μ̂₁'s noise.", options: { bold: true, color: RED } },
  ], { x: x2 + 0.3, y: top + 0.68, w: cw - 0.6, h: ch - 0.85, fontFace: BODY_FONT, fontSize: 12.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.35 });

  // Takeaway — bridges to the X-Learner section
  const tk_top = top + ch + 0.18, tk_h = 1.35;
  s.addShape("roundRect", { x: 0.6, y: tk_top, w: 12.1, h: tk_h, rectRadius: 0.08, fill: { color: NAVY_DK }, line: { type: "none" } });
  s.addText([
    { text: "So what does X-Learner actually fix?  ", options: { bold: true, color: GOLD } },
    { text: "Not a better model for the small arm — it stops asking μ̂₁ to predict outcomes for everyone. For each new-offer customer it uses their own real, observed Applied outcome, and only needs μ̂₀ (the reliable, large-sample old-offer model) to estimate what would've happened under the old offer instead. The minority arm's job shrinks from “run a whole noisy model on every customer” down to “fit one more regression on a small-but-real dataset.” Next: how that actually works, layer by layer.", options: { color: WHITE } },
  ], { x: 0.95, y: tk_top, w: 11.4, h: tk_h, valign: "middle", fontFace: BODY_FONT, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });

  pageNum(s, 7);
}

// ---------------------------------------------------------------------------
// Slide 8 — Comparison table + why X-learner
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "The Decision", "Why X-Learner, Specifically");

  const rows = [
    [
      { text: "", options: { fill: { color: NAVY } } },
      { text: "S-Learner", options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "center" } },
      { text: "T-Learner", options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "center" } },
      { text: "X-Learner", options: { fill: { color: GOLD }, color: WHITE, bold: true, align: "center" } },
    ],
    [
      { text: "How effect\nis estimated", options: { bold: true, color: NAVY } },
      { text: "Implicit — read off one shared model's new_offer term", options: {} },
      { text: "Difference of two\nindependently-fit models", options: {} },
      { text: "Directly regresses on an\nimputed treatment effect", options: { fill: { color: GOLD_BG } } },
    ],
    [
      { text: "Best when", options: { bold: true, color: NAVY } },
      { text: "Treatment effect is a strong,\neasily separable signal", options: {} },
      { text: "Treatment / control groups\nare large and balanced", options: {} },
      { text: "Groups are imbalanced —\ne.g. a small treatment group", options: { fill: { color: GOLD_BG } } },
    ],
    [
      { text: "Main risk", options: { bold: true, color: NAVY } },
      { text: "Regularization can shrink the\ntreatment signal toward zero", options: {} },
      { text: "Minority-arm model is noisy;\nerrors compound on subtraction", options: {} },
      { text: "More moving parts — more\nplaces to get cross-fitting wrong", options: { fill: { color: GOLD_BG } } },
    ],
  ];
  s.addTable(rows, {
    x: 0.6, y: 1.65, w: 12.1, h: 3.55,
    colW: [1.9, 3.4, 3.4, 3.4],
    fontFace: BODY_FONT, fontSize: 12, color: BODY, valign: "middle", align: "left",
    border: { type: "solid", color: "E4E7F0", pt: 1 },
    autoPage: false, margin: [0.08, 0.12, 0.08, 0.12], rowH: 0.85,
  });

  s.addShape("roundRect", { x: 0.6, y: 5.5, w: 12.1, h: 1.35, rectRadius: 0.08, fill: { color: NAVY_DK }, line: { type: "none" } });
  s.addText([
    { text: "Our real constraint: ", options: { bold: true, color: GOLD } },
    { text: "about 70% of customers stay on the existing offer (control) and only ~30% are rolled out to the new offer (treatment) — a deliberately smaller test slice to limit risk before scaling a new offer design. That's exactly T-Learner's weak point: μ̂₁, the new-offer-side model, is trained on the smaller group and is noisy. X-Learner's imputation step lets the new-offer-side effect estimate borrow strength from the reliable, well-estimated existing-offer-side model instead.", options: { color: WHITE } },
  ], {
    x: 0.95, y: 5.5, w: 11.4, h: 1.35, valign: "middle", fontFace: BODY_FONT, fontSize: 13, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25,
  });
  pageNum(s, 8);
}

// ---------------------------------------------------------------------------
// Slide 9 — X-Learner architecture (4-layer flow)
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "Architecture", "X-Learner, End to End");
  const steps = [
    { n: "1", t: "Outcome Models", d: "μ̂₀(x), μ̂₁(x) — one per arm, trained separately" },
    { n: "2", t: "Pseudo Effects", d: "Counterfactual predictions turn Applied into an imputed effect" },
    { n: "3", t: "Effect Models", d: "τC(x), τT(x) — regressors on the imputed effect" },
    { n: "4", t: "Weighted Fusion", d: "g(x) blends the two into one uplift score τ(x)" },
  ];
  const top = 2.1, bw = 2.55, bh = 2.6, gap = 0.62;
  const totalW = bw * 4 + gap * 3;
  const startX = (PW - totalW) / 2;
  steps.forEach((st, i) => {
    const x = startX + i * (bw + gap);
    const fill = i === 3 ? GOLD_BG : ICE_LT;
    const line = i === 3 ? GOLD : ICE;
    s.addShape("roundRect", { x, y: top, w: bw, h: bh, rectRadius: 0.1, fill: { color: fill }, line: { color: line, width: 1.5 } });
    s.addShape("ellipse", { x: x + bw / 2 - 0.32, y: top - 0.32, w: 0.64, h: 0.64, fill: { color: i === 3 ? GOLD : NAVY }, line: { color: WHITE, width: 2 } });
    s.addText(st.n, { x: x + bw / 2 - 0.32, y: top - 0.32, w: 0.64, h: 0.64, align: "center", valign: "middle", fontFace: HEAD_FONT, fontSize: 18, bold: true, color: WHITE, isTextBox: true, margin: 0 });
    s.addText(st.t, { x: x + 0.18, y: top + 0.5, w: bw - 0.36, h: 0.75, fontFace: HEAD_FONT, fontSize: 16, bold: true, color: NAVY, isTextBox: true, margin: 0, valign: "top" });
    s.addText(st.d, { x: x + 0.18, y: top + 1.3, w: bw - 0.36, h: 1.2, fontFace: BODY_FONT, fontSize: 11.5, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
    if (i < 3) {
      const ax = x + bw + 0.06;
      s.addText("→", { x: ax, y: top + bh / 2 - 0.35, w: gap - 0.12, h: 0.7, align: "center", valign: "middle", fontFace: BODY_FONT, fontSize: 26, bold: true, color: MUTED, isTextBox: true, margin: 0 });
    }
  });
  s.addText("Layers 1–3 are model training. Layer 4 happens at scoring time, for every customer.", {
    x: 0.6, y: top + bh + 0.55, w: 12.1, h: 0.4, align: "center", fontFace: BODY_FONT, fontSize: 12.5, italic: true, color: MUTED, isTextBox: true, margin: 0,
  });
  pageNum(s, 9);
}

// ---------------------------------------------------------------------------
// Slide 10 — Layer 1: two outcome models
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "Layer 1 of 4", "Two Outcome Models");
  s.addText([
    { text: "Control outcome model  ", options: { bold: true, color: NAVY } },
    { text: "learns μ̂₀(x) = E[Y(0) | X=x] — trained only on customers who stayed on the existing offer (control).\n", options: { color: BODY } },
    { text: "Treatment outcome model  ", options: { bold: true, color: NAVY } },
    { text: "learns μ̂₁(x) = E[Y(1) | X=x] — trained only on customers who got the new offer (treatment). (Y = 1 if they applied for the card, 0 if not.)", options: { color: BODY } },
  ], { x: 0.6, y: 1.4, w: 12.1, h: 0.62, fontFace: BODY_FONT, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });

  s.addShape("roundRect", { x: 0.6, y: 2.06, w: 12.1, h: 0.4, rectRadius: 0.06, fill: { color: "F2F2F5" }, line: { type: "none" } });
  s.addText("Both models are trained with out-of-fold / cross-fitted predictions — never a model scoring its own training rows. That avoids leakage later.", {
    x: 0.85, y: 2.06, w: 11.6, h: 0.4, valign: "middle", fontFace: BODY_FONT, fontSize: 10.5, italic: true, color: NAVY, isTextBox: true, margin: 0,
  });

  // TRAIN — each outcome model fit ONLY on its own arm
  const trainTop = 2.56, trainH = 1.7;
  s.addShape("roundRect", { x: 0.6, y: trainTop, w: 12.1, h: trainH, rectRadius: 0.1, fill: { color: ICE_LT }, line: { color: NAVY, width: 1 } });
  s.addText([
    { text: "TRAIN   ", options: { bold: true, color: NAVY } },
    { text: "each outcome model fit ONLY on its own arm", options: { color: MUTED, italic: true } },
  ], { x: 0.85, y: trainTop + 0.14, w: 11.6, h: 0.26, fontFace: BODY_FONT, fontSize: 11, isTextBox: true, margin: 0 });
  {
    const innerW = 11.6, gap = 0.5, miniW = (innerW - gap) / 2;
    const mx1 = 0.85, mx2 = mx1 + miniW + gap, labelY = trainTop + 0.48, tableY = trainTop + 0.72, tableH = trainH - 0.88;
    s.addText("Control (A,B,C) → μ̂₀ model", { x: mx1, y: labelY, w: miniW, h: 0.22, fontFace: BODY_FONT, fontSize: 10.5, bold: true, color: NAVY, isTextBox: true, margin: 0 });
    const ctrlHead = ["Customer", "Applied"].map(t => ({ text: t, options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "center", valign: "middle" } }));
    const ctrlRows = [ctrlHead, ...[["A", "0"], ["B", "1"], ["C", "0"]].map((r, i) => r.map((v, j) => ({
      text: v, options: { align: "center", valign: "middle", bold: j === 0, color: BODY, fill: { color: i % 2 === 0 ? WHITE : "F4F6FB" } },
    })))];
    s.addTable(ctrlRows, { x: mx1, y: tableY, w: miniW, h: tableH, colW: [miniW * 0.5, miniW * 0.5], fontFace: BODY_FONT, fontSize: 11, border: { type: "solid", color: "E4E7F0", pt: 1 }, autoPage: false, margin: [0.02, 0.05, 0.02, 0.05], rowH: tableH / 4 });

    s.addText("Treatment (D,E,F) → μ̂₁ model", { x: mx2, y: labelY, w: miniW, h: 0.22, fontFace: BODY_FONT, fontSize: 10.5, bold: true, color: GOLD, isTextBox: true, margin: 0 });
    const trtHead = ["Customer", "Applied"].map(t => ({ text: t, options: { fill: { color: GOLD }, color: WHITE, bold: true, align: "center", valign: "middle" } }));
    const trtRows = [trtHead, ...[["D", "1"], ["E", "0"], ["F", "1"]].map((r, i) => r.map((v, j) => ({
      text: v, options: { align: "center", valign: "middle", bold: j === 0, color: BODY, fill: { color: i % 2 === 0 ? WHITE : "F4F6FB" } },
    })))];
    s.addTable(trtRows, { x: mx2, y: tableY, w: miniW, h: tableH, colW: [miniW * 0.5, miniW * 0.5], fontFace: BODY_FONT, fontSize: 11, border: { type: "solid", color: "E4E7F0", pt: 1 }, autoPage: false, margin: [0.02, 0.05, 0.02, 0.05], rowH: tableH / 4 });
  }

  // SCORE — once trained, BOTH models score EVERY customer
  const scoreTop = trainTop + trainH + 0.15, scoreH = 2.5;
  s.addShape("roundRect", { x: 0.6, y: scoreTop, w: 12.1, h: scoreH, rectRadius: 0.1, fill: { color: GOLD_BG }, line: { color: GOLD, width: 1 } });
  s.addText([
    { text: "SCORE   ", options: { bold: true, color: GOLD } },
    { text: "once trained, BOTH models score EVERY customer — including their own counterfactual arm", options: { color: MUTED, italic: true } },
  ], { x: 0.85, y: scoreTop + 0.14, w: 11.6, h: 0.26, fontFace: BODY_FONT, fontSize: 11, isTextBox: true, margin: 0 });

  const header_row = ["Customer", "new_offer", "Applied", "μ̂₀(x)\npred. Applied, old offer", "μ̂₁(x)\npred. Applied, new offer"].map(t => (
    { text: t, options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "center", valign: "middle", fontSize: 9.5 } }
  ));
  const data = [
    ["A", "0 (old offer)", "0", "0.20", "0.55"],
    ["B", "0 (old offer)", "1", "0.60", "0.75"],
    ["C", "0 (old offer)", "0", "0.10", "0.30"],
    ["D", "1 (new offer)", "1", "0.35", "0.80"],
    ["E", "1 (new offer)", "0", "0.25", "0.40"],
    ["F", "1 (new offer)", "1", "0.50", "0.70"],
  ];
  const rows = [header_row, ...data.map((r, i) => r.map((v, j) => ({
    text: v,
    options: {
      align: j === 0 ? "left" : "center", valign: "middle",
      bold: j === 0,
      color: j === 1 ? (v.startsWith("0") ? NAVY : GOLD) : BODY,
      fill: { color: i % 2 === 0 ? WHITE : GOLD_BG },
    },
  })))];
  s.addTable(rows, {
    x: 1.0, y: scoreTop + 0.48, w: 11.3, h: scoreH - 0.63,
    colW: [1.5, 1.95, 1.5, 3.15, 3.2],
    fontFace: BODY_FONT, fontSize: 11, border: { type: "solid", color: "F0DFB0", pt: 1 },
    autoPage: false, margin: [0.04, 0.08, 0.04, 0.08], rowH: (scoreH - 0.63) / 7,
  });
  pageNum(s, 10);
}

// ---------------------------------------------------------------------------
// Slide 11 — Layer 2: pseudo effects
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "Layer 2 of 4", "Constructing Two Pseudo Effects");
  s.addText("X-Learner doesn't use μ̂₁ − μ̂₀ directly — it first imputes each unit's likely effect using the OTHER arm's model as counterfactual. (Applied = 1 if the customer applied for the card, 0 if not.)", {
    x: 0.6, y: 1.55, w: 12.1, h: 0.55, fontFace: BODY_FONT, fontSize: 13, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });

  // two formula + table columns
  const colW = 5.85, x1 = 0.6, x2 = 6.85, ytop = 2.25;
  formulaBox(s, "D̃ᵢᶜ = μ̂₁(Xᵢ) − Appliedᵢ", x1, ytop, colW, 0.6, { fill: GREEN, size: 15 });
  formulaBox(s, "D̃ᵢᵀ = Appliedᵢ − μ̂₀(Xᵢ)", x2, ytop, colW, 0.6, { fill: GOLD, size: 15 });

  s.addText("Control-side pseudo effect  ·  from A, B, C", { x: x1, y: ytop + 0.68, w: colW, h: 0.35, fontFace: BODY_FONT, fontSize: 12, bold: true, color: GREEN, isTextBox: true, margin: 0 });
  s.addText("Treatment-side pseudo effect  ·  from D, E, F", { x: x2, y: ytop + 0.68, w: colW, h: 0.35, fontFace: BODY_FONT, fontSize: 12, bold: true, color: GOLD, isTextBox: true, margin: 0 });

  function mkTable(rows, fillHead) {
    return [
      rows[0].map(t => ({ text: t, options: { fill: { color: fillHead }, color: WHITE, bold: true, align: "center", valign: "middle" } })),
      ...rows.slice(1).map((r, i) => r.map((v, j) => ({
        text: v, options: { align: "center", valign: "middle", bold: j === r.length - 1, color: j === r.length - 1 ? NAVY : BODY, fill: { color: i % 2 === 0 ? WHITE : ICE_LT } },
      }))),
    ];
  }
  const ctrlRows = mkTable([
    ["Cust.", "μ̂₁(x)", "Applied", "D̃ᶜ"],
    ["A", "0.55", "0", "0.55"],
    ["B", "0.75", "1", "−0.25"],
    ["C", "0.30", "0", "0.30"],
  ], GREEN);
  const trtRows = mkTable([
    ["Cust.", "Applied", "μ̂₀(x)", "D̃ᵀ"],
    ["D", "1", "0.35", "0.65"],
    ["E", "0", "0.25", "−0.25"],
    ["F", "1", "0.50", "0.50"],
  ], GOLD);
  s.addTable(ctrlRows, { x: x1, y: ytop + 1.1, w: colW, h: 1.9, colW: [1.3, 1.5, 1.3, 1.75], fontFace: BODY_FONT, fontSize: 13, border: { type: "solid", color: "E4E7F0", pt: 1 }, autoPage: false, margin: [0.05, 0.05, 0.05, 0.05], rowH: 0.46 });
  s.addTable(trtRows, { x: x2, y: ytop + 1.1, w: colW, h: 1.9, colW: [1.3, 1.3, 1.5, 1.75], fontFace: BODY_FONT, fontSize: 13, border: { type: "solid", color: "E4E7F0", pt: 1 }, autoPage: false, margin: [0.05, 0.05, 0.05, 0.05], rowH: 0.46 });

  s.addShape("roundRect", { x: 0.6, y: ytop + 3.2, w: 12.1, h: 1.15, rectRadius: 0.06, fill: { color: NAVY_DK }, line: { type: "none" } });
  s.addText([
    { text: "Why this beats a raw μ̂₁ − μ̂₀:  ", options: { bold: true, color: GOLD } },
    { text: "for a new-offer customer we don't need a model to guess their outcome — we use their own real, observed Applied value, and only borrow μ̂₀ (fit on the large old-offer group) to estimate what would've happened had they stayed on the old offer instead. The minority (treatment) arm no longer needs a full noisy model of its own — just one more regression on these pseudo effects. These are continuous, signed numbers, not 0/1 labels, and each becomes a training label in Layer 3.", options: { color: WHITE } },
  ], {
    x: 0.9, y: ytop + 3.2, w: 11.5, h: 1.15, valign: "middle", fontFace: BODY_FONT, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  pageNum(s, 11);
}

// ---------------------------------------------------------------------------
// Slide 12 — Training vs Scoring (the "what did A participate in" nuance)
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "A Common Follow-Up", "Training ≠ Scoring");
  s.addText("A frequent interviewer probe: “if A is a control-group customer, does A ever touch the treatment-side model?” Yes and no — at two different stages.", {
    x: 0.6, y: 1.55, w: 12.1, h: 0.5, fontFace: BODY_FONT, fontSize: 13, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });

  const cw = 5.85, x1 = 0.6, x2 = 6.85, top = 2.25, ch = 2.55;
  s.addShape("roundRect", { x: x1, y: top, w: cw, h: ch, rectRadius: 0.1, fill: { color: RED_BG }, line: { color: RED, width: 1.3 } });
  s.addText("✗  During TRAINING", { x: x1 + 0.3, y: top + 0.22, w: cw - 0.6, h: 0.4, fontFace: HEAD_FONT, fontSize: 16, bold: true, color: RED, isTextBox: true, margin: 0 });
  s.addText("A is control, so A's real Applied outcome and its counterfactual μ̂₁(A) build a control-side pseudo effect. That label trains effect model C only. A never contributes a label to effect model T's training set.", {
    x: x1 + 0.3, y: top + 0.75, w: cw - 0.6, h: ch - 1, fontFace: BODY_FONT, fontSize: 12.5, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25,
  });

  s.addShape("roundRect", { x: x2, y: top, w: cw, h: ch, rectRadius: 0.1, fill: { color: GREEN_BG }, line: { color: GREEN, width: 1.3 } });
  s.addText("✓  During SCORING", { x: x2 + 0.3, y: top + 0.22, w: cw - 0.6, h: 0.4, fontFace: HEAD_FONT, fontSize: 16, bold: true, color: GREEN, isTextBox: true, margin: 0 });
  s.addText("Once BOTH effect models are trained, they're just functions of X. Effect model T can score A too — it was never told A's outcome, but it doesn't need to have been. A gets both τC(A) and τT(A), fused like everyone else.", {
    x: x2 + 0.3, y: top + 0.75, w: cw - 0.6, h: ch - 1, fontFace: BODY_FONT, fontSize: 12.5, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25,
  });

  s.addShape("roundRect", { x: 0.6, y: top + ch + 0.35, w: 12.1, h: 0.95, rectRadius: 0.06, fill: { color: NAVY_DK }, line: { type: "none" } });
  s.addText([
    { text: "Why this matters:  ", options: { bold: true, color: GOLD } },
    { text: "it's what makes X-Learner one unified scorer instead of two disconnected ones — every customer, regardless of which arm they were actually in, gets a single fused uplift score at the end.", options: { color: WHITE } },
  ], { x: 0.9, y: top + ch + 0.35, w: 11.5, h: 0.95, valign: "middle", fontFace: BODY_FONT, fontSize: 12.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });
  pageNum(s, 12);
}

// ---------------------------------------------------------------------------
// Slide 13 — Layer 3: two effect models
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "Layer 3 of 4", "Two Effect Models");
  s.addText("The labels are now continuous effect values, not 0/1 — so both effect models are regressors, not classifiers.", {
    x: 0.6, y: 1.55, w: 12.1, h: 0.45, fontFace: BODY_FONT, fontSize: 13, color: BODY, isTextBox: true, margin: 0,
  });

  const cw = 5.85, x1 = 0.6, x2 = 6.85, top = 2.15, ch = 1.55;
  s.addShape("roundRect", { x: x1, y: top, w: cw, h: ch, rectRadius: 0.1, fill: { color: GREEN_BG }, line: { type: "none" } });
  s.addText("Effect Model C", { x: x1 + 0.3, y: top + 0.18, w: cw - 0.6, h: 0.4, fontFace: HEAD_FONT, fontSize: 16, bold: true, color: GREEN, isTextBox: true, margin: 0 });
  s.addText("Input X → predicts τC(x). Learned the pattern in the control-side pseudo effects.", { x: x1 + 0.3, y: top + 0.62, w: cw - 0.6, h: 0.8, fontFace: BODY_FONT, fontSize: 12.5, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  s.addShape("roundRect", { x: x2, y: top, w: cw, h: ch, rectRadius: 0.1, fill: { color: GOLD_BG }, line: { type: "none" } });
  s.addText("Effect Model T", { x: x2 + 0.3, y: top + 0.18, w: cw - 0.6, h: 0.4, fontFace: HEAD_FONT, fontSize: 16, bold: true, color: GOLD, isTextBox: true, margin: 0 });
  s.addText("Input X → predicts τT(x). Learned the pattern in the treatment-side pseudo effects.", { x: x2 + 0.3, y: top + 0.62, w: cw - 0.6, h: 0.8, fontFace: BODY_FONT, fontSize: 12.5, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2 });

  s.addShape("roundRect", { x: 0.6, y: top + ch + 0.2, w: 12.1, h: 0.85, rectRadius: 0.06, fill: { color: ICE_LT }, line: { type: "none" } });
  s.addText([
    { text: "What they output: ", options: { bold: true, color: NAVY } },
    { text: "“how big is this customer's incremental effect” — not “how likely are they to respond.” ", options: { color: NAVY } },
    { text: "Why two of them: ", options: { bold: true, color: NAVY } },
    { text: "each is trained on a different pseudo-effect sample — τC on the large control-side sample, τT on the small treatment-side one — so neither is fully trustworthy alone. Layer 4 decides how much to lean on each.", options: { color: NAVY } },
  ], { x: 0.85, y: top + ch + 0.2, w: 11.6, h: 0.85, valign: "middle", fontFace: BODY_FONT, fontSize: 11.5, isTextBox: true, margin: 0, lineSpacingMultiple: 1.22 });

  const wy = top + ch + 1.2;
  s.addText("Worked example — a new customer, G, scored by both:", { x: 0.6, y: wy, w: 12.1, h: 0.35, fontFace: BODY_FONT, fontSize: 13, bold: true, color: NAVY, isTextBox: true, margin: 0 });

  const bw = 2.7, bh = 1.1, bgap = 0.6, bx = (PW - bw * 2 - bgap) / 2, by = wy + 0.42;
  s.addShape("roundRect", { x: bx, y: by, w: bw, h: bh, rectRadius: 0.08, fill: { color: GREEN }, line: { type: "none" } });
  s.addText("τC(G)", { x: bx, y: by + 0.12, w: bw, h: 0.32, align: "center", fontFace: BODY_FONT, fontSize: 12.5, color: WHITE, isTextBox: true, margin: 0 });
  s.addText("0.12", { x: bx, y: by + 0.42, w: bw, h: 0.6, align: "center", fontFace: HEAD_FONT, fontSize: 27, bold: true, color: WHITE, isTextBox: true, margin: 0 });

  s.addShape("roundRect", { x: bx + bw + bgap, y: by, w: bw, h: bh, rectRadius: 0.08, fill: { color: GOLD }, line: { type: "none" } });
  s.addText("τT(G)", { x: bx + bw + bgap, y: by + 0.12, w: bw, h: 0.32, align: "center", fontFace: BODY_FONT, fontSize: 12.5, color: WHITE, isTextBox: true, margin: 0 });
  s.addText("0.28", { x: bx + bw + bgap, y: by + 0.42, w: bw, h: 0.6, align: "center", fontFace: HEAD_FONT, fontSize: 27, bold: true, color: WHITE, isTextBox: true, margin: 0 });

  s.addText("Two different estimates of the same customer's uplift — neither is the final answer yet.", {
    x: 0.6, y: by + bh + 0.12, w: 12.1, h: 0.35, align: "center", italic: true, fontFace: BODY_FONT, fontSize: 11.5, color: MUTED, isTextBox: true, margin: 0,
  });
  pageNum(s, 13);
}

// ---------------------------------------------------------------------------
// Slide 14 — Layer 4: weighted fusion
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "Layer 4 of 4", "Weighted Fusion");
  formulaBox(s, "τ(x) = g(x)·τT(x) + (1 − g(x))·τC(x)", 0.6, 1.6, 12.1, 0.75, { fill: NAVY, size: 20, align: "center" });
  s.addText("g(x) is the propensity score. τC and τT were each built from a different-sized sample with a different kind of noise (previous slide) — neither is fully reliable alone. g(x) is how the model calibrates how much to lean on each, instead of trusting just one arm's estimate.", {
    x: 0.6, y: 2.48, w: 12.1, h: 0.62, align: "center", italic: true, fontFace: BODY_FONT, fontSize: 11.5, color: MUTED, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });

  s.addText("Worked example, continuing with customer G:", { x: 0.6, y: 3.28, w: 12.1, h: 0.35, fontFace: BODY_FONT, fontSize: 13, bold: true, color: NAVY, isTextBox: true, margin: 0 });

  const rows = [
    ["τT(G)", "0.28", "×", "g(G)", "0.7", "=", "0.196"],
    ["τC(G)", "0.12", "×", "1 − g(G)", "0.3", "=", "0.036"],
  ];
  const y0 = 3.75, rh = 0.6;
  rows.forEach((r, i) => {
    const y = y0 + i * (rh + 0.15);
    s.addShape("roundRect", { x: 0.6, y, w: 8.9, h: rh, rectRadius: 0.06, fill: { color: i === 0 ? GOLD_BG : GREEN_BG }, line: { type: "none" } });
    s.addText(r[0], { x: 0.85, y, w: 1.3, h: rh, valign: "middle", fontFace: "Courier New", fontSize: 15, bold: true, color: i === 0 ? GOLD : GREEN, isTextBox: true, margin: 0 });
    s.addText(r[1], { x: 2.15, y, w: 1.1, h: rh, valign: "middle", fontFace: "Courier New", fontSize: 15, color: BODY, isTextBox: true, margin: 0 });
    s.addText(r[2], { x: 3.25, y, w: 0.4, h: rh, valign: "middle", align: "center", fontFace: BODY_FONT, fontSize: 15, color: MUTED, isTextBox: true, margin: 0 });
    s.addText(r[3], { x: 3.65, y, w: 1.9, h: rh, valign: "middle", fontFace: "Courier New", fontSize: 15, bold: true, color: BODY, isTextBox: true, margin: 0 });
    s.addText(r[4], { x: 5.55, y, w: 1.0, h: rh, valign: "middle", fontFace: "Courier New", fontSize: 15, color: BODY, isTextBox: true, margin: 0 });
    s.addText(r[5], { x: 6.55, y, w: 0.4, h: rh, valign: "middle", align: "center", fontFace: BODY_FONT, fontSize: 15, color: MUTED, isTextBox: true, margin: 0 });
    s.addText(r[6], { x: 6.95, y, w: 1.4, h: rh, valign: "middle", fontFace: "Courier New", fontSize: 15, bold: true, color: NAVY, isTextBox: true, margin: 0 });
  });

  const fx = 9.85, fy = y0 - 0.1, fw = 2.85, fh = rh * 2 + 0.15 + 0.2;
  s.addShape("roundRect", { x: fx, y: fy, w: fw, h: fh, rectRadius: 0.1, fill: { color: NAVY }, line: { type: "none" } });
  s.addText("τ(G) =", { x: fx, y: fy + 0.15, w: fw, h: 0.35, align: "center", fontFace: BODY_FONT, fontSize: 13, color: ICE, isTextBox: true, margin: 0 });
  s.addText("0.232", { x: fx, y: fy + 0.45, w: fw, h: 0.75, align: "center", fontFace: HEAD_FONT, fontSize: 36, bold: true, color: GOLD, isTextBox: true, margin: 0 });
  s.addText("final uplift score", { x: fx, y: fy + fh - 0.35, w: fw, h: 0.3, align: "center", fontFace: BODY_FONT, fontSize: 10.5, italic: true, color: ICE, isTextBox: true, margin: 0 });

  s.addShape("roundRect", { x: 0.6, y: 5.55, w: 12.1, h: 0.85, rectRadius: 0.06, fill: { color: ICE_LT }, line: { type: "none" } });
  s.addText("0.196 + 0.036 = 0.232 — this single number is what gets ranked, thresholded, and acted on. Everything before this slide exists to produce it responsibly.", {
    x: 0.9, y: 5.55, w: 11.5, h: 0.85, valign: "middle", fontFace: BODY_FONT, fontSize: 13, color: NAVY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });
  pageNum(s, 14);
}

// ---------------------------------------------------------------------------
// Slide 15 — Validation
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "Does It Actually Work?", "Validation, Two Ways");
  s.addText("You never observe both outcomes for the same person, so uplift can't be checked row-by-row like a classifier. The standard tool is a Qini curve — rank by score, then compare treated vs. control response as you move down the ranking.", {
    x: 0.6, y: 1.5, w: 12.1, h: 0.65, fontFace: BODY_FONT, fontSize: 12.5, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });

  s.addImage({ path: "qini_curve.png", x: 0.6, y: 2.25, w: 6.5, h: 4.55 });
  s.addShape("rect", { x: 0.6, y: 2.25, w: 6.5, h: 4.55, fill: { type: "none" }, line: { color: "E4E7F0", width: 1 } });

  const stx = 7.5, stw = 5.2;
  const stats = [
    { n: "0.93", l: "Spearman correlation between estimated and true uplift — only checkable because this demo uses synthetic data with known ground truth." },
    { n: "96%", l: "of the oracle's achievable Qini gain captured by the X-learner — how close the model gets to the best possible ranking." },
    { n: "+8.5%", l: "incremental response in the top uplift decile — vs. a population-wide average lift near zero." },
  ];
  let sy = 2.25;
  stats.forEach((st) => {
    const sh = 1.42;
    s.addShape("roundRect", { x: stx, y: sy, w: stw, h: sh, rectRadius: 0.08, fill: { color: ICE_LT }, line: { type: "none" } });
    s.addText(st.n, { x: stx + 0.25, y: sy + 0.12, w: 2.0, h: sh - 0.24, valign: "middle", fontFace: HEAD_FONT, fontSize: 30, bold: true, color: NAVY, isTextBox: true, margin: 0 });
    s.addText(st.l, { x: stx + 2.05, y: sy + 0.14, w: stw - 2.3, h: sh - 0.28, valign: "middle", fontFace: BODY_FONT, fontSize: 10.5, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
    sy += sh + 0.22;
  });
  pageNum(s, 15);
}

// ---------------------------------------------------------------------------
// Slide 16 — From score to decision
// ---------------------------------------------------------------------------
{
  const s = newSlide(WHITE);
  header(s, "Putting It To Work", "From Score to Decision");

  const rows = [
    [
      { text: "Customer", options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "center", valign: "middle" } },
      { text: "Final uplift τ(x)", options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "center", valign: "middle" } },
      { text: "Action", options: { fill: { color: NAVY }, color: WHITE, bold: true, align: "center", valign: "middle" } },
    ],
    [
      { text: "G", options: { align: "center", valign: "middle", bold: true, fill: { color: GREEN_BG } } },
      { text: "0.232", options: { align: "center", valign: "middle", bold: true, color: GREEN, fill: { color: GREEN_BG } } },
      { text: "Switch to new offer — first", options: { align: "center", valign: "middle", color: GREEN, fill: { color: GREEN_BG } } },
    ],
    [
      { text: "H", options: { align: "center", valign: "middle", bold: true, fill: { color: GOLD_BG } } },
      { text: "0.05", options: { align: "center", valign: "middle", bold: true, color: GOLD, fill: { color: GOLD_BG } } },
      { text: "Borderline — only if budget allows", options: { align: "center", valign: "middle", color: GOLD, fill: { color: GOLD_BG } } },
    ],
    [
      { text: "I", options: { align: "center", valign: "middle", bold: true, fill: { color: RED_BG } } },
      { text: "−0.02", options: { align: "center", valign: "middle", bold: true, color: RED, fill: { color: RED_BG } } },
      { text: "Keep old offer — new one likely backfires", options: { align: "center", valign: "middle", color: RED, fill: { color: RED_BG } } },
    ],
  ];
  s.addTable(rows, {
    x: 0.6, y: 1.6, w: 6.0, h: 2.6, colW: [1.4, 2.0, 2.6],
    fontFace: BODY_FONT, fontSize: 12, border: { type: "solid", color: "E4E7F0", pt: 1 },
    autoPage: false, margin: [0.06, 0.08, 0.06, 0.08], rowH: 0.65,
  });
  s.addText("The final score drives three things: rank the file high-to-low, roll out the new offer down to a budget/capacity cutoff, and treat a low or negative score as a decision to keep that customer on the old offer — not just a low priority.", {
    x: 0.6, y: 4.4, w: 6.0, h: 1.4, fontFace: BODY_FONT, fontSize: 12, color: BODY, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25,
  });

  // right column: the real insight from the demo
  s.addShape("roundRect", { x: 7.0, y: 1.6, w: 5.7, h: 4.9, rectRadius: 0.1, fill: { color: NAVY }, line: { type: "none" } });
  s.addText("THE ARGUMENT FOR UPLIFT MODELING", { x: 7.35, y: 1.9, w: 5.0, h: 0.35, fontFace: BODY_FONT, fontSize: 11.5, bold: true, color: GOLD, charSpacing: 1.5, isTextBox: true, margin: 0 });
  s.addText("In the full simulation, switching everyone to the new offer nets an overall incremental lift close to zero — persuadable gains are offset by sleeping-dog losses.", {
    x: 7.35, y: 2.35, w: 5.0, h: 1.1, fontFace: BODY_FONT, fontSize: 13, color: WHITE, isTextBox: true, margin: 0, lineSpacingMultiple: 1.25,
  });
  s.addText("A plain response model — or a simple before/after read — would look at that flat average and conclude the campaign doesn't work.", {
    x: 7.35, y: 3.5, w: 5.0, h: 1.0, fontFace: BODY_FONT, fontSize: 13, color: "C9D2F0", isTextBox: true, margin: 0, lineSpacingMultiple: 1.25, italic: true,
  });
  s.addShape("line", { x: 7.35, y: 4.6, w: 4.65, h: 0, line: { color: "3A459A", width: 1 } });
  s.addText("Uplift modeling reveals the real, exploitable heterogeneity hiding underneath that flat average.", {
    x: 7.35, y: 4.8, w: 5.0, h: 1.0, fontFace: HEAD_FONT, fontSize: 15.5, bold: true, color: WHITE, isTextBox: true, margin: 0, lineSpacingMultiple: 1.3,
  });
  pageNum(s, 16);
}

// ---------------------------------------------------------------------------
// Slide 17 — One-minute answer + closing
// ---------------------------------------------------------------------------
{
  const s = newSlide(NAVY);
  s.addText("THE ONE-MINUTE ANSWER", { x: 0.7, y: 0.55, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 13, bold: true, color: GOLD, charSpacing: 2, isTextBox: true, margin: 0 });
  s.addText("If I only get one minute to explain X-Learner", { x: 0.7, y: 0.95, w: 11.5, h: 0.6, fontFace: HEAD_FONT, fontSize: 26, bold: true, color: WHITE, isTextBox: true, margin: 0 });

  const steps = [
    ["1", "Split & train", "Split historical data into treatment / control, train one outcome model per arm."],
    ["2", "Impute effects", "For each unit, use the OTHER arm's model as its counterfactual to build a pseudo effect."],
    ["3", "Learn the effect", "Train two regressors directly on those pseudo effects — output is uplift, not a response probability."],
    ["4", "Fuse & rank", "Blend both effect models with propensity weight g(x) into one score; rank and target on it."],
  ];
  const top = 1.65, rh = 0.88, gap = 0.14;
  steps.forEach((st, i) => {
    const y = top + i * (rh + gap);
    s.addShape("ellipse", { x: 0.7, y: y + 0.1, w: 0.52, h: 0.52, fill: { color: GOLD }, line: { type: "none" } });
    s.addText(st[0], { x: 0.7, y: y + 0.1, w: 0.52, h: 0.52, align: "center", valign: "middle", fontFace: HEAD_FONT, fontSize: 17, bold: true, color: NAVY_DK, isTextBox: true, margin: 0 });
    s.addText(st[1], { x: 1.5, y, w: 3.0, h: rh, valign: "middle", fontFace: HEAD_FONT, fontSize: 15.5, bold: true, color: WHITE, isTextBox: true, margin: 0 });
    s.addText(st[2], { x: 4.55, y, w: 8.1, h: rh, valign: "middle", fontFace: BODY_FONT, fontSize: 12.5, color: ICE, isTextBox: true, margin: 0, lineSpacingMultiple: 1.15 });
  });

  const closeY = top + 4 * (rh + gap) + 0.12;
  s.addShape("line", { x: 0.7, y: closeY, w: 11.9, h: 0, line: { color: "3A459A", width: 1 } });
  s.addText("This is the estimator family behind a live prescreen campaign that lifted response 20% and cut acquisition cost 16% — this deck and notebook reproduce the method on synthetic data.", {
    x: 0.7, y: closeY + 0.18, w: 11.9, h: 0.6, fontFace: BODY_FONT, fontSize: 11.5, italic: true, color: "9FB4E8", isTextBox: true, margin: 0, lineSpacingMultiple: 1.2,
  });

  s.addNotes(
    "1-minute Chinese answer (verbatim, for rehearsal):\n\n" +
    "X-learner 的流程是这样的：\n" +
    "1. 先把历史实验数据分成 treatment（新 offer）和 control（旧 offer）两组，分别训练两个 outcome model，一个学新 offer 下的结果，一个学旧 offer 下的结果。\n" +
    "2. 然后对每个样本做反事实预测，构造 pseudo effect：对 control 组样本，用 treatment model 预测值减真实结果；对 treatment 组样本，用真实结果减 control model 预测值。构造两类 pseudo effect（control-side pseudo effect, treatment-side pseudo effect）。\n" +
    "3. 接着用这些 pseudo effect 分别训练两个 effect model，effect model C 和 effect model T。它们输出的是连续的 uplift 分数，而不是回复概率。\n" +
    "4. 最后再用一个权重 g(x) 把两个 effect model 的结果融合，得到最终 uplift，用来排序和选人。这样做的好处是，它不只是看谁最可能回复，而是看谁因为换成新 offer 会产生最大的增量响应。\n\n" +
    "English equivalent:\n" +
    "1. Split historical experiment data into treatment (new offer) and control (old offer), train two outcome models — one for what happens with the new offer, one for the existing offer.\n" +
    "2. For each unit, do a counterfactual prediction to build a pseudo effect: for control units, treatment-model prediction minus actual outcome; for treatment units, actual outcome minus control-model prediction. Two kinds of pseudo effect result.\n" +
    "3. Train two effect models on these pseudo effects. Their output is a continuous uplift score, not a response probability.\n" +
    "4. Finally, blend the two effect models with a weight g(x) into the final uplift score, used for ranking and targeting. The benefit: you're not just looking at who's most likely to respond, but who will produce the biggest incremental response from switching to the new offer."
  );
  pageNum(s, 17);
}

pres.writeFile({ fileName: "uplift_model_deck.pptx" }).then(() => console.log("deck written"));

