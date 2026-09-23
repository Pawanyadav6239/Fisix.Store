// Builds CleanCarbon_Assurance_CBAM_Verification.pptx
// Usage: node build-deck.js   (needs pptxgenjs, react, react-dom, react-icons, sharp)
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

// ---------- Palette & type ----------
const C = {
  deep: "0B3B36",     // deep pine – dominant dark
  pine: "12574F",
  emerald: "1E8A6E",
  mint: "DDEFE8",
  mist: "F3F8F6",
  amber: "E8A93A",    // sharp accent
  ink: "1C2B28",
  muted: "5C6F6A",
  white: "FFFFFF",
};
const HEAD = "Cambria";
const BODY = "Calibri";
const W = 13.333, H = 7.5;

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// Icon inside a filled circle – the deck's recurring motif
async function badge(slide, Comp, x, y, d, fill, fg) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } });
  const pad = d * 0.26;
  slide.addImage({ data: await icon(Comp, fg), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
}

function title(slide, text, opts = {}) {
  slide.addText(text, {
    x: 0.6, y: 0.45, w: W - 1.2, h: 0.9, fontFace: HEAD, fontSize: 34, bold: true,
    color: opts.color || C.deep, margin: 0, valign: "top", isTextBox: true,
  });
}
function kicker(slide, text, color) {
  slide.addText(text.toUpperCase(), {
    x: 0.6, y: 0.18, w: 8, h: 0.3, fontFace: BODY, fontSize: 11, bold: true,
    charSpacing: 3, color: color || C.emerald, margin: 0, isTextBox: true,
  });
}
function footer(slide, n, dark) {
  slide.addText("CleanCarbon Assurance  |  cbamassure.com", {
    x: 0.6, y: H - 0.45, w: 6, h: 0.3, fontFace: BODY, fontSize: 9,
    color: dark ? "9FC3B8" : C.muted, margin: 0, isTextBox: true,
  });
  slide.addText(String(n), {
    x: W - 1.1, y: H - 0.45, w: 0.5, h: 0.3, fontFace: BODY, fontSize: 9, align: "right",
    color: dark ? "9FC3B8" : C.muted, margin: 0, isTextBox: true,
  });
}
const card = (fill) => ({
  fill: { color: fill || C.white }, line: { color: fill || C.white }, rectRadius: 0.12,
  shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 90, opacity: 0.12 },
});

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "CleanCarbon Assurance";
  pres.company = "CleanCarbon Assurance";
  pres.title = "CleanCarbon Assurance – Independent CBAM Verification";

  // ================= 1. Cover =================
  {
    const s = pres.addSlide();
    s.background = { color: C.deep };
    // large decorative rings (motif: circles)
    s.addShape("ellipse", { x: 8.6, y: -1.4, w: 6.4, h: 6.4, fill: { color: C.pine }, line: { color: C.pine } });
    s.addShape("ellipse", { x: 10.2, y: 3.6, w: 4.2, h: 4.2, fill: { color: C.emerald, transparency: 55 }, line: { color: C.emerald, transparency: 55 } });
    await badge(s, fa.FaLeaf, 10.35, 1.05, 2.3, C.amber, C.deep);

    s.addText("CLEANCARBON ASSURANCE", {
      x: 0.8, y: 1.3, w: 7.5, h: 0.4, fontFace: BODY, fontSize: 14, bold: true, charSpacing: 4,
      color: C.amber, margin: 0, isTextBox: true,
    });
    s.addText("Independent CBAM Verification for Exporters to the EU", {
      x: 0.8, y: 1.85, w: 8.2, h: 2.2, fontFace: HEAD, fontSize: 44, bold: true, color: C.white,
      margin: 0, valign: "top", isTextBox: true,
    });
    s.addText("Verified embedded-emissions data your EU customers can rely on, so you don't have to fall back on default values.", {
      x: 0.8, y: 4.25, w: 7.6, h: 1.0, fontFace: BODY, fontSize: 18, color: "CFE5DD",
      margin: 0, valign: "top", isTextBox: true,
    });
    s.addText("cbamassure.com   ·   nb@cbamassure.com", {
      x: 0.8, y: 6.3, w: 7, h: 0.4, fontFace: BODY, fontSize: 13, color: C.white, margin: 0, isTextBox: true,
    });
    s.addNotes("Introduce CleanCarbon Assurance: a specialist, independent verification body focused on the EU Carbon Border Adjustment Mechanism (CBAM). Our job is to give EU importers confidence in the embedded-emissions data your installation reports.");
  }

  // ================= 2. CBAM timeline =================
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    kicker(s, "The regulatory context");
    title(s, "CBAM has moved from reporting to paying");
    s.addText("The EU's Carbon Border Adjustment Mechanism (Regulation (EU) 2023/956) puts a carbon price on the emissions embedded in imported goods. From 2026, costs depend directly on the emissions data you supply.", {
      x: 0.6, y: 1.35, w: 11.8, h: 0.8, fontFace: BODY, fontSize: 15, color: C.muted, margin: 0, valign: "top", isTextBox: true,
    });

    const steps = [
      ["Oct 2023 – Dec 2025", "Transitional period", "Quarterly reporting only. No financial obligation, and verification not required."],
      ["1 Jan 2026", "Definitive period begins", "Authorised declarants import CBAM goods. 2026 embedded emissions count toward the first bill."],
      ["2027", "Certificates & first declaration", "CBAM certificate sales open. The annual declaration for 2026 imports is due by 30 Sep 2027."],
      ["2026 – 2034", "Free allocation phase-out", "EU ETS free allowances fall year by year, so the CBAM cost per tonne keeps rising."],
    ];
    const y0 = 3.05, x0 = 0.9, gap = 3.0;
    s.addShape("line", { x: x0 + 0.3, y: y0 + 0.3, w: gap * 3, h: 0, line: { color: C.mint, width: 4 } });
    for (let i = 0; i < steps.length; i++) {
      const x = x0 + i * gap;
      const hot = i === 1;
      s.addShape("ellipse", { x, y: y0, w: 0.6, h: 0.6, fill: { color: hot ? C.amber : C.emerald }, line: { color: C.white, width: 3 } });
      s.addText(String(i + 1), { x, y: y0, w: 0.6, h: 0.6, align: "center", valign: "middle", fontFace: BODY, bold: true, fontSize: 16, color: hot ? C.deep : C.white, margin: 0, isTextBox: true });
      s.addText(steps[i][0], { x: x - 0.05, y: y0 + 0.85, w: 2.75, h: 0.35, fontFace: BODY, fontSize: 13, bold: true, color: C.emerald, margin: 0, isTextBox: true });
      s.addText(steps[i][1], { x: x - 0.05, y: y0 + 1.2, w: 2.75, h: 0.7, fontFace: HEAD, fontSize: 18, bold: true, color: C.deep, margin: 0, valign: "top", isTextBox: true });
      s.addText(steps[i][2], { x: x - 0.05, y: y0 + 1.95, w: 2.7, h: 1.3, fontFace: BODY, fontSize: 13, color: C.ink, margin: 0, valign: "top", isTextBox: true });
    }
    footer(s, 2);
    s.addNotes("Key message: 2026 data is already live. Emissions embedded in goods imported during 2026 will be paid for through the declaration due in 2027, so exporters need verified data for 2026 production. Check the latest EU implementing acts before quoting exact dates, since the Commission has adjusted some timelines through the 2025 simplification package.");
  }

  // ================= 3. Default vs verified =================
  {
    const s = pres.addSlide();
    s.background = { color: C.mist };
    kicker(s, "Why verification matters");
    title(s, "Verified data vs. default values");

    const colW = 5.75, y = 1.75, h = 4.6;
    // Left: default values
    s.addShape("roundRect", { x: 0.6, y, w: colW, h, ...card(C.white) });
    await badge(s, fa.FaExclamationTriangle, 0.95, y + 0.35, 0.8, "F6E3C2", "A8661A");
    s.addText("Without verified data", { x: 1.95, y: y + 0.45, w: 4.2, h: 0.6, fontFace: HEAD, fontSize: 22, bold: true, color: C.ink, margin: 0, isTextBox: true });
    s.addText([
      { text: "Your EU importer must use EU default values for your goods", options: { bullet: true, breakLine: true } },
      { text: "Default values carry mark-ups that grow each year", options: { bullet: true, breakLine: true } },
      { text: "A higher CBAM bill for your customer makes your product less competitive", options: { bullet: true, breakLine: true } },
      { text: "Your investment in cleaner production goes unrecognised", options: { bullet: true } },
    ], { x: 0.95, y: y + 1.35, w: 5.1, h: 3.0, fontFace: BODY, fontSize: 15, color: C.ink, paraSpaceAfter: 10, valign: "top", margin: 0, isTextBox: true });

    // Right: verified
    const rx = 0.6 + colW + 0.5;
    s.addShape("roundRect", { x: rx, y, w: colW, h, ...card(C.deep) });
    await badge(s, fa.FaCheckCircle, rx + 0.35, y + 0.35, 0.8, C.amber, C.deep);
    s.addText("With CleanCarbon verification", { x: rx + 1.35, y: y + 0.45, w: 4.3, h: 0.6, fontFace: HEAD, fontSize: 22, bold: true, color: C.white, margin: 0, isTextBox: true });
    s.addText([
      { text: "Actual embedded emissions, verified to CBAM rules, are accepted in the importer's declaration", options: { bullet: true, breakLine: true } },
      { text: "A lower and more predictable carbon cost for your EU buyers", options: { bullet: true, breakLine: true } },
      { text: "A commercial edge in tenders and supplier selection", options: { bullet: true, breakLine: true } },
      { text: "Credible evidence of your decarbonisation progress", options: { bullet: true } },
    ], { x: rx + 0.35, y: y + 1.35, w: 5.1, h: 3.0, fontFace: BODY, fontSize: 15, color: "E4F1EC", paraSpaceAfter: 10, valign: "top", margin: 0, isTextBox: true });
    footer(s, 3);
    s.addNotes("In the definitive period, actual embedded emissions can only be used in a CBAM declaration when they have been verified by an accredited verifier. Otherwise the declarant falls back on default values, which the Commission sets conservatively and increases with mark-ups over time.");
  }

  // ================= 4. Sectors =================
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    kicker(s, "Scope");
    title(s, "CBAM goods we verify");
    const items = [
      [fa.FaIndustry, "Iron & Steel", "Pig iron, DRI, crude steel, flat & long products, fasteners, articles of steel"],
      [fa.FaCubes, "Aluminium", "Unwrought aluminium, bars, profiles, sheets, foil and articles"],
      [fa.FaBuilding, "Cement", "Clinker, Portland and other hydraulic cements, calcined clay"],
      [fa.FaSeedling, "Fertilisers", "Ammonia, nitric acid, urea, ammonium nitrate, mixed NPK fertilisers"],
      [fa.FaAtom, "Hydrogen", "Hydrogen from all production routes"],
      [fa.FaBolt, "Electricity", "Electricity imported into the EU"],
    ];
    const cw = 3.85, ch = 2.35, gx = 0.3, gy = 0.35, x0 = 0.6, y0 = 1.55;
    for (let i = 0; i < items.length; i++) {
      const c = i % 3, r = Math.floor(i / 3);
      const x = x0 + c * (cw + gx), y = y0 + r * (ch + gy);
      s.addShape("roundRect", { x, y, w: cw, h: ch, ...card(C.mist) });
      await badge(s, items[i][0], x + 0.3, y + 0.3, 0.75, C.emerald, C.white);
      s.addText(items[i][1], { x: x + 1.25, y: y + 0.38, w: cw - 1.45, h: 0.6, fontFace: HEAD, fontSize: 20, bold: true, color: C.deep, margin: 0, valign: "middle", isTextBox: true });
      s.addText(items[i][2], { x: x + 0.3, y: y + 1.2, w: cw - 0.6, h: 1.0, fontFace: BODY, fontSize: 13, color: C.ink, margin: 0, valign: "top", isTextBox: true });
    }
    footer(s, 4);
    s.addNotes("Includes relevant precursors (e.g. sintered ore, pig iron, clinker, ammonia) whose embedded emissions flow into the final good. Downstream steel and aluminium products are the largest category by number of exporters.");
  }

  // ================= 5. Services =================
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    kicker(s, "What we offer");
    title(s, "Our services");

    // Left hero panel
    s.addShape("roundRect", { x: 0.6, y: 1.55, w: 4.6, h: 5.0, ...card(C.deep) });
    await badge(s, fa.FaShieldAlt, 0.95, 1.9, 1.0, C.amber, C.deep);
    s.addText("Core service", { x: 0.95, y: 3.1, w: 3.9, h: 0.35, fontFace: BODY, fontSize: 12, bold: true, charSpacing: 3, color: C.amber, margin: 0, isTextBox: true });
    s.addText("CBAM Installation Emissions Verification", { x: 0.95, y: 3.45, w: 3.95, h: 1.2, fontFace: HEAD, fontSize: 22, bold: true, color: C.white, margin: 0, valign: "top", isTextBox: true });
    s.addText("Independent verification of your installation's emissions report and the specific embedded emissions (direct and indirect) of each CBAM good, ending in a verification report you can share with EU declarants.", {
      x: 0.95, y: 4.35, w: 3.95, h: 1.9, fontFace: BODY, fontSize: 13, color: "D3E8E1", margin: 0, valign: "top", isTextBox: true,
    });

    const rows = [
      [fa.FaClipboardCheck, "Verification readiness review", "A limited pre-assessment of your monitoring plan, data flows and evidence before the formal verification, so there are no surprises on site."],
      [fa.FaGlobeEurope, "GHG verification to ISO 14064-3", "Organisation- and product-level GHG statements for customers, investors and voluntary disclosures."],
      [fa.FaHandshake, "EU customer support", "Clear verification reports and guidance on sharing verified data with your EU importers through the CBAM registry."],
    ];
    const rx = 5.6;
    for (let i = 0; i < rows.length; i++) {
      const y = 1.65 + i * 1.65;
      await badge(s, rows[i][0], rx, y, 0.85, C.mint, C.pine);
      s.addText(rows[i][1], { x: rx + 1.15, y: y - 0.02, w: 6.0, h: 0.45, fontFace: HEAD, fontSize: 19, bold: true, color: C.deep, margin: 0, isTextBox: true });
      s.addText(rows[i][2], { x: rx + 1.15, y: y + 0.45, w: 6.0, h: 0.95, fontFace: BODY, fontSize: 14, color: C.ink, margin: 0, valign: "top", isTextBox: true });
    }
    footer(s, 5);
    s.addNotes("To protect impartiality, we do not design monitoring plans or calculate the emissions we later verify. The readiness review only identifies gaps; your team or your consultant closes them.");
  }

  // ================= 6. Process =================
  {
    const s = pres.addSlide();
    s.background = { color: C.mist };
    kicker(s, "How it works");
    title(s, "Our verification process");
    const steps = [
      [fa.FaFileSignature, "Engagement", "Impartiality and competence checks, scope confirmation, contract"],
      [fa.FaSearch, "Risk analysis", "Review of installation, production processes, monitoring methodology and data risks"],
      [fa.FaMapMarkedAlt, "Verification plan", "Sampling plan and data tests aimed at the highest-risk areas"],
      [fa.FaHardHat, "Site visit", "Physical inspection of the installation, meters, records and staff interviews"],
      [fa.FaCalculator, "Data testing", "Recalculation of embedded emissions, precursors, electricity and allocation to goods"],
      [fa.FaFileAlt, "Report & opinion", "Findings raised and closed, independent technical review, verification report issued"],
    ];
    const n = steps.length, bw = 1.88, gap = 0.19, x0 = 0.6, y = 1.75;
    for (let i = 0; i < n; i++) {
      const x = x0 + i * (bw + gap);
      s.addShape("roundRect", { x, y, w: bw, h: 4.0, ...card(C.white) });
      s.addText(String(i + 1).padStart(2, "0"), { x: x + 0.2, y: y + 0.2, w: 0.8, h: 0.45, fontFace: HEAD, fontSize: 22, bold: true, color: C.amber, margin: 0, isTextBox: true });
      await badge(s, steps[i][0], x + bw - 0.85, y + 0.2, 0.65, C.emerald, C.white);
      s.addText(steps[i][1], { x: x + 0.2, y: y + 1.1, w: bw - 0.4, h: 0.8, fontFace: HEAD, fontSize: 15, bold: true, color: C.deep, margin: 0, valign: "top", isTextBox: true });
      s.addText(steps[i][2], { x: x + 0.2, y: y + 1.9, w: bw - 0.4, h: 2.3, fontFace: BODY, fontSize: 13, color: C.ink, margin: 0, valign: "top", isTextBox: true });
    }
    footer(s, 6);
    s.addNotes("Aligned with the CBAM verification requirements in the implementing acts and with ISO 14065 / ISO 14064-3 principles. Verification is performed to reasonable assurance against a 5% materiality threshold. Typical duration depends on installation complexity, usually a few weeks from receipt of complete data.");
  }

  // ================= 7. What we need =================
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    kicker(s, "Preparing for verification");
    title(s, "What we will need from you");

    const left = [
      "Installation description, boundaries and production processes",
      "Monitoring methodology / monitoring plan used for CBAM",
      "Activity data: fuels, materials, production volumes by good",
      "Measurement records, meter calibration and lab analyses",
    ];
    const right = [
      "Purchased precursors and their embedded emissions (with evidence)",
      "Electricity consumption and emission factor sources",
      "Emissions report / CBAM communication template for the period",
      "Carbon price paid in the country of origin, if any",
    ];
    const list = (arr) => arr.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < arr.length - 1 } }));
    s.addText(list(left), { x: 0.6, y: 1.65, w: 5.5, h: 3.6, fontFace: BODY, fontSize: 16, color: C.ink, paraSpaceAfter: 14, valign: "top", margin: 0, isTextBox: true });
    s.addText(list(right), { x: 6.5, y: 1.65, w: 5.5, h: 3.6, fontFace: BODY, fontSize: 16, color: C.ink, paraSpaceAfter: 14, valign: "top", margin: 0, isTextBox: true });

    s.addShape("roundRect", { x: 0.6, y: 5.45, w: W - 1.2, h: 1.1, ...card(C.mint) });
    await badge(s, fa.FaLightbulb, 0.85, 5.62, 0.75, C.amber, C.deep);
    s.addText("Tip: having data organised by installation and production process before we begin shortens the verification and reduces cost.", {
      x: 1.85, y: 5.55, w: 10.5, h: 0.9, fontFace: BODY, fontSize: 15, italic: true, color: C.deep, margin: 0, valign: "middle", isTextBox: true,
    });
    footer(s, 7);
    s.addNotes("We share a detailed document request list at engagement. Where operators have used the Commission's communication template for installations, this is a good starting point.");
  }

  // ================= 8. Why us =================
  {
    const s = pres.addSlide();
    s.background = { color: C.deep };
    kicker(s, "Why CleanCarbon Assurance", C.amber);
    title(s, "A verification partner built for CBAM", { color: C.white });
    const items = [
      [fa.FaBullseye, "CBAM specialists", "CBAM is our core focus, not a side line to general certification."],
      [fa.FaUserTie, "Qualified lead verifiers", "Led by ISO 14064 lead verifiers with hands-on GHG assurance experience in heavy industry."],
      [fa.FaBalanceScale, "Strict impartiality", "We don't consult on the data we verify, so EU buyers can trust our opinion."],
      [fa.FaMapMarkerAlt, "Close to exporters", "India-based team: responsive site visits, local context, lower travel cost."],
      [fa.FaEuroSign, "EU-aligned method", "Work follows the EU CBAM implementing rules, supported by ISO 14064-3 and ISO 14065 principles."],
      [fa.FaComments, "Clear communication", "Plain-language findings, predictable timelines and transparent fixed-fee proposals."],
    ];
    const cw = 3.85, ch = 2.2, gx = 0.3, gy = 0.3, x0 = 0.6, y0 = 1.6;
    for (let i = 0; i < items.length; i++) {
      const c = i % 3, r = Math.floor(i / 3);
      const x = x0 + c * (cw + gx), y = y0 + r * (ch + gy);
      s.addShape("roundRect", { x, y, w: cw, h: ch, fill: { color: C.pine }, line: { color: C.pine }, rectRadius: 0.12 });
      await badge(s, items[i][0], x + 0.3, y + 0.3, 0.7, C.amber, C.deep);
      s.addText(items[i][1], { x: x + 1.2, y: y + 0.3, w: cw - 1.4, h: 0.7, fontFace: HEAD, fontSize: 18, bold: true, color: C.white, margin: 0, valign: "middle", isTextBox: true });
      s.addText(items[i][2], { x: x + 0.3, y: y + 1.15, w: cw - 0.6, h: 0.95, fontFace: BODY, fontSize: 13, color: "D3E8E1", margin: 0, valign: "top", isTextBox: true });
    }
    footer(s, 8, true);
    s.addNotes("Emphasise specialist focus, the lead verifier's ISO 14064 credentials, and our impartiality commitment.");
  }

  // ================= 9. Accreditation & getting started =================
  {
    const s = pres.addSlide();
    s.background = { color: C.white };
    kicker(s, "Accreditation & next steps");
    title(s, "Plan your 2026 verification now");

    // Accreditation status card
    s.addShape("roundRect", { x: 0.6, y: 1.6, w: 5.0, h: 4.4, ...card(C.mist) });
    await badge(s, fa.FaAward, 0.95, 1.95, 0.95, C.emerald, C.white);
    s.addText("Accreditation status", { x: 0.95, y: 3.1, w: 4.3, h: 0.5, fontFace: HEAD, fontSize: 20, bold: true, color: C.deep, margin: 0, isTextBox: true });
    s.addText("CleanCarbon Assurance has applied for accreditation as a CBAM verifier, and the accreditation process is under way. We will confirm status and scope in writing in every proposal and before any verification opinion is issued.", {
      x: 0.95, y: 3.65, w: 4.35, h: 2.6, fontFace: BODY, fontSize: 14, color: C.ink, margin: 0, valign: "top", isTextBox: true,
    });

    // Engagement steps
    const steps = [
      ["Discovery call", "We learn about your installation(s), CBAM goods and EU customers."],
      ["Scoping & proposal", "A fixed-fee proposal with timeline and document request list."],
      ["Readiness review (optional)", "Early gap check so data is ready before the verification window."],
      ["Verification", "Planned around your production calendar, well before your customers' 30 Sep 2027 declaration deadline."],
    ];
    const x = 6.1;
    for (let i = 0; i < steps.length; i++) {
      const y = 1.65 + i * 1.22;
      s.addShape("ellipse", { x, y, w: 0.6, h: 0.6, fill: { color: i === 0 ? C.amber : C.deep }, line: { color: i === 0 ? C.amber : C.deep } });
      s.addText(String(i + 1), { x, y, w: 0.6, h: 0.6, align: "center", valign: "middle", fontFace: BODY, fontSize: 16, bold: true, color: i === 0 ? C.deep : C.white, margin: 0, isTextBox: true });
      s.addText(steps[i][0], { x: x + 0.85, y: y - 0.02, w: 5.7, h: 0.4, fontFace: HEAD, fontSize: 18, bold: true, color: C.deep, margin: 0, isTextBox: true });
      s.addText(steps[i][1], { x: x + 0.85, y: y + 0.38, w: 5.7, h: 0.7, fontFace: BODY, fontSize: 14, color: C.ink, margin: 0, valign: "top", isTextBox: true });
    }
    footer(s, 9);
    s.addNotes("Be transparent: accreditation has been applied for and is in progress. Early engagement lets exporters reserve verification slots ahead of the expected capacity crunch in the first half of 2027.");
  }

  // ================= 10. Contact =================
  {
    const s = pres.addSlide();
    s.background = { color: C.deep };
    s.addShape("ellipse", { x: 8.9, y: -1.2, w: 6.0, h: 6.0, fill: { color: C.pine }, line: { color: C.pine } });
    s.addShape("ellipse", { x: 10.6, y: 4.3, w: 3.6, h: 3.6, fill: { color: C.emerald, transparency: 55 }, line: { color: C.emerald, transparency: 55 } });
    await badge(s, fa.FaLeaf, 10.55, 1.2, 2.0, C.amber, C.deep);

    s.addText("LET'S TALK", { x: 0.8, y: 0.9, w: 6, h: 0.4, fontFace: BODY, fontSize: 14, bold: true, charSpacing: 4, color: C.amber, margin: 0, isTextBox: true });
    s.addText("Get your CBAM data verified", { x: 0.8, y: 1.35, w: 7.8, h: 1.4, fontFace: HEAD, fontSize: 36, valign: "top", bold: true, color: C.white, margin: 0, isTextBox: true });

    // Contact card
    s.addShape("roundRect", { x: 0.8, y: 3.0, w: 7.6, h: 3.6, fill: { color: C.pine }, line: { color: C.pine }, rectRadius: 0.12 });
    await badge(s, fa.FaUserTie, 1.15, 3.35, 0.95, C.amber, C.deep);
    s.addText("Kashyap Trivedi", { x: 2.35, y: 3.35, w: 5.8, h: 0.5, fontFace: HEAD, fontSize: 24, bold: true, color: C.white, margin: 0, isTextBox: true });
    s.addText("Lead Verifier (ISO 14064)", { x: 2.35, y: 3.85, w: 5.8, h: 0.4, fontFace: BODY, fontSize: 15, color: C.amber, margin: 0, isTextBox: true });

    const lines = [
      [fa.FaEnvelope, "kashyap@cbamassure.com"],
      [fa.FaPhoneAlt, "+91 83470008344"],
      [fa.FaInbox, "nb@cbamassure.com  (new enquiries)"],
      [fa.FaGlobe, "cbamassure.com"],
    ];
    for (let i = 0; i < lines.length; i++) {
      const y = 4.6 + i * 0.45;
      s.addImage({ data: await icon(lines[i][0], "9FC3B8"), x: 1.2, y: y + 0.07, w: 0.26, h: 0.26 });
      s.addText(lines[i][1], { x: 1.65, y, w: 6.5, h: 0.4, fontFace: BODY, fontSize: 16, color: C.white, margin: 0, valign: "middle", isTextBox: true });
    }
    s.addNotes("Close with a clear call to action: book a discovery call via nb@cbamassure.com or contact Kashyap directly.");
  }

  const out = path.join(__dirname, "CleanCarbon_Assurance_CBAM_Verification.pptx");
  await pres.writeFile({ fileName: out });
  console.log("wrote", out);
})();
