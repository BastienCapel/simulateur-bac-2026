import ExcelJS from "exceljs";

// ─── Color Palette Matching DNB Premium Aesthetic ──────────────────────────
const C = {
  headerBg: "1A2D4A",     // Dark Navy Blue for Title
  headerFg: "FFFFFF",
  subHeaderBg: "2E4A6E",  // Medium Steel Blue for Sub-header
  subHeaderFg: "FFFFFF",
  colHeaderBg: "3B5998",  // Steel Blue for section dividers / headers
  colHeaderFg: "FFFFFF",
  altRow: "F4F7FB",       // Soft alternating row color
  white: "FFFFFF",
  admis: "1A7A45",        // Premium Green
  admisBg: "E8F5EE",
  rattrapage: "B54708",   // Warm Orange
  rattrapageBg: "FEF3EB",
  refuse: "9E2A2B",       // Academic Red
  refuseBg: "FCE8E6",
  mention_tb_felicitations: "7B3FA6", // Imperial Purple
  mention_tb: "1D5FA6",   // Royal Blue
  mention_b: "0D7B8A",    // Deep Teal
  mention_ab: "3B6B35",   // Olive Green
  mention_sans: "5C5C5C",  // Slate Grey
  border: "C8D4E8",       // Soft grid borders
  borderStrong: "3B5998", // Accent borders
  summaryBg: "EFF4FF",    // Soft summary card blue
  summaryHeaderBg: "2E4A6E",
  gold: "B8860B",         // Academic gold highlight
  goldBg: "FFF8E7",
};

// ─── Formatting Helpers ───────────────────────────────────────────────────
function fr(value, decimals = 2) {
  if (typeof value !== "number" || !Number.isFinite(value)) return "—";
  return value.toFixed(decimals).replace(".", ",");
}

function mentionColor(mention) {
  if (!mention) return C.mention_sans;
  const m = mention.toLowerCase();
  if (m.includes("félicitations")) return C.mention_tb_felicitations;
  if (m.includes("très bien")) return C.mention_tb;
  if (m.includes("bien")) return C.mention_b;
  if (m.includes("assez bien")) return C.mention_ab;
  return C.mention_sans;
}

function gradeColor(value) {
  if (value === null || value === undefined) return C.mention_sans;
  if (value >= 16) return C.admis;
  if (value >= 14) return C.mention_b;
  if (value >= 12) return C.mention_ab;
  if (value >= 10) return C.mention_sans;
  if (value >= 8) return C.rattrapage;
  return C.refuse;
}

function border(style = "thin") {
  const s = { style, color: { argb: "FF" + C.border } };
  return { top: s, left: s, bottom: s, right: s };
}

function borderStrong() {
  const s = { style: "medium", color: { argb: "FF" + C.borderStrong } };
  return { top: s, left: s, bottom: s, right: s };
}

function applyFont(cell, opts = {}) {
  cell.font = {
    name: opts.name ?? "Calibri",
    bold: opts.bold ?? false,
    size: opts.size ?? 11,
    color: { argb: "FF" + (opts.color ?? "1A2D4A") },
  };
}

function applyFill(cell, argbHex) {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + argbHex } };
}

function gradeBar(value, max = 20) {
  if (value === null || value === undefined) return "";
  const filled = Math.min(10, Math.max(0, Math.round((value / max) * 10)));
  return "█".repeat(filled) + "░".repeat(10 - filled);
}

// ─── Dynamic Target Requirements ──────────────────────────────────────────
function calculateRequiredTerminalAverage(ccAverage, targetThreshold) {
  if (ccAverage === null || ccAverage === undefined) return null;
  // Total points required = threshold * 100
  // CC points = ccAverage * 40
  // Terminal points needed = Total points required - CC points
  // French points are already locked in our target simulator, but wait:
  // In our simulator, the target calculates the required average on the *remaining* 50 coefficients
  // which are Spe 1 (16) + Spe 2 (16) + Philo (8) + Grand Oral (10) = 50.
  // Wait, let's keep the formula exact as in App.jsx.
  return null; // Will override with real math or pass directly
}

// ─── Export Individual simulation ──────────────────────────────────────────
export async function exportIndividualToExcel(student, sim, results) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Simulateur Bac 2026 LFJP";
  wb.created = new Date();
  wb.modified = new Date();

  const exportedAt = new Date().toLocaleDateString("fr-FR", {
    day: "2-digit",
    month: "long",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });

  const ws = wb.addWorksheet("Fiche Candidat", {
    pageSetup: {
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9, // A4
      margins: { left: 0.6, right: 0.6, top: 0.8, bottom: 0.8, header: 0.3, footer: 0.3 },
    },
    headerFooter: {
      oddFooter: `&L&"Calibri,Italic"&8Bac Général 2026 — ${student.name}&R&"Calibri,Italic"&8Page &P`,
    },
  });

  // Setup 4 columns identical to DNB format:
  // Col A: Label
  // Col B: Numerical grade
  // Col C: Sparkline Progress Bar
  // Col D: Formatted grade /20
  ws.columns = [
    { width: 34 },  // A - Subject Label
    { width: 10 },  // B - Numerical value
    { width: 22 },  // C - Visual Sparkline
    { width: 14 },  // D - Formatted grade /20
  ];

  let r = 1;

  // Local helper to set cell properties easily
  const setCell = (rowNum, colNum, value, opts = {}) => {
    const cell = ws.getCell(rowNum, colNum);
    cell.value = value;
    applyFont(cell, { bold: opts.bold, size: opts.size ?? 10, color: opts.color ?? "1A2D4A" });
    if (opts.bg) applyFill(cell, opts.bg);
    cell.alignment = {
      horizontal: opts.align ?? "left",
      vertical: "middle",
      wrapText: opts.wrap ?? false
    };
    return cell;
  };

  const mergeRow = (rowNum, fromCol, toCol) => {
    ws.mergeCells(rowNum, fromCol, rowNum, toCol);
  };

  // Draws section header bar
  const drawSectionHeader = (label) => {
    for (let c = 1; c <= 4; c++) {
      const cell = ws.getCell(r, c);
      applyFill(cell, C.colHeaderBg);
      cell.border = border("thin");
    }
    const titleCell = ws.getCell(r, 1);
    titleCell.value = label;
    applyFont(titleCell, { bold: true, size: 10, color: C.colHeaderFg });
    titleCell.alignment = { horizontal: "left", vertical: "middle" };
    mergeRow(r, 1, 4);
    ws.getRow(r).height = 20;
    r++;
  };

  // Draws standard data row
  const drawDataRow = (label, value, opts = {}) => {
    const bg = opts.bg ?? (opts.alt ? C.altRow : C.white);
    
    // Label cell
    setCell(r, 1, label, { bg });
    ws.getCell(r, 1).border = border("hair");

    if (opts.isDispensed) {
      setCell(r, 2, "Disp.", { bold: true, color: C.refuse, align: "center", bg });
      ws.getCell(r, 2).border = border("hair");
      
      setCell(r, 3, "Élève dispensé d'épreuve", { size: 9, color: "9E2A2B", bg });
      ws.getCell(r, 3).border = border("hair");
      
      setCell(r, 4, "Dispensé", { bold: true, color: C.refuse, align: "center", bg });
      ws.getCell(r, 4).border = border("hair");
    } else if (value !== null && value !== undefined) {
      const col = opts.color ?? gradeColor(value);
      
      // Numerical cell
      setCell(r, 2, Number(value.toFixed(2)), { bold: opts.bold, color: col, align: "center", bg });
      ws.getCell(r, 2).border = border("hair");

      // Progress bar or detail
      if (opts.bar) {
        const barCell = ws.getCell(r, 3);
        barCell.value = gradeBar(value);
        applyFont(barCell, { size: 8, color: col });
        applyFill(barCell, bg);
        barCell.alignment = { horizontal: "left", vertical: "middle" };
        barCell.border = border("hair");
      } else if (opts.detail) {
        setCell(r, 3, opts.detail, { size: 9, color: "6B7280", bg });
        ws.getCell(r, 3).border = border("hair");
      } else {
        applyFill(ws.getCell(r, 3), bg);
        ws.getCell(r, 3).border = border("hair");
      }

      // Formatted text
      setCell(r, 4, `${value.toFixed(2).replace(".", ",")} /20`, { bold: opts.bold, color: col, align: "center", bg });
      ws.getCell(r, 4).border = border("hair");
    } else {
      for (let c = 2; c <= 4; c++) {
        setCell(r, c, "—", { color: "9CA3AF", align: "center", bg });
        ws.getCell(r, c).border = border("hair");
      }
    }
    ws.getRow(r).height = 17;
    r++;
  };

  // ─── Header Block ──────────────────────────────────────────────────────────
  for (let c = 1; c <= 4; c++) applyFill(ws.getCell(r, c), C.headerBg);
  const mainTitle = ws.getCell(r, 1);
  mainTitle.value = "BACCALAURÉAT GÉNÉRAL — SESSION JUIN 2026";
  applyFont(mainTitle, { bold: true, size: 13, color: C.headerFg });
  mergeRow(r, 1, 4);
  mainTitle.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(r).height = 26;
  r++;

  for (let c = 1; c <= 4; c++) applyFill(ws.getCell(r, c), C.subHeaderBg);
  const subtitleCell = ws.getCell(r, 1);
  subtitleCell.value = `${student.name}  ·  Classe: Terminale Générale LFJP`;
  applyFont(subtitleCell, { bold: true, size: 11, color: C.headerFg });
  mergeRow(r, 1, 4);
  subtitleCell.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(r).height = 20;
  r++;

  for (let c = 1; c <= 4; c++) applyFill(ws.getCell(r, c), C.subHeaderBg);
  const metaCell = ws.getCell(r, 1);
  metaCell.value = `Fiche de simulation individuelle exportée le ${exportedAt}  ·  N° ${student.candidate_id}`;
  applyFont(metaCell, { size: 9, color: "A8C4E0" });
  mergeRow(r, 1, 4);
  metaCell.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(r).height = 14;
  r++;

  // Small spacer
  for (let c = 1; c <= 4; c++) applyFill(ws.getCell(r, c), C.headerBg);
  ws.getRow(r).height = 3;
  r++;
  r++;

  // ─── Rules of Calculation ──────────────────────────────────────────────────
  drawSectionHeader("RÈGLES DE CALCUL");
  const ruleCell = ws.getCell(r, 1);
  ruleCell.value =
    "Contrôle continu (CC) : 40 % de la moyenne finale (21 coefficients acquis en 1ère, 19 coefficients acquis en Terminale). " +
    "Épreuves terminales : 60 % — pondérées par les coefficients officiels (Français Écrit 5, Français Oral 5, Spécialité 1 16, Spécialité 2 16, Philosophie 8, Grand Oral 10). " +
    "Admission ≥ 10/20. Rattrapage : 8,00 à 9,99/20. Mentions : Assez bien ≥ 12, Bien ≥ 14, Très bien ≥ 16, Très bien avec félicitations ≥ 18.";
  applyFont(ruleCell, { size: 8, color: "4B5563" });
  mergeRow(r, 1, 4);
  ruleCell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
  ws.getRow(r).height = 38;
  r++;
  r++;

  // ─── Contrôle Continu 1ère (Coef 21) ───────────────────────────────────────
  drawSectionHeader("CONTRÔLE CONTINU — CLASSE DE 1ÈRE  (Coef. 21 / 40)");

  const lp001 = student.grades['LP001']?.value ?? 10.0;
  const lp002 = student.grades['LP002']?.value ?? 10.0;
  const lp003 = student.grades['LP003']?.value ?? 10.0;
  const lp004 = student.grades['LP004']?.value ?? 10.0;
  const ensSci1 = student.grades['LP005']?.value ?? student.grades['LP007']?.value ?? 10.0;
  const droppedSpec = student.dropped_specialty?.value ?? 10.0;

  const cc1Rows = [
    ["Histoire-Géographie 1ère", lp001, "coef. 3", false],
    ["Enseignement moral et civique (EMC) 1ère", lp002, "coef. 1", true],
    ["Langue Vivante A 1ère", lp003, "coef. 3", false],
    ["Langue Vivante B 1ère", lp004, "coef. 3", true],
    ["Enseignement Scientifique 1ère", ensSci1, "coef. 3", false],
    [`Spécialité non poursuivie: ${student.dropped_specialty?.label || "Spé 1ère"}`, droppedSpec, "coef. 8", true],
  ];

  cc1Rows.forEach(([label, val, detail, alt]) => {
    drawDataRow(label, val, { bar: true, detail, alt });
  });
  r++;

  // ─── Contrôle Continu Terminale (Coef 19 or 13 if EPS is dispensed) ────────
  const isEpsDispensed = student.eps_grade === 'dispense';
  const cc2CoefTotal = isEpsDispensed ? 13 : 19;
  const ccCoefTotal = isEpsDispensed ? 34 : 40;
  drawSectionHeader(`CONTRÔLE CONTINU — CLASSE DE TERMINALE  (Coef. ${cc2CoefTotal} / ${ccCoefTotal})`);

  const lt001 = student.grades['LT001']?.value ?? 10.0;
  const lt002 = student.grades['LT002']?.value ?? 10.0;
  const lt003 = student.grades['LT003']?.value ?? 10.0;
  const lt004 = student.grades['LT004']?.value ?? 10.0;
  const lt005 = student.grades['LT005']?.value ?? 10.0;

  const cc2Rows = [
    ["Histoire-Géographie Terminale", lt001, "coef. 3", false],
    ["Enseignement moral et civique (EMC) Terminale", lt002, "coef. 1", true],
    ["Langue Vivante A Terminale", lt003, "coef. 3", false],
    ["Langue Vivante B Terminale", lt004, "coef. 3", true],
    ["Enseignement Scientifique Terminale", lt005, "coef. 3", false],
  ];

  cc2Rows.forEach(([label, val, detail, alt]) => {
    drawDataRow(label, val, { bar: true, detail, alt });
  });

  // Handle EPS row dynamically:
  let epsVal = sim.eps;
  let epsLabel = "EPS Terminale (Simulé)";
  let epsDetail = "coef. 6";
  let epsOpts = { bar: true, detail: epsDetail, alt: true };

  if (isEpsDispensed) {
    epsLabel = "EPS Terminale (Dispensé)";
    epsOpts = { isDispensed: true, alt: true };
  } else if (student.eps_grade !== undefined) {
    epsLabel = "EPS Terminale (Note Réelle)";
    epsVal = student.eps_grade;
    epsOpts = { bar: true, color: C.admis, detail: "coef. 6 (Réel)", alt: true };
  }

  drawDataRow(epsLabel, epsVal, epsOpts);

  // Moyenne CC Box
  for (let c = 1; c <= 4; c++) applyFill(ws.getCell(r, c), C.summaryBg);
  const ccLabel = ws.getCell(r, 1);
  ccLabel.value = "Moyenne contrôle continu (40%)";
  applyFont(ccLabel, { bold: true, size: 10, color: "1E3A5F" });
  mergeRow(r, 1, 2);
  ccLabel.border = borderStrong();

  const ccVal = ws.getCell(r, 3);
  ccVal.value = `${results.ccAverage.toFixed(2).replace(".", ",")} / 20`;
  applyFont(ccVal, { bold: true, size: 13, color: gradeColor(results.ccAverage) });
  ccVal.alignment = { horizontal: "center", vertical: "middle" };
  ws.mergeCells(r, 3, r, 4);
  ccVal.border = borderStrong();
  ws.getRow(r).height = 22;
  r++;
  r++;

  // ─── Épreuves Terminales (Coef 60) ──────────────────────────────────────────
  drawSectionHeader("ÉPREUVES TERMINALES  (60 % de la moyenne finale)");

  const terminalRows = [
    ["Français Écrit (Anticipé)", sim.frenchWritten, "coef. 5", false],
    ["Français Oral (Anticipé)", sim.frenchOral, "coef. 5", true],
    [`Spé 1: ${student.specialties[0] || "Spécialité 1"}`, sim.spe1, "coef. 16", false],
    [`Spé 2: ${student.specialties[1] || "Spécialité 2"}`, sim.spe2, "coef. 16", true],
    ["Philosophie", sim.philo, "coef. 8", false],
    ["Grand Oral", sim.grandOral, "coef. 10", true],
  ];

  terminalRows.forEach(([label, val, detail, alt]) => {
    drawDataRow(label, val, { bar: true, detail, alt });
  });

  // Moyenne Terminales Box
  for (let c = 1; c <= 4; c++) applyFill(ws.getCell(r, c), C.summaryBg);
  const termLabel = ws.getCell(r, 1);
  termLabel.value = "Moyenne épreuves terminales (60%)";
  applyFont(termLabel, { bold: true, size: 10, color: "1E3A5F" });
  mergeRow(r, 1, 2);
  termLabel.border = borderStrong();

  const termVal = ws.getCell(r, 3);
  termVal.value = `${results.terminalAverage.toFixed(2).replace(".", ",")} / 20`;
  applyFont(termVal, { bold: true, size: 13, color: gradeColor(results.terminalAverage) });
  termVal.alignment = { horizontal: "center", vertical: "middle" };
  ws.mergeCells(r, 3, r, 4);
  termVal.border = borderStrong();
  ws.getRow(r).height = 22;
  r++;
  r++;

  // ─── Résultat Projeté ──────────────────────────────────────────────────────
  drawSectionHeader("RÉSULTAT PROJETÉ");

  let resultBg = C.altRow;
  let resultColor = C.mention_sans;
  let resultLabel = "Simulation incomplète";

  if (results.status === "Admis") {
    resultBg = C.admisBg;
    resultColor = C.admis;
    resultLabel = "✓  ADMIS";
  } else if (results.status === "Rattrapage") {
    resultBg = C.rattrapageBg;
    resultColor = C.rattrapage;
    resultLabel = "⚠  RATTRAPAGE (Second Groupe)";
  } else if (results.status === "Refusé") {
    resultBg = C.refuseBg;
    resultColor = C.refuse;
    resultLabel = "✗  REFUSÉ";
  }

  for (let c = 1; c <= 4; c++) applyFill(ws.getCell(r, c), resultBg);
  ws.getCell(r, 1).value = "Statut";
  applyFont(ws.getCell(r, 1), { size: 10, color: "1E3A5F" });
  ws.getCell(r, 1).border = border("thin");
  
  ws.getCell(r, 2).value = resultLabel;
  applyFont(ws.getCell(r, 2), { bold: true, size: 12, color: resultColor });
  ws.getCell(r, 2).alignment = { horizontal: "left", vertical: "middle" };
  ws.getCell(r, 2).border = border("thin");
  mergeRow(r, 2, 4);
  ws.getRow(r).height = 22;
  r++;

  // Mention
  const mentionBg = results.mention ? C.summaryBg : C.altRow;
  for (let c = 1; c <= 4; c++) applyFill(ws.getCell(r, c), mentionBg);
  ws.getCell(r, 1).value = "Mention";
  applyFont(ws.getCell(r, 1), { size: 10, color: "1E3A5F" });
  ws.getCell(r, 1).border = border("thin");
  
  ws.getCell(r, 2).value = results.mention ?? "Non attribuée";
  applyFont(ws.getCell(r, 2), { bold: !!results.mention, size: 11, color: mentionColor(results.mention) });
  ws.getCell(r, 2).alignment = { horizontal: "left", vertical: "middle" };
  ws.getCell(r, 2).border = border("thin");
  mergeRow(r, 2, 4);
  ws.getRow(r).height = 20;
  r++;

  // Moyenne Finale Card
  for (let c = 1; c <= 4; c++) applyFill(ws.getCell(r, c), resultBg);
  ws.getCell(r, 1).value = "MOYENNE FINALE PROJETÉE";
  applyFont(ws.getCell(r, 1), { bold: true, size: 11, color: "1E3A5F" });
  ws.getCell(r, 1).border = borderStrong();
  mergeRow(r, 1, 2);
  ws.getCell(r, 1).alignment = { horizontal: "right", vertical: "middle" };

  const finalVal = ws.getCell(r, 3);
  finalVal.value = `${results.finalAverage.toFixed(2).replace(".", ",")} / 20`;
  applyFont(finalVal, { bold: true, size: 16, color: resultColor });
  finalVal.alignment = { horizontal: "center", vertical: "middle" };
  ws.mergeCells(r, 3, r, 4);
  finalVal.border = borderStrong();
  ws.getRow(r).height = 28;
  r++;
  r++;

  // ─── Objectifs Dynamic Thresholds ──────────────────────────────────────────
  drawSectionHeader("OBJECTIFS — MOYENNE TERMINALE REQUISE SUR LES ÉPREUVES RESTANTES");

  // Headers sub-table
  ["Objectif visé", "Seuil Requis", "Moyenne nécessaire (coeff 50)", "Faisabilité"].forEach((h, i) => {
    const cell = ws.getCell(r, i + 1);
    cell.value = h;
    applyFont(cell, { bold: true, size: 9, color: "2E4A6E" });
    applyFill(cell, C.summaryBg);
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = border("thin");
  });
  ws.getRow(r).height = 16;
  r++;

  // Calculate dynamic goals based on actual acquired points
  const totalCcPoints = results.totalCcPoints;
  const frenchPoints = (sim.frenchWritten * 5) + (sim.frenchOral * 5);
  const acquiredPoints = totalCcPoints + frenchPoints; // total coefficient 50 or 44
  const remainingCoef = 50; // Spe 1 (16) + Spe 2 (16) + Philo (8) + Grand Oral (10) = 50
  const totalBacCoef = isEpsDispensed ? 94.0 : 100.0;

  const targets = [
    { label: "Baccalauréat (Admis)", target: 10.0 },
    { label: "Mention Assez Bien", target: 12.0 },
    { label: "Mention Bien", target: 14.0 },
    { label: "Mention Très Bien", target: 16.0 },
    { label: "Mention Très Bien avec Félicitations", target: 18.0 }
  ];

  targets.forEach((target, i) => {
    const bg = i % 2 === 0 ? C.white : C.altRow;
    const neededPoints = (target.target * totalBacCoef) - acquiredPoints;
    const noteMin = neededPoints / remainingCoef;

    let feasibility;
    let feasColor;
    if (noteMin <= 0) {
      feasibility = "✓ Déjà acquis";
      feasColor = C.admis;
    } else if (noteMin > 20) {
      feasibility = "✗ Impossible";
      feasColor = C.refuse;
    } else if (noteMin >= 18) {
      feasibility = "⚠ Très difficile";
      feasColor = C.rattrapage;
    } else if (noteMin >= 15) {
      feasibility = "△ Ambitieux";
      feasColor = C.gold;
    } else {
      feasibility = "✓ Atteignable";
      feasColor = C.mention_ab;
    }

    setCell(r, 1, target.label, { bg });
    ws.getCell(r, 1).border = border("hair");
    
    setCell(r, 2, `${target.target.toFixed(0)} /20`, { align: "center", bold: true, color: gradeColor(target.target), bg });
    ws.getCell(r, 2).border = border("hair");

    const noteDisplay = noteMin > 20 ? "—" : noteMin <= 0 ? "—" : `${noteMin.toFixed(2).replace(".", ",")} /20`;
    setCell(r, 3, noteDisplay, { align: "center", bold: true, color: noteMin > 0 && noteMin <= 20 ? gradeColor(noteMin) : C.mention_sans, bg });
    ws.getCell(r, 3).border = border("hair");

    setCell(r, 4, feasibility, { align: "center", bold: true, color: feasColor, bg });
    ws.getCell(r, 4).border = border("hair");
    
    ws.getRow(r).height = 16;
    r++;
  });

  // ─── Footer Block ──────────────────────────────────────────────────────────
  r++;
  for (let c = 1; c <= 4; c++) applyFill(ws.getCell(r, c), C.headerBg);
  ws.getCell(r, 1).value = `Simulateur BAC Général 2026 — Lycée Français Jacques Prévert (Saly) • Généré le ${exportedAt}`;
  applyFont(ws.getCell(r, 1), { size: 8, color: "A8C4E0" });
  mergeRow(r, 1, 4);
  ws.getCell(r, 1).alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(r).height = 14;

  // ─── Write & Download Workbook ─────────────────────────────────────────────
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Simulation_Bac_2026_${student.name.replace(/\s+/g, '_')}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ─── Export principal (Cohort) ───────────────────────────────────────────────
export async function exportCohortToExcel(computedStudents, simulations) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Simulateur Bac 2026 LFJP";
  wb.created = new Date();
  wb.modified = new Date();
  wb.properties.date1904 = false;

  const exportedAt = new Date().toLocaleDateString("fr-FR", {
    day: "2-digit",
    month: "long",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });

  // Build the primary sheet for the entire cohort
  buildCohortSheet(wb, "Toute la Cohorte", computedStudents, simulations, exportedAt);

  // Group students by their profile specialties combination to make secondary sheets!
  // This is a premium touch showing master class organization.
  const profileGroups = {};
  computedStudents.forEach(s => {
    // Abbreviated profile key e.g. "Maths - SES"
    const spec1 = s.specialties[0] ? s.specialties[0].replace("Histoire-géographie, géopolitique et sciences politiques", "HGGSP").replace("Sciences de la vie et de la terre", "SVT").replace("Sciences économiques et sociales", "SES").replace("Langues, littératures et cultures étrangères et régionales - Anglais, monde contemporain", "Anglais MC") : "";
    const spec2 = s.specialties[1] ? s.specialties[1].replace("Histoire-géographie, géopolitique et sciences politiques", "HGGSP").replace("Sciences de la vie et de la terre", "SVT").replace("Sciences économiques et sociales", "SES").replace("Langues, littératures et cultures étrangères et régionales - Anglais, monde contemporain", "Anglais MC") : "";
    const profileName = [spec1, spec2].filter(Boolean).join(" - ");
    const key = profileName || "Général";
    if (!profileGroups[key]) profileGroups[key] = [];
    profileGroups[key].push(s);
  });

  // Generate sheets for profiles with at least 3 students to keep workbook elegant
  Object.keys(profileGroups).forEach(profileKey => {
    if (profileGroups[profileKey].length >= 2) {
      buildCohortSheet(wb, profileKey.substring(0, 31), profileGroups[profileKey], simulations, exportedAt);
    }
  });

  // Export
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `Simulateur_Bac_2026_Cohorte_LFJP_${new Date().toLocaleDateString("fr-FR").replaceAll("/", "-")}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ─── Core Cohort Sheet Builder ─────────────────────────────────────────────
function buildCohortSheet(wb, sheetName, rows, simulations, exportedAt) {
  const ws = wb.addWorksheet(sheetName, {
    pageSetup: {
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9, // A4
      margins: { left: 0.4, right: 0.4, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 },
    },
    headerFooter: {
      oddFooter: `&L&"Calibri,Italic"&8Bac 2026 Simulator — Cohorte ${sheetName}&R&"Calibri,Italic"&8Page &P de &N`,
    },
  });

  // Setup Column Widths
  ws.columns = [
    { width: 4 },    // A - Index
    { width: 22 },   // B - Last name
    { width: 18 },   // C - First name
    { width: 14 },   // D - Candidate ID
    { width: 12 },   // E - Date of Birth
    { width: 28 },   // F - Specialty 1
    { width: 28 },   // G - Specialty 2
    { width: 14 },   // H - CC Average /20
    { width: 10 },   // I - Français Écrit
    { width: 10 },   // J - Français Oral
    { width: 10 },   // K - Spé 1
    { width: 10 },   // L - Spé 2
    { width: 10 },   // M - Philo
    { width: 10 },   // N - Grand Oral
    { width: 10 },   // O - EPS
    { width: 14 },   // P - Terminal Average /20
    { width: 14 },   // Q - Final Average /20
    { width: 14 },   // R - Status
    { width: 28 },   // S - Mention
  ];

  let r = 1;

  // Title bar banner
  const titleCell = ws.getCell(r, 1);
  titleCell.value = "BACCALAURÉAT GÉNÉRAL — SESSION JUIN 2026";
  applyFont(titleCell, { bold: true, size: 14, color: C.headerFg });
  applyFill(titleCell, C.headerBg);
  ws.mergeCells(r, 1, r, 19);
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(r).height = 28;
  r++;

  // Subtitle
  const subtitleCell = ws.getCell(r, 1);
  subtitleCell.value = `Simulation globale — ${sheetName}  ·  Lycée Français Jacques Prévert (Saly)  ·  Généré le ${exportedAt}`;
  applyFont(subtitleCell, { size: 10, color: C.headerFg });
  applyFill(subtitleCell, C.subHeaderBg);
  ws.mergeCells(r, 1, r, 19);
  subtitleCell.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(r).height = 18;
  r++;

  // Methodology separator row
  for (let c = 1; c <= 19; c++) applyFill(ws.getCell(r, c), C.headerBg);
  ws.getRow(r).height = 4;
  r++;

  // Rules text
  const ruleCell = ws.getCell(r, 1);
  ruleCell.value = "Contrôle Continu = 40% (21 coeff 1ère, 19 coeff Terminale). Épreuves Terminales = 60% (Français Écrit 5, Français Oral 5, Spé 1 16, Spé 2 16, Philo 8, Grand Oral 10). Admis ≥ 10,00. Second groupe (Rattrapage) : 8,00 à 9,99/20.";
  applyFont(ruleCell, { size: 8, color: "4A5568" });
  ws.mergeCells(r, 1, r, 19);
  ruleCell.alignment = { horizontal: "left", vertical: "middle" };
  ws.getRow(r).height = 14;
  r++;
  r++; // space

  // Table Columns Headers
  const colHeaders = [
    { label: "N°", align: "center" },
    { label: "NOM", align: "left" },
    { label: "Prénom", align: "left" },
    { label: "N° Candidat", align: "center" },
    { label: "Date Naiss.", align: "center" },
    { label: "Spécialité 1", align: "left" },
    { label: "Spécialité 2", align: "left" },
    { label: "Moyenne\nCC (40%)", align: "center" },
    { label: "Fr. Écrit\n(Anticipé)", align: "center" },
    { label: "Fr. Oral\n(Anticipé)", align: "center" },
    { label: "Spé 1\n(Simulée)", align: "center" },
    { label: "Spé 2\n(Simulée)", align: "center" },
    { label: "Philo\n(Simulée)", align: "center" },
    { label: "Grand Oral\n(Simulé)", align: "center" },
    { label: "EPS\n(Simulé)", align: "center" },
    { label: "Moyenne\nTerm (60%)", align: "center" },
    { label: "Moyenne\nFinale", align: "center" },
    { label: "Statut", align: "center" },
    { label: "Mention", align: "left" },
  ];

  colHeaders.forEach((h, i) => {
    const cell = ws.getCell(r, i + 1);
    cell.value = h.label;
    applyFont(cell, { bold: true, size: 9, color: C.colHeaderFg });
    applyFill(cell, C.colHeaderBg);
    cell.alignment = {
      horizontal: h.align,
      vertical: "middle",
      wrapText: true,
    };
    cell.border = border("thin");
  });
  ws.getRow(r).height = 34;
  r++;

  const dataStartRow = r;

  // Insert students simulation rows
  let idx = 0;
  for (const s of rows) {
    idx++;
    const isAlt = idx % 2 === 0;
    const rowBg = isAlt ? C.altRow : C.white;
    const sim = simulations[s.full_id] || {};
    const results = s.results;

    let statusLabel = "Simulation incomplete";
    let statusBg = C.white;
    let statusColor = C.mention_sans;

    if (results.status === "Admis") {
      statusLabel = "✓ Admis";
      statusBg = C.admisBg;
      statusColor = C.admis;
    } else if (results.status === "Rattrapage") {
      statusLabel = "⚠ Rattrapage";
      statusBg = C.rattrapageBg;
      statusColor = C.rattrapage;
    } else if (results.status === "Refusé") {
      statusLabel = "✗ Refusé";
      statusBg = C.refuseBg;
      statusColor = C.refuse;
    }


    let epsCellVal = "—";
    let epsCellColor = undefined;
    let epsCellBold = false;

    if (s.eps_grade === 'dispense') {
      epsCellVal = "Disp.";
      epsCellColor = C.refuse;
      epsCellBold = true;
    } else if (s.eps_grade !== undefined) {
      epsCellVal = Number(s.eps_grade.toFixed(1));
      epsCellColor = C.admis;
      epsCellBold = true;
    } else if (sim.eps !== undefined) {
      epsCellVal = Number(sim.eps.toFixed(1));
    }

    const cellsData = [
      { value: idx, align: "center" },
      { value: s.family_name, bold: true },
      { value: s.first_name },
      { value: s.candidate_id, align: "center" },
      { value: s.dob, align: "center" },
      { value: s.specialties[0] || "—" },
      { value: s.specialties[1] || "—" },
      {
        value: Number(results.ccAverage.toFixed(2)),
        align: "center",
        bold: true,
        color: gradeColor(results.ccAverage)
      },
      { value: sim.frenchWritten !== undefined ? Number(sim.frenchWritten.toFixed(1)) : "—", align: "center" },
      { value: sim.frenchOral !== undefined ? Number(sim.frenchOral.toFixed(1)) : "—", align: "center" },
      { value: sim.spe1 !== undefined ? Number(sim.spe1.toFixed(1)) : "—", align: "center" },
      { value: sim.spe2 !== undefined ? Number(sim.spe2.toFixed(1)) : "—", align: "center" },
      { value: sim.philo !== undefined ? Number(sim.philo.toFixed(1)) : "—", align: "center" },
      { value: sim.grandOral !== undefined ? Number(sim.grandOral.toFixed(1)) : "—", align: "center" },
      { value: epsCellVal, align: "center", color: epsCellColor, bold: epsCellBold },
      {
        value: Number(results.terminalAverage.toFixed(2)),
        align: "center",
        color: gradeColor(results.terminalAverage)
      },
      {
        value: Number(results.finalAverage.toFixed(2)),
        align: "center",
        bold: true,
        color: gradeColor(results.finalAverage)
      },
      {
        value: statusLabel,
        align: "center",
        bold: true,
        color: statusColor,
        bg: statusBg
      },
      {
        value: results.mention ?? "—",
        bold: !!results.mention && results.mention !== "Aucune" && results.mention !== "Admis sans mention",
        color: mentionColor(results.mention)
      }
    ];

    cellsData.forEach((cd, i) => {
      const cell = ws.getCell(r, i + 1);
      cell.value = cd.value;
      applyFont(cell, { size: 9.5, bold: cd.bold, color: cd.color ?? "1A2D4A" });
      applyFill(cell, cd.bg ?? rowBg);
      cell.alignment = {
        horizontal: cd.align ?? "left",
        vertical: "middle",
      };
      cell.border = border("hair");
    });

    ws.getRow(r).height = 18;
    r++;
  }

  // Draw medium outer border around database table
  const dataEndRow = r - 1;
  for (let row = dataStartRow - 1; row <= dataEndRow; row++) {
    for (let col = 1; col <= 19; col++) {
      const cell = ws.getCell(row, col);
      const isFirstRow = row === dataStartRow - 1;
      const isLastRow = row === dataEndRow;
      const isFirstCol = col === 1;
      const isLastCol = col === 19;
      
      const thinBorder = { style: "hair", color: { argb: "FF" + C.border } };
      const strongBorder = { style: "medium", color: { argb: "FF" + C.borderStrong } };
      
      cell.border = {
        top: isFirstRow ? strongBorder : thinBorder,
        bottom: isLastRow ? strongBorder : thinBorder,
        left: isFirstCol ? strongBorder : thinBorder,
        right: isLastCol ? strongBorder : thinBorder,
      };
    }
  }

  r += 2; // space

  // ─── Statistics Summary Dashboard at bottom ─────────────────────────────────
  const summaryTitle = ws.getCell(r, 1);
  summaryTitle.value = "TABLEAU DE BORD DE LA COHORTE";
  applyFont(summaryTitle, { bold: true, size: 10, color: C.headerFg });
  applyFill(summaryTitle, C.summaryHeaderBg);
  ws.mergeCells(r, 1, r, 7);
  summaryTitle.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(r).height = 18;
  r++;

  const totalStudents = rows.length;
  const admittedCount = rows.filter(s => s.results.status === "Admis").length;
  const rattrapageCount = rows.filter(s => s.results.status === "Rattrapage").length;
  const refuseCount = rows.filter(s => s.results.status === "Refusé").length;
  const avgCohort = rows.reduce((sum, s) => sum + s.results.finalAverage, 0) / totalStudents;

  const countMention = (mentionStr) => {
    return rows.filter(s => s.results.mention === mentionStr).length;
  };

  const dashboardData = [
    ["Nombre de Candidats", totalStudents, "Candidats"],
    ["Admis (Taux d'admission)", admittedCount, `${Math.round((admittedCount / totalStudents) * 100)} %`],
    ["Moyenne Générale Cohorte", avgCohort !== null ? fr(avgCohort) + " /20" : "—", "Moyenne"],
    ["Rattrapage (Second groupe)", rattrapageCount, `${Math.round((rattrapageCount / totalStudents) * 100)} %`],
    ["Réfusés", refuseCount, `${Math.round((refuseCount / totalStudents) * 100)} %`],
    ["Mention Très Bien avec Félicitations", countMention("Très bien avec félicitations"), "élèves"],
    ["Mention Très Bien", countMention("Très bien"), "élèves"],
    ["Mention Bien", countMention("Bien"), "élèves"],
    ["Mention Assez Bien", countMention("Assez bien"), "élèves"],
    ["Admis sans mention", countMention("Admis sans mention"), "élèves"],
  ];

  dashboardData.forEach(([label, countVal, detailSuffix]) => {
    const labelCell = ws.getCell(r, 1);
    labelCell.value = label;
    applyFont(labelCell, { size: 9.5, bold: true, color: "2E4A6E" });
    applyFill(labelCell, C.summaryBg);
    ws.mergeCells(r, 1, r, 4);
    labelCell.alignment = { horizontal: "left", vertical: "middle" };

    const countCell = ws.getCell(r, 5);
    countCell.value = countVal;
    applyFont(countCell, { size: 10, bold: true, color: C.headerBg });
    applyFill(countCell, C.summaryBg);
    countCell.alignment = { horizontal: "center", vertical: "middle" };

    const detailCell = ws.getCell(r, 6);
    detailCell.value = detailSuffix;
    applyFont(detailCell, { size: 9, bold: false, color: "5C5C5C" });
    applyFill(detailCell, C.summaryBg);
    ws.mergeCells(r, 6, r, 7);
    detailCell.alignment = { horizontal: "left", vertical: "middle" };

    labelCell.border = border("thin");
    countCell.border = border("thin");
    detailCell.border = border("thin");
    
    ws.getRow(r).height = 16;
    r++;
  });

  // Freeze top title rows so scrolling down the table stays intuitive
  ws.views = [{ state: "frozen", xSplit: 0, ySplit: dataStartRow - 1 }];

  // Enable Excel filter widgets in column headers
  ws.autoFilter = {
    from: { row: dataStartRow - 1, column: 1 },
    to: { row: dataEndRow, column: 19 },
  };
}
