(function () {
  const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  const state = {
    heightUnit: "cm",
    weightUnit: "kg",
  };

  const el = {
    form: document.getElementById("pcnsl-form"),
    status: document.getElementById("status"),
    fillSample: document.getElementById("fill-sample"),
    downloadOncopharm: document.getElementById("download-oncopharm"),
    name: document.getElementById("name"),
    pwi: document.getElementById("pwi"),
    caseNumber: document.getElementById("case-number"),
    contactNumber: document.getElementById("contact-number"),
    age: document.getElementById("age"),
    sex: document.getElementById("sex"),
    address: document.getElementById("address"),
    assessment: document.getElementById("assessment"),
    heightToggle: document.getElementById("height-toggle"),
    heightCmWrap: document.getElementById("height-cm-wrap"),
    heightFtInWrap: document.getElementById("height-ftin-wrap"),
    heightCm: document.getElementById("height-cm"),
    heightFt: document.getElementById("height-ft"),
    heightIn: document.getElementById("height-in"),
    weightToggle: document.getElementById("weight-toggle"),
    weightKgWrap: document.getElementById("weight-kg-wrap"),
    weightLbsWrap: document.getElementById("weight-lbs-wrap"),
    weightKg: document.getElementById("weight-kg"),
    weightLbs: document.getElementById("weight-lbs"),
    dateReferred: document.getElementById("date-referred"),
    mtxStart: document.getElementById("mtx-start"),
    includeRituximab: document.getElementById("include-rituximab"),
    rituximabOffset: document.getElementById("rituximab-offset"),
    inductionCycles: document.getElementById("induction-cycles"),
    inductionInterval: document.getElementById("induction-interval"),
    consolidationCycles: document.getElementById("consolidation-cycles"),
    consolidationInterval: document.getElementById("consolidation-interval"),
    rituximabMgm2: document.getElementById("rituximab-mgm2"),
    rituximabRound: document.getElementById("rituximab-round"),
    rituximabVialMg: document.getElementById("rituximab-vial-mg"),
    rituximabPrice: document.getElementById("rituximab-price"),
    rituximabVial100Mg: document.getElementById("rituximab-vial-100-mg"),
    rituximabPrice100: document.getElementById("rituximab-price-100"),
    mtxGm2: document.getElementById("mtx-gm2"),
    mtxRound: document.getElementById("mtx-round"),
    mtxVialG: document.getElementById("mtx-vial-g"),
    mtxPrice: document.getElementById("mtx-price"),
    metricBsa: document.getElementById("metric-bsa"),
    metricRituximab: document.getElementById("metric-rituximab"),
    metricMtx: document.getElementById("metric-mtx"),
    metricTotal: document.getElementById("metric-total"),
    scheduleTable: document.getElementById("schedule-table"),
    drugTable: document.getElementById("drug-table"),
  };

  function num(value, fallback = 0) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : fallback;
  }

  function asDate(value) {
    if (!value) return null;
    return new Date(`${value}T00:00:00`);
  }

  function addDays(date, days) {
    const next = new Date(date);
    next.setDate(next.getDate() + days);
    return next;
  }

  function fmtDate(date) {
    return new Intl.DateTimeFormat("en-US", {
      month: "long",
      day: "numeric",
      year: "numeric",
    }).format(date);
  }

  function fmtShort(date) {
    return new Intl.DateTimeFormat("en-US", {
      month: "short",
      day: "numeric",
      year: "numeric",
    }).format(date);
  }

  function money(value) {
    return new Intl.NumberFormat("en-PH", {
      style: "currency",
      currency: "PHP",
      minimumFractionDigits: 2,
    }).format(value);
  }

  function peso(value) {
    return `Php ${new Intl.NumberFormat("en-PH", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(value)}`;
  }

  function computeBsa(weightKg, heightCm) {
    return 0.007184 * Math.pow(heightCm, 0.725) * Math.pow(weightKg, 0.425);
  }

  function round1(value) {
    return Math.round(value * 10) / 10;
  }

  function displayNumber(value, decimals = 1) {
    const rounded = Number(value.toFixed(decimals));
    return Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(decimals);
  }

  function getHeightCm() {
    if (state.heightUnit === "cm") return num(el.heightCm.value);
    return ((num(el.heightFt.value) * 12) + num(el.heightIn.value)) * 2.54;
  }

  function getWeightKg() {
    if (state.weightUnit === "kg") return num(el.weightKg.value);
    return num(el.weightLbs.value) * 0.45359237;
  }

  function syncHeightFieldsFromCm(cm) {
    if (!cm) return;
    el.heightCm.value = displayNumber(round1(cm));
    const totalInches = cm / 2.54;
    const feet = Math.floor(totalInches / 12);
    const inches = round1(totalInches - feet * 12);
    el.heightFt.value = String(feet);
    el.heightIn.value = displayNumber(inches);
  }

  function syncWeightFieldsFromKg(kg) {
    if (!kg) return;
    el.weightKg.value = displayNumber(round1(kg));
    el.weightLbs.value = displayNumber(round1(kg / 0.45359237));
  }

  function roundTo(value, increment) {
    const step = Math.max(num(increment, 1), 0.01);
    return Math.max(Math.round(value / step) * step, step);
  }

  function plural(value, unit) {
    return `${value} ${unit}${Number(value) === 1 ? "" : "s"}`;
  }

  function doseG(value) {
    return `${Number(value).toFixed(Number(value) % 1 ? 1 : 0)} g`;
  }

  function vialLine(count, strength, unit) {
    if (!count) return "";
    return `${count} x ${strength}${unit} vial`;
  }

  function joinVialLines(parts) {
    return parts.filter(Boolean).join(" and ");
  }

  function rituximabVialDescription(model, multiplier = 1) {
    if (!model.includeRituximab) return "Not included";
    return joinVialLines([
      vialLine(model.rituximabVials500PerDose * multiplier, model.rituximabVialMg, "mg"),
      vialLine(model.rituximabVials100PerDose * multiplier, model.rituximabVial100Mg, "mg"),
    ]);
  }

  function computeRituximabVials(doseMg, vial500Mg, price500, vial100Mg, price100) {
    const largeMg = Math.max(vial500Mg, 1);
    const smallMg = Math.max(vial100Mg, 1);
    const maxLarge = Math.ceil(doseMg / largeMg) + 2;
    const maxSmall = Math.ceil(doseMg / smallMg) + 2;
    let best = null;
    for (let large = 0; large <= maxLarge; large += 1) {
      for (let small = 0; small <= maxSmall; small += 1) {
        const covered = (large * largeMg) + (small * smallMg);
        if (covered < doseMg) continue;
        const excess = covered - doseMg;
        const cost = (large * price500) + (small * price100);
        const totalVials = large + small;
        if (
          !best ||
          excess < best.excess ||
          (excess === best.excess && cost < best.cost) ||
          (excess === best.excess && cost === best.cost && totalVials < best.totalVials)
        ) {
          best = { large, small, covered, excess, cost, totalVials };
        }
      }
    }
    return best || { large: 0, small: 0, covered: 0, excess: 0, cost: 0, totalVials: 0 };
  }

  function cleanLines(text) {
    return String(text || "")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
  }

  function makeSchedule(firstMtx, cycles, interval, rituximabOffset) {
    return Array.from({ length: cycles }, (_, index) => {
      const mtxDate = addDays(firstMtx, index * interval);
      return {
        cycle: index + 1,
        methotrexate: mtxDate,
        rituximab: addDays(mtxDate, rituximabOffset),
      };
    });
  }

  function buildModel() {
    const weight = getWeightKg();
    const height = getHeightCm();
    const mtxStart = asDate(el.mtxStart.value);
    const dateReferred = asDate(el.dateReferred.value);
    if (!weight || !height || !mtxStart || !dateReferred) {
      throw new Error("Please complete height, weight, date referred, and first methotrexate date.");
    }

    const bsa = Math.round(computeBsa(weight, height) * 100) / 100;
    const includeRituximab = el.includeRituximab.checked;
    const inductionCycles = Math.min(Math.max(Math.round(num(el.inductionCycles.value, 6)), 1), 8);
    const consolidationCycles = Math.min(Math.max(Math.round(num(el.consolidationCycles.value, 6)), 1), 6);
    const inductionInterval = Math.max(Math.round(num(el.inductionInterval.value, 14)), 1);
    const consolidationInterval = Math.max(Math.round(num(el.consolidationInterval.value, 28)), 1);
    const rituximabOffset = Math.round(num(el.rituximabOffset.value, 1));
    const induction = makeSchedule(mtxStart, inductionCycles, inductionInterval, rituximabOffset);
    const lastInductionMtx = induction[induction.length - 1].methotrexate;
    const firstConsolidationMtx = addDays(lastInductionMtx, consolidationInterval);
    const consolidation = makeSchedule(firstConsolidationMtx, consolidationCycles, consolidationInterval, rituximabOffset);

    const rituximabDose = roundTo(bsa * num(el.rituximabMgm2.value, 375), num(el.rituximabRound.value, 50));
    const mtxDoseG = roundTo(bsa * num(el.mtxGm2.value, 8), num(el.mtxRound.value, 0.5));
    const rituximabVialMg = Math.max(num(el.rituximabVialMg.value, 500), 1);
    const rituximabVial100Mg = Math.max(num(el.rituximabVial100Mg.value, 100), 1);
    const rituximabPrice = num(el.rituximabPrice.value);
    const rituximabPrice100 = num(el.rituximabPrice100.value, 8343.75);
    const rituximabVialPlan = includeRituximab ? computeRituximabVials(
      rituximabDose,
      rituximabVialMg,
      rituximabPrice,
      rituximabVial100Mg,
      rituximabPrice100,
    ) : { large: 0, small: 0, covered: 0, excess: 0, cost: 0, totalVials: 0 };
    const rituximabVialsPerDose = rituximabVialPlan.totalVials;
    const mtxVialsPerDose = Math.ceil(mtxDoseG / Math.max(num(el.mtxVialG.value, 1), 0.1));
    const rituximabCycles = inductionCycles + consolidationCycles;
    const mtxCycles = inductionCycles + consolidationCycles;
    const rituximabTotalVials = rituximabVialsPerDose * rituximabCycles;
    const mtxTotalVials = mtxVialsPerDose * mtxCycles;
    const rituximabTotal = rituximabVialPlan.cost * rituximabCycles;
    const mtxTotal = mtxTotalVials * num(el.mtxPrice.value);

    return {
      name: String(el.name.value || "").trim().toUpperCase(),
      pwi: String(el.pwi.value || "").trim(),
      caseNumber: String(el.caseNumber.value || "").trim(),
      contactNumber: String(el.contactNumber.value || "").trim(),
      age: Math.round(num(el.age.value)),
      sex: String(el.sex.value || "").trim().toUpperCase(),
      address: String(el.address.value || "").trim().toUpperCase(),
      assessment: cleanLines(el.assessment.value),
      weight: round1(weight),
      height: round1(height),
      bsa,
      dateReferred,
      induction,
      consolidation,
      inductionCycles,
      consolidationCycles,
      includeRituximab,
      rituximabDose,
      rituximabVialsPerDose,
      rituximabVials500PerDose: rituximabVialPlan.large,
      rituximabVials100PerDose: rituximabVialPlan.small,
      rituximabTotalVials,
      rituximabTotal500Vials: rituximabVialPlan.large * rituximabCycles,
      rituximabTotal100Vials: rituximabVialPlan.small * rituximabCycles,
      rituximabDosePrepared: rituximabVialPlan.covered,
      rituximabDoseExcess: rituximabVialPlan.excess,
      rituximabCostPerDose: rituximabVialPlan.cost,
      rituximabTotal,
      mtxDoseG,
      mtxVialsPerDose,
      mtxTotalVials,
      mtxTotal,
      total: rituximabTotal + mtxTotal,
      rituximabPrice,
      rituximabPrice100,
      mtxPrice: num(el.mtxPrice.value),
      rituximabVialMg,
      rituximabVial100Mg,
      mtxVialG: num(el.mtxVialG.value, 1),
    };
  }

  function setStatus(message, error = false) {
    el.status.textContent = message;
    el.status.style.color = error ? "#9b2f1c" : "#5d6a70";
  }

  function scheduleRows(model) {
    return [
      ...model.induction.map((item) => ({ phase: `Induction ${item.cycle}`, ...item })),
      ...model.consolidation.map((item) => ({ phase: `Consolidation ${item.cycle}`, ...item })),
    ];
  }

  function wrapText(text, maxChars) {
    const words = String(text || "").split(/\s+/).filter(Boolean);
    const lines = [];
    let current = "";
    words.forEach((word) => {
      const next = current ? `${current} ${word}` : word;
      if (next.length > maxChars && current) {
        lines.push(current);
        current = word;
      } else {
        current = next;
      }
    });
    if (current) lines.push(current);
    return lines;
  }

  function updateSummary() {
    try {
      const model = buildModel();
      el.metricBsa.textContent = `${model.bsa.toFixed(2)} m2`;
      el.metricRituximab.textContent = model.includeRituximab ? `${model.rituximabDose} mg` : "Removed";
      el.metricMtx.textContent = `${model.mtxDoseG.toFixed(model.mtxDoseG % 1 ? 1 : 0)} g`;
      el.metricTotal.textContent = money(model.total);
      el.scheduleTable.innerHTML = scheduleRows(model)
        .map((row) => `<tr><td>${row.phase}</td><td>${model.includeRituximab ? fmtShort(row.rituximab) : "Removed"}</td><td>${fmtShort(row.methotrexate)}</td></tr>`)
        .join("");
      el.drugTable.innerHTML = [
        model.includeRituximab ? `<tr><td>Rituximab</td><td>${model.rituximabDose} mg</td><td>${rituximabVialDescription(model, model.inductionCycles + model.consolidationCycles)}</td><td>${money(model.rituximabTotal)}</td></tr>` : "",
        `<tr><td>Methotrexate</td><td>${model.mtxDoseG} g</td><td>${plural(model.mtxTotalVials, "vial")}</td><td>${money(model.mtxTotal)}</td></tr>`,
      ].join("");
      setStatus("");
    } catch (error) {
      el.metricBsa.textContent = "-";
      el.metricRituximab.textContent = "-";
      el.metricMtx.textContent = "-";
      el.metricTotal.textContent = "-";
      el.scheduleTable.innerHTML = "";
      el.drugTable.innerHTML = "";
      setStatus(error.message, true);
    }
  }

  function base64ToUint8Array(base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }

  async function loadDocx(base64, label) {
    if (!base64) throw new Error(`${label} template data is missing.`);
    return JSZip.loadAsync(base64ToUint8Array(base64));
  }

  function getTextNodes(doc) {
    return Array.from(doc.getElementsByTagName("w:t"));
  }

  function setText(nodes, index, value) {
    if (nodes[index]) nodes[index].textContent = String(value);
  }

  function setMany(nodes, updates) {
    updates.forEach(([index, value]) => setText(nodes, index, value));
  }

  function replaceParagraphWithLines(nodes, startIndex, maxExistingLines, lines) {
    for (let i = 0; i < maxExistingLines; i += 1) {
      setText(nodes, startIndex + i, lines[i] || "");
    }
  }

  function directRows(table) {
    return Array.from(table.childNodes).filter((node) => node.nodeName === "w:tr");
  }

  function rowTexts(row) {
    return Array.from(row.getElementsByTagName("w:t"));
  }

  function rowIncludes(row, text) {
    return rowTexts(row).some((node) => node.textContent === text);
  }

  function fillRituximabRow(row, cycle, date) {
    const nodes = rowTexts(row);
    if (nodes.length >= 4) {
      nodes[0].textContent = String(cycle);
      nodes[1].textContent = "Rituximab";
      nodes[2].textContent = "1";
      nodes[3].textContent = fmtDate(date);
    }
  }

  function fillCycleDrugRow(row, cycle, drug, date) {
    const nodes = rowTexts(row);
    if (nodes.length >= 4) {
      nodes[0].textContent = String(cycle);
      nodes[1].textContent = drug;
      nodes[2].textContent = "1";
      nodes[3].textContent = fmtDate(date);
    }
  }

  function fillMethotrexateRow(row, date) {
    const nodes = rowTexts(row);
    if (nodes.length >= 3) {
      const offset = nodes.length === 3 ? 0 : 1;
      if (nodes.length === 4) nodes[0].textContent = "";
      nodes[offset].textContent = "Methotrexate";
      nodes[offset + 1].textContent = "1";
      nodes[offset + 2].textContent = fmtDate(date);
    }
  }

  function fillScheduleTable(doc, tableIndex, schedule, includeRituximab = true) {
    const table = doc.getElementsByTagName("w:tbl")[tableIndex];
    if (!table) return;
    const rows = directRows(table);
    const headerIndex = rows.findIndex((row) => rowIncludes(row, "Cycle"));
    const rituxTemplate = rows.find((row) => rowIncludes(row, "Rituximab"));
    const mtxTemplate = rows.find((row) => rowIncludes(row, "Methotrexate"));
    if (headerIndex < 0 || !rituxTemplate || !mtxTemplate) return;

    rows.slice(headerIndex + 1).forEach((row) => table.removeChild(row));
    schedule.forEach((item, index) => {
      if (includeRituximab) {
        const rituxRow = rituxTemplate.cloneNode(true);
        const mtxRow = mtxTemplate.cloneNode(true);
        fillRituximabRow(rituxRow, index + 1, item.rituximab);
        fillMethotrexateRow(mtxRow, item.methotrexate);
        table.appendChild(rituxRow);
        table.appendChild(mtxRow);
      } else {
        const mtxRow = rituxTemplate.cloneNode(true);
        fillCycleDrugRow(mtxRow, index + 1, "Methotrexate", item.methotrexate);
        table.appendChild(mtxRow);
      }
    });
  }

  function serializeDocx(zip, doc) {
    zip.file("word/document.xml", new XMLSerializer().serializeToString(doc));
    return zip.generateAsync({ type: "uint8array", mimeType: DOCX_MIME });
  }

  function blankIndexes(nodes, indexes) {
    indexes.forEach((index) => setText(nodes, index, ""));
  }

  async function generateProtocol(model) {
    const zip = await loadDocx(PCNSL_PROTOCOL_DOCX_BASE64, "Protocol");
    const xml = await zip.file("word/document.xml").async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const nodes = getTextNodes(doc);
    const r = model.rituximabDose;
    const m = model.mtxDoseG;

    setMany(nodes, [
      [1, ` ${model.name}`],
      [3, fmtDate(model.dateReferred)],
      [5, ` ${model.address || "-"}`],
      [7, ` ${model.age} y/${model.sex}`],
      [9, model.pwi],
    ]);
    fillScheduleTable(doc, 1, model.induction, model.includeRituximab);
    fillScheduleTable(doc, 2, model.consolidation, model.includeRituximab);
    if (model.includeRituximab) {
      setMany(nodes, [
        [108, `${r} mg`],
        [109, String(model.inductionCycles)],
        [110, `${model.rituximabVialMg} mg/vial; ${model.rituximabVial100Mg} mg/vial`],
        [111, rituximabVialDescription(model, model.inductionCycles)],
        [112, `${peso(model.rituximabPrice)}/${model.rituximabVialMg}mg vial; ${peso(model.rituximabPrice100)}/${model.rituximabVial100Mg}mg vial`],
        [113, peso(model.rituximabCostPerDose * model.inductionCycles)],
        [127, `${r} mg`],
        [128, String(model.consolidationCycles)],
        [129, `${model.rituximabVialMg} mg/vial; ${model.rituximabVial100Mg} mg/vial`],
        [130, rituximabVialDescription(model, model.consolidationCycles)],
        [131, `${peso(model.rituximabPrice)}/${model.rituximabVialMg}mg vial; ${peso(model.rituximabPrice100)}/${model.rituximabVial100Mg}mg vial`],
        [132, peso(model.rituximabCostPerDose * model.consolidationCycles)],
      ]);
    } else {
      blankIndexes(nodes, [105, 106, 107, 108, 109, 110, 111, 112, 113, 124, 125, 126, 127, 128, 129, 130, 131, 132]);
    }
    setMany(nodes, [
      [117, `${m} g`],
      [118, String(model.inductionCycles)],
      [120, String(model.mtxVialsPerDose * model.inductionCycles)],
      [121, `${peso(model.mtxPrice)}/vial`],
      [122, peso(model.mtxVialsPerDose * model.inductionCycles * model.mtxPrice)],
      [136, `${m} g`],
      [137, String(model.consolidationCycles)],
      [139, String(model.mtxVialsPerDose * model.consolidationCycles)],
      [140, `${peso(model.mtxPrice)}/vial`],
      [141, peso(model.mtxVialsPerDose * model.consolidationCycles * model.mtxPrice)],
      [143, peso(model.rituximabCostPerDose + (model.mtxVialsPerDose * model.mtxPrice))],
      [145, peso(model.total)],
    ]);

    return serializeDocx(zip, doc);
  }

  async function generateEndorsement(model) {
    const zip = await loadDocx(PCNSL_ENDORSEMENT_DOCX_BASE64, "Endorsement");
    const xml = await zip.file("word/document.xml").async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const nodes = getTextNodes(doc);

    setMany(nodes, [
      [6, model.name],
      [8, model.caseNumber || "-"],
      [12, `${model.age} / ${model.sex}`],
      [16, fmtDate(model.dateReferred)],
      [20, model.contactNumber || "-"],
      [22, `Height ${Math.round(model.height)}cm; Weight ${model.weight} kg; BSA ${model.bsa.toFixed(2)} m`],
      [23, "2"],
      [34, model.includeRituximab ? `For induction chemotherapy with Rituximab and High-Dose Methotrexate every 2 weeks for ${model.inductionCycles} cycles then every 4 weeks for ${model.consolidationCycles} consolidation cycles` : `For induction chemotherapy with High-Dose Methotrexate every 2 weeks for ${model.inductionCycles} cycles then every 4 weeks for ${model.consolidationCycles} consolidation cycles`],
      [137, model.includeRituximab ? `Administer Rituximab ${model.rituximabDose} mg in enough PNSS to make 500mL via infusion pump x 50cc/hour for the first 30 minutes, then increase by 50cc/hr every 30 minutes to maximum 150cc/hr` : ""],
      [153, `Hook methotrexate ${model.mtxDoseG}g (8g/m`],
    ]);
    if (!model.includeRituximab) {
      blankIndexes(nodes, [114, 115, 116, 117, 118, 119, 120, 128, 129, 130, 131, 132, 133, 134, 135, 136, 138, 139, 140]);
    }
    replaceParagraphWithLines(nodes, 25, 7, model.assessment.length ? model.assessment : [model.pwi]);
    fillScheduleTable(doc, 3, model.induction, model.includeRituximab);
    fillScheduleTable(doc, 4, model.consolidation, model.includeRituximab);

    return serializeDocx(zip, doc);
  }

  async function generateMethotrexateLetter(model) {
    const zip = await loadDocx(METHOTREXATE_INPATIENT_DOCX_BASE64, "Methotrexate inpatient");
    const xml = await zip.file("word/document.xml").async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const nodes = getTextNodes(doc);
    const total = model.mtxVialsPerDose * model.mtxPrice;
    setMany(nodes, [
      [1, fmtDate(model.dateReferred)],
      [15, model.name],
      [18, model.caseNumber || "-"],
      [23, model.pwi || "Primary CNS Lymphoma"],
      [40, `8g/BSA = 8g x ${model.bsa.toFixed(2)} m`],
      [42, ` = ${doseG(model.mtxDoseG)}`],
      [43, plural(model.mtxVialsPerDose, "vial")],
      [46, ` ${new Intl.NumberFormat("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(total)}`],
      [48, ` Methotrexate ${doseG(model.mtxDoseG)} (8g/BSA) in enough `],
    ]);
    return serializeDocx(zip, doc);
  }

  async function generateRituximabLetter(model) {
    const zip = await loadDocx(RITUXIMAB_INPATIENT_DOCX_BASE64, "Rituximab inpatient");
    const xml = await zip.file("word/document.xml").async("string");
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const nodes = getTextNodes(doc);
    const total = model.rituximabCostPerDose;
    setMany(nodes, [
      [1, fmtDate(model.dateReferred)],
      [15, model.name],
      [19, model.caseNumber || "-"],
      [24, model.pwi || "Primary CNS Lymphoma"],
      [37, `${model.rituximabDose}mg`],
      [39, ` ${new Intl.NumberFormat("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(model.rituximabPrice)} per ${model.rituximabVialMg}mg vial`],
      [41, ` ${new Intl.NumberFormat("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(model.rituximabPrice100)} per ${model.rituximabVial100Mg}mg vial`],
      [42, `375mg/BSA = 375mg x ${model.bsa.toFixed(2)} m`],
      [44, ` = ${model.rituximabDose} mg`],
      [45, rituximabVialDescription(model)],
      [48, ` ${new Intl.NumberFormat("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(total)}`],
      [50, ` Rituximab ${model.rituximabDose}mg (375mg/BSA) in enough `],
    ]);
    return serializeDocx(zip, doc);
  }

  function replacePdfAnnotationText(pdfDoc, replacements, indexedReplacements = {}) {
    const { PDFName, PDFString } = window.PDFLib;
    const page = pdfDoc.getPages()[0];
    const annots = page.node.Annots();
    if (!annots) return false;
    let changed = false;
    for (let i = 0; i < annots.size(); i += 1) {
      const annot = pdfDoc.context.lookup(annots.get(i));
      const contents = annot.get(PDFName.of("Contents"));
      const current = contents && typeof contents.decodeText === "function" ? contents.decodeText() : "";
      const nextText = Object.prototype.hasOwnProperty.call(indexedReplacements, i)
        ? indexedReplacements[i]
        : replacements[current];
      if (typeof nextText === "string") {
        annot.set(PDFName.of("Contents"), PDFString.of(nextText));
        annot.delete(PDFName.of("AP"));
        changed = true;
      }
    }
    return changed;
  }

  async function generateOncopharmPdf(model) {
    if (!window.PDFLib) throw new Error("PDF generator library is missing.");
    const { PDFDocument, StandardFonts, rgb } = window.PDFLib;
    const pdfDoc = await PDFDocument.load(base64ToUint8Array(CHEMO_BLANK_PDF_BASE64));
    const textboxEdited = replacePdfAnnotationText(pdfDoc, {
      "[Name]": model.name,
      "[Diagnosis]": model.pwi || "Primary CNS Lymphoma",
      "[age]": String(model.age),
      "[sex]": model.sex,
      "Methotrexate": "Methotrexate",
      "Rituximab": model.includeRituximab ? "Rituximab" : "",
      "1 day (day 1)": `Day 1 - ${fmtShort(model.induction[0].methotrexate)}`,
      "1 day (day 2)": model.includeRituximab ? `Day 2 - ${fmtShort(model.induction[0].rituximab)}` : "",
      "in enough PNSS to make 500mL": "in enough PNSS to make 500mL",
    }, {
      5: `${model.bsa.toFixed(2)} m2`,
      10: `${model.mtxDoseG * 1000} mg`,
      11: model.includeRituximab ? `${model.rituximabDose} mg` : "",
      13: model.includeRituximab ? "in enough PNSS to make 500mL" : "",
      19: `${model.bsa.toFixed(2)} m2`,
      24: `${model.mtxDoseG * 1000} mg`,
      25: model.includeRituximab ? `${model.rituximabDose} mg` : "",
      27: model.includeRituximab ? "in enough PNSS to make 500mL" : "",
    });
    if (textboxEdited) {
      return pdfDoc.save();
    }

    const page = pdfDoc.getPages()[0];
    const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
    const bold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
    const { width, height } = page.getSize();
    const ink = rgb(0.04, 0.08, 0.1);
    const accent = rgb(0, 0.33, 0.36);

    function draw(text, x, y, size = 9, useBold = false, color = ink) {
      page.drawText(String(text || ""), { x, y, size, font: useBold ? bold : font, color });
    }

    function line(y, fromX = 42, toX = width - 42) {
      page.drawLine({ start: { x: fromX, y }, end: { x: toX, y }, thickness: 0.6, color: rgb(0.75, 0.8, 0.82) });
    }

    page.drawRectangle({
      x: 32,
      y: height - 132,
      width: width - 64,
      height: 92,
      color: rgb(1, 1, 1),
      borderColor: rgb(0, 0.33, 0.36),
      borderWidth: 1,
      opacity: 0.92,
    });
    draw("ONCOPHARM PROTOCOL", 44, height - 62, 14, true, accent);
    draw(`Date: ${fmtDate(model.dateReferred)}`, 420, height - 62, 9, false);
    draw(`Patient: ${model.name}`, 44, height - 84, 10, true);
    draw(`Age/Sex: ${model.age}/${model.sex}`, 330, height - 84, 9);
    draw(`Case No.: ${model.caseNumber || "-"}`, 430, height - 84, 9);
    draw(`Diagnosis: ${model.pwi || "Primary CNS Lymphoma"}`, 44, height - 104, 9);
    draw(`Height: ${displayNumber(model.height)} cm   Weight: ${displayNumber(model.weight)} kg   BSA: ${model.bsa.toFixed(2)} m2`, 44, height - 122, 9);

    let y = height - 158;
    draw("Computed Chemotherapy", 44, y, 11, true, accent);
    y -= 14;
    line(y + 4);
    draw("Drug", 44, y - 8, 8, true);
    draw("Dose", 188, y - 8, 8, true);
    draw("Preparation", 292, y - 8, 8, true);
    draw("Total Items", 400, y - 8, 8, true);
    draw("Estimated Cost", 482, y - 8, 8, true);
    y -= 26;
    if (model.includeRituximab) {
      draw("Rituximab", 44, y, 9, true);
      draw(`${model.rituximabDose} mg`, 188, y, 9);
      draw(`${model.rituximabVialMg} mg + ${model.rituximabVial100Mg} mg`, 292, y, 9);
      draw(rituximabVialDescription(model, model.inductionCycles + model.consolidationCycles), 400, y, 8);
      draw(peso(model.rituximabTotal), 482, y, 9);
      y -= 18;
    }
    draw("Methotrexate", 44, y, 9, true);
    draw(`${model.mtxDoseG} g`, 188, y, 9);
    draw(`${model.mtxVialG} g/vial`, 292, y, 9);
    draw(`${model.mtxTotalVials} vials`, 400, y, 9);
    draw(peso(model.mtxTotal), 482, y, 9);
    y -= 18;
    draw(`Total estimate: ${peso(model.total)}`, 400, y, 10, true);

    y -= 34;
    draw("Schedule", 44, y, 11, true, accent);
    y -= 14;
    line(y + 4);
    draw("Phase", 44, y - 8, 8, true);
    draw("Rituximab", 180, y - 8, 8, true);
    draw("Methotrexate", 318, y - 8, 8, true);
    y -= 24;
    scheduleRows(model).forEach((row) => {
      draw(row.phase, 44, y, 8);
      draw(model.includeRituximab ? fmtShort(row.rituximab) : "Removed", 180, y, 8);
      draw(fmtShort(row.methotrexate), 318, y, 8);
      y -= 14;
    });

    y -= 10;
    draw("Protocol Notes", 44, y, 11, true, accent);
    y -= 16;
    [
      model.includeRituximab ? `Rituximab ${model.rituximabDose} mg in enough PNSS to make 500 mL; infusion pump per institutional protocol.` : "",
      `Methotrexate ${model.mtxDoseG} g in enough PNSS to make at least 500 mL; run at 125 cc/hr.`,
      `Induction: ${model.inductionCycles} cycles every 2 weeks. Consolidation: ${model.consolidationCycles} cycles every 4 weeks.`,
    ].filter(Boolean).forEach((text) => {
      wrapText(text, 92).forEach((lineText) => {
        draw(lineText, 54, y, 8);
        y -= 11;
      });
      y -= 2;
    });

    y -= 4;
    draw("Assessment", 44, y, 11, true, accent);
    y -= 16;
    const assessment = model.assessment.length ? model.assessment : [model.pwi || "Primary CNS Lymphoma"];
    assessment.slice(0, 6).forEach((lineText) => {
      wrapText(lineText, 96).forEach((wrapped) => {
        draw(wrapped, 54, y, 8);
        y -= 10;
      });
    });

    return pdfDoc.save();
  }

  function download(bytes, fileName, type) {
    const blob = new Blob([bytes], { type });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1200);
  }

  async function handleSubmit(event) {
    event.preventDefault();
    try {
      setStatus("Generating Word documents...");
      const model = buildModel();
      const protocol = await generateProtocol(model);
      const endorsement = await generateEndorsement(model);
      const methotrexateLetter = await generateMethotrexateLetter(model);
      const rituximabLetter = model.includeRituximab ? await generateRituximabLetter(model) : null;
      const bundle = new JSZip();
      const safe = (model.name || "patient").replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "") || "patient";
      bundle.file(`${safe}_PCNSL_protocol.docx`, protocol);
      bundle.file(`${safe}_PCNSL_endorsement.docx`, endorsement);
      bundle.file(`${safe}_methotrexate_inpatient_JL.docx`, methotrexateLetter);
      if (rituximabLetter) bundle.file(`${safe}_rituximab_inpatient_JL.docx`, rituximabLetter);
      const zipBytes = await bundle.generateAsync({ type: "uint8array" });
      download(zipBytes, `${safe}_PCNSL_documents.zip`, "application/zip");
      setStatus("Word documents generated successfully.");
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  async function handleOncopharmDownload() {
    try {
      setStatus("Generating oncopharm protocol PDF...");
      const model = buildModel();
      const bytes = await generateOncopharmPdf(model);
      const safe = (model.name || "patient").replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "") || "patient";
      download(bytes, `${safe}_oncopharm_protocol.pdf`, "application/pdf");
      setStatus("Oncopharm protocol PDF generated successfully.");
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  function loadSample() {
    el.name.value = "Juan Dela Cruz";
    el.pwi.value = "Primary CNS Lymphoma";
    el.caseNumber.value = "4024566";
    el.contactNumber.value = "09464293595";
    el.age.value = "47";
    el.sex.value = "F";
    el.address.value = "Metro Manila";
    el.assessment.value = "Primary CNS Lymphoma";
    syncHeightFieldsFromCm(145);
    syncWeightFieldsFromKg(55);
    el.dateReferred.value = "2026-04-17";
    el.mtxStart.value = "2026-04-18";
    el.rituximabOffset.value = "1";
    el.inductionCycles.value = "6";
    el.consolidationCycles.value = "6";
    updateSummary();
  }

  function setHeightUnit(unit) {
    const cm = getHeightCm();
    syncHeightFieldsFromCm(cm);
    state.heightUnit = unit;
    el.heightCmWrap.classList.toggle("hidden", unit !== "cm");
    el.heightFtInWrap.classList.toggle("hidden", unit !== "ftin");
    [...el.heightToggle.querySelectorAll("button")].forEach((button) => {
      button.classList.toggle("active", button.dataset.unit === unit);
    });
    updateSummary();
  }

  function setWeightUnit(unit) {
    const kg = getWeightKg();
    syncWeightFieldsFromKg(kg);
    state.weightUnit = unit;
    el.weightKgWrap.classList.toggle("hidden", unit !== "kg");
    el.weightLbsWrap.classList.toggle("hidden", unit !== "lbs");
    [...el.weightToggle.querySelectorAll("button")].forEach((button) => {
      button.classList.toggle("active", button.dataset.unit === unit);
    });
    updateSummary();
  }

  el.form.addEventListener("submit", handleSubmit);
  el.downloadOncopharm.addEventListener("click", handleOncopharmDownload);
  el.fillSample.addEventListener("click", loadSample);
  el.heightToggle.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-unit]");
    if (button) setHeightUnit(button.dataset.unit);
  });
  el.weightToggle.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-unit]");
    if (button) setWeightUnit(button.dataset.unit);
  });
  document.querySelectorAll("input, select, textarea").forEach((input) => {
    input.addEventListener("input", updateSummary);
    input.addEventListener("change", updateSummary);
  });

  loadSample();
})();
