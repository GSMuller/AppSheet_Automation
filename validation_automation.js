// No Nosso Caso , temos uma plataforma de cadastramento administrativo que criamos para nossa controladoria.

// =================================================================
// === CONSTANTES DE CONFIGURAÇÃO ==================================
// =================================================================

const SHEET_NAMES = {
  CADASTRO: "CADASTRO",
  DOCUMENTACAO: "DOCUMENTAÇÃO",
  CACHE_CADASTRO: "CACHE_MANUAL",
  CACHE_DOC: "CACHE_DOCUMENTAÇÃO",
  BASE_APOLLO: "BASE_APOLLO_DIARIA"
};

const COLORS = {
  GREEN: "#E6F8E6", YELLOW: "#FFFDE7", RED: "#FDEAEA",
  GREY: "#F5F5F5", BLUE: "#E6F7FF", WHITE: "#FFFFFF"
};

const STATUS = {
  VALID: "Documentos Validados!",
  PENDING: "Pendente Documentação!",
  MISSING_PREFIX: "Falta:",
  CHECK_BONUS: "Falta: Verificar Bônus",
  NA: "Não se aplica",
  NO_BONUS: "Sem Bônus/Venda Direta"
};

const CELLMARK = {
  VALID: "Validado! ✅",
  PENDING: "Pendente! ⚠️",
  PENDING_ALL_RED: "Pendente! ❌",
  NA: "Não se Aplica"
};

const DOC_COL_MAP = {
  proposta: 13, sinal: 14, comprovante_endereco: 15, nota_fiscal: 16, procuracao_atpv: 17,
  cnh: 18, contrato_comp_venda: 19, crv: 20, comp_vinculo: 21, guia: 22, comp_guia: 23, reembolso: 24
};
const COL_TO_DOC_KEY = Object.fromEntries(Object.entries(DOC_COL_MAP).map(([k, v]) => [v, k]));
const ALL_DOC_COLUMNS = Object.values(DOC_COL_MAP);

const DOC_LABEL_BY_KEY = {
  proposta: "Proposta", sinal: "Sinal", comprovante_endereco: "Comprovante Endereço",
  nota_fiscal: "Nota Fiscal", procuracao_atpv: "Procuração / ATPV", cnh: "CNH",
  contrato_comp_venda: "Contrato Comp. Venda", crv: "CRV", comp_vinculo: "Comp. Vínculo",
  guia: "Guia", comp_guia: "Comp. Guia", reembolso: "Reembolso"
};
const NORMAL_LABEL_TO_DOC_KEY = Object.fromEntries(Object.entries(DOC_LABEL_BY_KEY).map(([k, v]) => [normalizeText(v), k]));

const BONUS_DOC_MAP = {
  varejo: ["proposta", "sinal", "comprovante_endereco"],
  equalizacao: ["proposta", "sinal", "comprovante_endereco"],
  taxa_0: ["proposta", "sinal", "comprovante_endereco"],
  seguro: ["proposta", "sinal", "comprovante_endereco"],
  trade_in: ["nota_fiscal", "procuracao_atpv", "cnh", "contrato_comp_venda", "crv", "comp_vinculo"],
  ipva: ["nota_fiscal", "proposta", "sinal", "guia", "comp_guia"],
  wallbox: ["proposta", "sinal", "nota_fiscal"],
  portatil: ["proposta", "sinal", "nota_fiscal"]
};
const BONUS_SYNONYMS = {
  "taxa 0": "taxa_0", "trade in": "trade_in", "trade-in": "trade_in",
  "sem bonus": "sem_bonus", "sem bônus": "sem_bonus", "venda direta": "venda_direta"
};

// =================================================================
// === FUNÇÃO DE ABERTURA - MENU ===================================
// =================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⚙️ Operações")
    .addSubMenu(ui.createMenu("➡️ Cadastro")
      .addItem("🔄 Atualizar Apollo", "atualizarDadosCadastro")
      .addItem("💾 Guardar Informações", "guardarInformacoesCadastro")
      .addItem("📌 Reorganizar Dados", "reorganizarDadosCadastro"))
    .addSubMenu(ui.createMenu("➡️ Documentação")
      .addItem("🔄 Atualizar Apollo", "atualizarDadosDocumentacao")
      .addItem("💾 Guardar Informações", "guardarInformacoesDocumentacao")
      .addItem("📌 Reorganizar Dados", "reorganizarDadosDocumentacao"))
    .addSeparator()
    .addItem("ℹ️ Créditos", "mostrarInfo")
    .addToUi();

  configurarColunaDocumentacao_();
}

// =================================================================
// === CONFIGURAÇÕES DE DOCUMENTAÇÃO ===============================
// =================================================================

function configurarColunaDocumentacao_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DOCUMENTACAO);
  if (!sheet) return;

  const header = sheet.getRange("1:1").getValues()[0];
  if (header[11] !== "Documentação") {
    sheet.insertColumnAfter(11);
    sheet.getRange("L1").setValue("Documentação");
  }

  const statusOptions = [
    STATUS.VALID, STATUS.PENDING, STATUS.NA,
    STATUS.CHECK_BONUS, STATUS.NO_BONUS,
    ...Object.values(DOC_LABEL_BY_KEY).map(l => `Falta: ${l}`)
  ];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statusOptions, true)
    .setAllowInvalid(false).build();

  sheet.getRange(`L2:L${sheet.getMaxRows()}`).setDataValidation(rule);
}

// =================================================================
// === EVENTOS =====================================================
// =================================================================

/** Evento ao editar DOCUMENTAÇÃO */
function onEdit(e) {
  if (!e) return;
  const sheet = e.range.getSheet();
  const row = e.range.getRow(), col = e.range.getColumn();

  if (sheet.getName() !== SHEET_NAMES.DOCUMENTACAO || row === 1) return;

  if (col === 10) {
    toastMsg(`Linha ${row}: processando Bônus...`);
    processarBonus_(sheet, row);
    toastMsg(`Linha ${row}: concluído`);
  } else if (col === 12) {
    toastMsg(`Linha ${row}: processando Documentação...`);
    if (!e.value || e.value.trim() === "") {
      sheet.getRange(row, 13, 1, 12).clearContent();
      aplicarFormatacao_(sheet, row, "");
      toastMsg(`Linha ${row}: colunas M:X limpas (L vazio)`);
    } else {
      processarDropdown_(sheet, row, e.value);
    }
    toastMsg(`Linha ${row}: concluído`);
  }
}

/** Validação periódica de documentação */
function validarDocumentacaoPeriodicamente() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DOCUMENTACAO);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const colStatus = sheet.getRange(2, 12, lastRow - 1).getValues();
  const colDocs = sheet.getRange(2, 13, lastRow - 1).getValues();

  colStatus.forEach((row, i) => {
    if (row[0] && !colDocs[i][0]) processarDropdown_(sheet, i + 2, row[0]);
  });
}

// =================================================================
// === FUNÇÕES AUXILIARES ==========================================
// =================================================================

function mostrarInfo() {
  SpreadsheetApp.getUi().alert(
    "🚀 Painel Apollo\n\n" +
    "Author: Felipe Marcondes\n" +
    "Co-Author: Giovanni Muller\n" +
    "Versão Final 1.0"
  );
}

function normalizeText(s) {
  return String(s || "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().trim();
}

function toastMsg(msg) {
  SpreadsheetApp.getActive().toast(msg, "Painel Apollo", 3);
}

// =================================================================
// === PROCESSAMENTO DOCUMENTAÇÃO ==================================
// =================================================================

function getRequiredDocs_(bonusValue) {
  const str = String(bonusValue || "").trim();
  if (!str) return { type: "empty", required: new Set() };

  const upper = str.toUpperCase();
  if (upper.includes("SEM BÔNUS") || upper.includes("VENDA DIRETA"))
    return { type: "no_bonus", required: new Set() };

  const tokens = str.split(",").map(t => t.trim()).filter(Boolean);
  const req = new Set();
  tokens.forEach(t => {
    const key = BONUS_SYNONYMS[normalizeText(t)] || normalizeText(t).replace(/\s+/g, "_");
    if (BONUS_DOC_MAP[key]) BONUS_DOC_MAP[key].forEach(d => req.add(d));
  });

  return { type: "regular", required: req };
}

function processarBonus_(sheet, row) {
  const bonusVal = sheet.getRange(row, 10).getValue();
  const { type, required } = getRequiredDocs_(bonusVal);
  let status, vals;

  if (type === "empty") {
    status = STATUS.CHECK_BONUS;
    vals = preencherArrayDocs_(null, "ALL_RED");
  } else if (type === "no_bonus") {
    status = STATUS.NO_BONUS;
    vals = preencherArrayDocs_(new Set(), "NA_ONLY");
  } else {
    vals = preencherArrayDocs_(required, "REQUIRED_PENDING");
    status = todosDocsValidos_(vals, required) ? STATUS.VALID : STATUS.PENDING;
    if (status === STATUS.VALID) vals = preencherArrayDocs_(required, "ALL_VALID");
  }

  sheet.getRange(row, 12).setValue(status);
  sheet.getRange(row, 13, 1, 12).setValues([vals]);
  aplicarFormatacao_(sheet, row, status);
}

function processarDropdown_(sheet, row, value) {
  const edited = value.trim();
  const bonusVal = sheet.getRange(row, 10).getValue();
  const { type, required } = getRequiredDocs_(bonusVal);

  let vals;
  if (edited === STATUS.CHECK_BONUS) {
    vals = preencherArrayDocs_(null, "ALL_RED");
    sheet.getRange(row, 12).setValue(STATUS.CHECK_BONUS);
  } else if (edited === STATUS.VALID) {
    vals = preencherArrayDocs_(required, "ALL_VALID");
  } else if (edited === STATUS.PENDING) {
    vals = preencherArrayDocs_(required, "REQUIRED_PENDING");
  } else if (edited.startsWith(STATUS.MISSING_PREFIX)) {
    const key = NORMAL_LABEL_TO_DOC_KEY[normalizeText(edited.replace(/^Falta:\s*/i, ""))];
    vals = preencherArrayDocs_(required, "ONE_MISSING", key);
  } else if ([STATUS.NA, STATUS.NO_BONUS].includes(edited) || type === "no_bonus") {
    vals = preencherArrayDocs_(new Set(), "NA_ONLY");
  }

  sheet.getRange(row, 13, 1, 12).setValues([vals]);
  aplicarFormatacao_(sheet, row, edited);
}

function preencherArrayDocs_(requiredSet, mode, missingKey) {
  const arr = [];
  requiredSet = requiredSet || new Set();

  for (const col of ALL_DOC_COLUMNS) {
    const key = COL_TO_DOC_KEY[col];
    switch (mode) {
      case "ALL_RED": arr.push(CELLMARK.PENDING_ALL_RED); break;
      case "NA_ONLY": arr.push(CELLMARK.NA); break;
      case "ALL_VALID": arr.push(requiredSet.has(key) ? CELLMARK.VALID : CELLMARK.NA); break;
      case "REQUIRED_PENDING": arr.push(requiredSet.has(key) ? CELLMARK.PENDING : CELLMARK.NA); break;
      case "ONE_MISSING": arr.push(requiredSet.has(key) ? (key === missingKey ? CELLMARK.PENDING : CELLMARK.VALID) : CELLMARK.NA); break;
    }
  }
  return arr;
}

function todosDocsValidos_(vals, requiredSet) {
  for (let i = 0; i < ALL_DOC_COLUMNS.length; i++) {
    const key = COL_TO_DOC_KEY[ALL_DOC_COLUMNS[i]];
    if (requiredSet.has(key)) {
      const v = String(vals[i] || "").toUpperCase().trim();
      if (!(v.startsWith("VALIDADO!") || v === "OK!")) return false;
    }
  }
  return true;
}

function aplicarFormatacao_(sheet, row, status) {
  const vals = sheet.getRange(row, 13, 1, 12).getValues()[0];
  const bgs = new Array(12).fill(COLORS.WHITE);

  vals.forEach((val, i) => {
    if (status === STATUS.VALID) {
      if (val.startsWith("Validado!")) bgs[i] = COLORS.GREEN;
      else if (val.startsWith("Não se Aplica")) bgs[i] = COLORS.GREY;
    } else if (status === STATUS.CHECK_BONUS) {
      vals[i] = CELLMARK.PENDING_ALL_RED; bgs[i] = COLORS.RED;
    } else if ([STATUS.PENDING, STATUS.MISSING_PREFIX].some(s => status.startsWith(s))) {
      if (val.startsWith("Pendente!")) bgs[i] = COLORS.YELLOW;
      else if (val.startsWith("Validado!")) bgs[i] = COLORS.GREEN;
      else if (val.startsWith("Não se Aplica")) bgs[i] = COLORS.GREY;
    } else if (status === STATUS.NO_BONUS) {
      if (val.startsWith("Não se Aplica")) bgs[i] = COLORS.GREY;
      else if (val.startsWith("Validado!")) bgs[i] = COLORS.BLUE;
    } else if (status === STATUS.NA) {
      bgs[i] = COLORS.GREY;
    }
  });

  sheet.getRange(row, 13, 1, 12).setValues([vals]);
  sheet.getRange(row, 13, 1, 12).setBackgrounds([bgs]);
}


/**
 * Painel de Operações - Automação Apollo
 * 
 * Author: Felipe Marcondes
 * Co-Author: Giovanni Muller
 * Versão Final 2.0
 */
