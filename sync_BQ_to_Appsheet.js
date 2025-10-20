    // ======================================================
    // FUNÇÃO PRINCIPAL 
    // ======================================================

function syncBigQueryToAppSheet() {
  const startTime = new Date();
  // Função helper movida para dentro para manter tudo junto
  const formatTime = () => Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");

  try {
    logHeader('INICIANDO SINCRONIZAÇÃO BIGQUERY → APPSHEET');

    // ======================================================
    // 🔧 CONFIGURAÇÕES PRINCIPAIS
    // ======================================================
    const CONFIG = {
      PROJECT_ID: 'projeto-looker-servopa', // ID DO PROJETO NO GOOGLE CLOUD - GCP
      DATASET_ID: 'silver',  // DATASET DE ONDE ESTÁ A TABLE
      TABLE_ID_BQ: 'LK_UNION_SILVER_VENDAS_VEICULOS', // NOME DA TABLE ONDE ESTÁ AS INFORMAÇÕES
      APP_ID: '51243e20-3be5-4784-81fa-8da217c096e8',  // CÓDIGO IDENTIFICADOR DO APLICATIVO
      TABLE_NAME_APPSHEET: 'DB_VENDAS_BQ', // NOME DA TABLE QUE RECEBE AS INFORMAÇÕES DO GCP
      ACCESS_KEY: 'V2-6T2yT-nPXML-uljnX-dNYN5-o6HHX-klD24-K7YM0-zUGMc', //ACESS KEY DA API DO APPSHEET
      BATCH_SIZE: 500  // QUANTIDADE DE REGISTROS POR LOTE
    };

    const APPSHEET_URL = `https://api.appsheet.com/api/v2/apps/${CONFIG.APP_ID}/tables/${CONFIG.TABLE_NAME_APPSHEET}/Action`; // ENDPOINT - PARA GRAVAR SEMPRE USAR /ACTION 

    // ======================================================
    // 🧠 CONSULTA AO BIGQUERY
    // ======================================================
    logSection('CONSULTA AO BIGQUERY');

    const query = `
      SELECT *
      FROM \`${CONFIG.PROJECT_ID}.${CONFIG.DATASET_ID}.${CONFIG.TABLE_ID_BQ}\`
      WHERE DATE(DTA_PROCESSAMENTO) > '2025-01-01'
        AND CAST(EMPRESA AS STRING) = '25'
        AND CAST(USADO AS INT64) = 0
    `; // AQUI ENCIMA É ONDE FICA OS FILTROS QUE PRECISO DO BQ, DATA DE VENDA, CÓD EMPRESA E SE É USADO OU VEÍCULO NOVO

    const request = { query, useLegacySql: false };
    const queryResults = BigQuery.Jobs.query(request, CONFIG.PROJECT_ID);
    const rows = queryResults.rows || [];

    if (!rows.length) {
      Logger.log(`Nenhum registro encontrado (${formatTime()}).`);
      return;
    }

    Logger.log(`${rows.length.toLocaleString()} registros retornados (${formatTime()})`);

    // ======================================================
    // 🔄 TRATAMENTO E MONTAGEM DO JSON
    // ======================================================
    logSection('TRATAMENTO E MONTAGEM DO JSON');

    const headers = queryResults.schema.fields.map(f => f.name);
    const jsonData = rows.map(r => {
      const obj = {};
      r.f.forEach((cell, i) => {
        obj[headers[i]] = cell?.v ?? null;
      });
      return obj;
    });

    Logger.log(`Total de registros prontos para envio: ${jsonData.length}`);

    // ======================================================
    // 🧩 DIVISÃO EM LOTES
    // ======================================================
    logSection('PREPARAÇÃO DOS LOTES');
    const batches = chunkArray(jsonData, CONFIG.BATCH_SIZE);
    Logger.log(`Total de lotes: ${batches.length} (aprox. ${CONFIG.BATCH_SIZE} registros/lote)`);

    // ======================================================
    // 🚀 ENVIO DOS DADOS PARA O APPSHEET (EM PARALELO)
    // ======================================================
    logSection('PREPARANDO REQUISIÇÕES PARALELAS');

    // Primeiro Passo - Criar um array de "requests" para o fetchAll
    const allRequests = batches.map((batch, i) => {
      const payload = {
        Action: 'Add',
        Rows: batch
      };

      Logger.log(`Preparando Lote ${i + 1}/${batches.length} com ${batch.length} registros.`);

      return {
        url: APPSHEET_URL,
        method: 'post',
        headers: {
          'ApplicationAccessKey': CONFIG.ACCESS_KEY,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true // Essencial para o fetchAll
      };
    });

    // Segundo Passo - Dispara TODAS as requisições de uma vez
    logSection('ENVIANDO DADOS (TODOS OS LOTES EM PARALELO)');
    const fetchStartTime = new Date();
    Logger.log(`[${formatTime()}] Enviando ${allRequests.length} lotes...`);

    const responses = UrlFetchApp.fetchAll(allRequests);

    const elapsedFetch = ((new Date() - fetchStartTime) / 1000).toFixed(1);
    Logger.log(`[${formatTime()}] Respostas recebidas em ${elapsedFetch}s.`);

    // Terceiro Passo - Verificar as respostas
    let hasErrors = false;
    responses.forEach((response, i) => {
      const status = response.getResponseCode();
      if (status === 200) {
        Logger.log(`Lote ${i + 1}: SUCESSO (HTTP ${status})`);
      } else {
        hasErrors = true;
        const body = response.getContentText();
        Logger.log(`Lote ${i + 1}: FALHA (HTTP ${status})`);
        Logger.log(`Retorno (Lote ${i + 1}): ${body.substring(0, 500)}...`);
      }
    });

    if (hasErrors) {
      throw new Error('Uma ou mais falhas ocorreram durante o envio dos lotes ao AppSheet.');
    }

    // ======================================================
    // ✅ FINALIZAÇÃO
    // ======================================================
    logSection('FINALIZAÇÃO');
    const totalTime = ((new Date() - startTime) / 1000 / 60).toFixed(2);
    Logger.log(`Execução concluída com sucesso em ${totalTime} min`);

  } catch (err) {
    Logger.log(`ERRO CRÍTICO: ${err.message}\n${err.stack}`);
  }
}

/**
 * ============================================================
 * 🔹 Função auxiliar: Divide um array em blocos (lotes)
 * ============================================================
 */
function chunkArray(array, size) {
  const result = [];
  for (let i = 0, len = array.length; i < len; i += size) {
    result.push(array.slice(i, i + size));
  }
  return result;
}

/**
 * ============================================================
 * 🧭 Funções de LOG (simplificadas)
 * ============================================================
 */
function logHeader(title) {
  Logger.log(`\n=== ${title} ===`);
}

function logSection(title) {
  Logger.log(`\n--- ${title} ---`);
}

/**
 * ============================================================
 * SINCRONIZAÇÃO BigQuery → AppSheet (API v2 - Endpoint /Action)
 * ============================================================
 * ✅ Lê dados do BigQuery e grava no AppSheet via API v2
 * ============================================================
 * Autor: Felipe Marcondes
 * Co-Author: Giovanni Müller
 * Versão: 1.0 - 20/10/2025
 * FUNCIONANDO PORRA!!!
 * ============================================================
 */
