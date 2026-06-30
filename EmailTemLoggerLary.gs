function enviarEmailsGestorLary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  // 📅 Mapeia meses para abreviações em minúsculas
  const mesesAbrev = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
  const mesAtualIndex = hoje.getMonth();
  const anoAtualFull = hoje.getFullYear();
  const anoAtual = String(anoAtualFull).slice(-2);

  const mesAtual = mesesAbrev[mesAtualIndex];
  const proximoMes = mesesAbrev[(mesAtualIndex + 1) % 12];
  const proximoAno = mesAtualIndex === 11 ? String(anoAtualFull + 1).slice(-2) : anoAtual;

  const abasParaVerificar = [`${mesAtual}/${anoAtual}`, `${proximoMes}/${proximoAno}`];

  // 🚫 Lista de cargos bloqueados
  const cargosBloqueados = [
    "ENCARREGADO ALMOXARIFADO",
    "ENCARREGADO LOGÍSTICA",
    "ENCARREGADO MANUTENÇÃO",
    "ENCARREGADO SUPORTE E OPERAÇÕES",
    "ASSISTENTE LOGÍSTICA",
    "AJUDANTE MOTORISTA",
    "ASSISTENTE DE LOGISTICA",
    "AUTONOMO - MOTORISTA",
    "AUXILIAR ALMOXARIFADO",
    "AUXILIAR ALMOXARIFADO II",
    "AUXILIAR DESPACHO AÉREO II",
    "AUXILIAR LOGÍSTICA (26 HORAS)",
    "AUXILIAR LOGÍSTICA I",
    "AUXILIAR LOGÍSTICA II",
    "AUXILIAR LOGÍSTICA III",
    "AUXILIAR MANUTENÇÃO",
    "AUXILIAR PÁTIO",
    "MOTORISTA",
    "JOVEM APRENDIZ ADMINISTRATIVO",
    "JOVEM APRENDIZ OPERAÇÃO",
    "JOVEM APRENDIZ DE ADMINISTRAÇÃO LOGÍSTICA",
    "OPERADOR EMPILHADEIRA"
  ].map(c => c.toUpperCase().trim());

  // 🚀 OTIMIZAÇÃO: Busca as imagens do Drive APENAS UMA VEZ fora do loop
  const logoTopoId = "118W37ZpXSk4QrmTxY9SoZKlfP62EEVl5";
  let logoTopoBlob, logoRodapeBlob;
  try {
    logoTopoBlob = DriveApp.getFileById(logoTopoId).getBlob();
    logoRodapeBlob = logoTopoBlob;
  } catch(e) {
    Logger.log("❌ CRÍTICO: Falha ao carregar imagens do Drive. Detalhe: " + e.message);
    return;
  }

  // 🛡️ Controle de controle de duplicados
  const enviadosHoje = new Set();

  // 🔁 Percorre as abas atuais e do próximo mês
  abasParaVerificar.forEach(nomeAba => {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) {
      Logger.log("ℹ️ Aba não encontrada (ignorada): " + nomeAba);
      return;
    }

    Logger.log("📋 Iniciando leitura da aba: " + nomeAba);

    const dados = aba.getDataRange().getValues();
    const cores = aba.getDataRange().getBackgrounds();
    
    let enviosSucesso = 0;

    // 📨 Percorre todas as linhas da aba
    for (let i = 1; i < dados.length; i++) {
      const cargo = (dados[i][15] || "").toString().trim().toUpperCase(); // Coluna P
      const tipoTrabalho = (dados[i][10] || "").toString().trim().toLowerCase(); // Coluna K
      const corCelulaI = cores[i][8].toLowerCase(); // Coluna I
      const corColunaK = cores[i][10].toLowerCase(); // Coluna K
      const nomeLogger = (dados[i][5] || "").toString().trim();   // Coluna F

      // 1. FILTROS SILENCIOSOS
      if (cargosBloqueados.includes(cargo)) continue;
      if (tipoTrabalho !== "remoto") {
        aba.getRange(i + 1, 24).setValue("Ignorado - não é regular remoto");
        continue;
      }

      if (corCelulaI !== "#ffffff" && corCelulaI !== "white") continue;

      if (corColunaK !== "#ffffff" && corColunaK !== "white") {
        aba.getRange(i + 1, 24).setValue("Ignorado - coluna K colorida");
        continue;
      }

      const emailGestor = (dados[i][18] || "").toString().trim(); // Coluna S

      if (!emailGestor || !nomeLogger) {
        Logger.log(`⚠️ Alerta: Linha ${i + 1} qualificada, mas falta e-mail do gestor ou nome do Logger.`);
        continue;
      }

      // 2. VERIFICAÇÃO DE DATAS
      const dataAlternativa = dados[i][4]; // Coluna E
      const dataAdmissaoOriginal = dados[i][13]; // Coluna N
      let dataReferencia = null;

      if (dataAlternativa) {
        dataReferencia = new Date(dataAlternativa);
      } else if (dataAdmissaoOriginal) {
        dataReferencia = new Date(dataAdmissaoOriginal);
      } else {
        Logger.log(`⚠️ Alerta [Gestor: ${emailGestor}]: Logger ${nomeLogger} está sem data de admissão.`);
        continue;
      }

      dataReferencia.setHours(0, 0, 0, 0);

      // 📆 Envia 6 dias antes da data de referência
      const dataEnvio = new Date(dataReferencia);
      dataEnvio.setDate(dataEnvio.getDate() - 6);

      if (hoje.getTime() !== dataEnvio.getTime()) continue;

      // Controle de duplicidade na execução
      const chaveEnvio = `${emailGestor}_${nomeLogger}_${dataReferencia.toDateString()}`;
      if (enviadosHoje.has(chaveEnvio)) continue;
      enviadosHoje.add(chaveEnvio);

      // 🗓️ Datas do Xboarding formatadas
      const xboarding1 = new Date(dataReferencia);
      const xboarding2 = new Date(dataReferencia);
      xboarding2.setDate(xboarding2.getDate() + 1);

      const formatarDataResumida = d => {
        const dd = String(d.getDate()).padStart(2, '0');
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        return `${dd}/${mm}`;
      };

      // 🌟 NOVO: Extrai com segurança o primeiro nome em MAIÚSCULO para o assunto
      const primeiroNomeLogger = nomeLogger.split(" ")[0].toUpperCase();

      // 3. ENVIO DO EMAIL
      try {
        // Atenção: O assunto agora usa CRASE ( ` ) para injetar a variável corretamente
        const assunto = `🚨 Tem Logger chegando no seu time | ${primeiroNomeLogger}`;
        
        const corpoHtml = HtmlService.createTemplateFromFile('templateTemLoggerLary');
        corpoHtml.nomeLogger = nomeLogger; // Aqui continua passando o nome completo para o corpo do HTML
        corpoHtml.primeiroNomeLogger = primeiroNomeLogger;
        corpoHtml.xboarding1 = formatarDataResumida(xboarding1);
        corpoHtml.xboarding2 = formatarDataResumida(xboarding2);
        const htmlFinal = corpoHtml.evaluate().getContent();

        MailApp.sendEmail({
          to: emailGestor,
          subject: assunto,
          htmlBody: htmlFinal,
          inlineImages: { logoTopo: logoTopoBlob, logoRodape: logoRodapeBlob }
        });

        aba.getRange(i + 1, 24).setValue('E-mail enviado');
        Logger.log(`✅ Sucesso [Gestor: ${emailGestor}]: Notificação enviada para o Logger: ${primeiroNomeLogger}.`);
        enviosSucesso++;
      } catch(e) {
        Logger.log(`❌ Erro [Gestor: ${emailGestor}]: Falha ao notificar sobre ${nomeLogger}. Detalhe: ` + e.message);
        aba.getRange(i + 1, 24).setValue("Erro ao enviar");
      }
    }
    Logger.log(`🏁 Resumo da aba ${nomeAba}: Total de gestores notificados hoje: ${enviosSucesso}`);
  });
}
