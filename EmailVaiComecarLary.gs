function enviarEmailVaiComecarLary() {
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

  // 🔁 Percorre as abas mapeadas
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
      const emailLogger = (dados[i][6] || "").toString().trim(); // Coluna G

      // 1. FILTROS SILENCIOSOS
      if (cargosBloqueados.includes(cargo)) continue;
      if (tipoTrabalho !== "remoto") {
        aba.getRange(i + 1, 24).setValue("Ignorado - não é remoto");
        continue;
      }

      if (corCelulaI !== "#ffffff" && corCelulaI !== "white") continue;

      if (corColunaK !== "#ffffff" && corColunaK !== "white") {
        aba.getRange(i + 1, 24).setValue("Ignorado - coluna K colorida");
        continue;
      }

      if (!emailLogger) {
        Logger.log(`⚠️ Alerta: Linha ${i + 1} é remota mas está sem e-mail do Logger cadastrado.`);
        continue;
      }

      // 2. VERIFICAÇÃO DE DATAS (E > N)
      const dataAlternativa = dados[i][4]; // Coluna E
      const dataAdmissaoOriginal = dados[i][13]; // Coluna N
      let dataAdmissao;

      if (dataAlternativa) {
        dataAdmissao = new Date(dataAlternativa);
      } else if (dataAdmissaoOriginal) {
        dataAdmissao = new Date(dataAdmissaoOriginal);
      } else {
        Logger.log(`⚠️ Alerta [${emailLogger}]: Sem data de admissão preenchida.`);
        continue;
      }

      dataAdmissao = new Date(dataAdmissao.getFullYear(), dataAdmissao.getMonth(), dataAdmissao.getDate(), 0, 0, 0, 0);
      
      // 📆 REGRA: Envia EXATAMENTE no mesmo dia da admissão
      const dataEnvio = new Date(dataAdmissao);

      if (hoje.getTime() !== dataEnvio.getTime()) continue;

      // 🗓️ Formatação padrão de data
      const formatarData = (d) => {
        const dd = String(d.getDate()).padStart(2, "0");
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        const yyyy = d.getFullYear();
        return `${dd}/${mm}/${yyyy}`;
      };

      // 3. ENVIO DO EMAIL
      try {
        const assunto = "Nosso Xboarding já vai começar! 🐇💙";
        const corpoHtml = HtmlService.createTemplateFromFile("templateVaiComecarLary");
  
        // 🌟 NOVO: Passa a data completa e a data resumida (DD/MM) para o HTML
        const dataCompleta = formatarData(dataAdmissao);
        corpoHtml.dataAdmissao = dataCompleta;
        corpoHtml.dataCurta = dataCompleta.slice(0, 5); // Corta o "/YYYY" e mantém apenas "DD/MM"

        const htmlFinal = corpoHtml.evaluate().getContent();

        MailApp.sendEmail({
          to: emailLogger,
          subject: assunto,
          htmlBody: htmlFinal,
          inlineImages: { logoTopo: logoTopoBlob, logoRodape: logoRodapeBlob }
        });

        // 📝 Marca na coluna X
        aba.getRange(i + 1, 24).setValue("templateVaiComecarLary enviado");
        Logger.log(`✅ Sucesso [${emailLogger}]: E-mail 'Vai Começar' enviado com sucesso.`);
        enviosSucesso++;
      } catch(e) {
        Logger.log(`❌ Erro [${emailLogger}]: Falha no envio do e-mail de início. Detalhe: ` + e.message);
        aba.getRange(i + 1, 24).setValue("Erro ao enviar");
      }
    }
    Logger.log(`🏁 Resumo da aba ${nomeAba}: Total de e-mails de início enviados: ${enviosSucesso}`);
  });
}
