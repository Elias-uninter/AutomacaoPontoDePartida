function enviarEmailsBoasVindasLary() {
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

  // 🚀 OTIMIZAÇÃO: Busca os assets específicos do Drive APENAS UMA VEZ fora das leituras
  const logoTopoId = "1ZLBjve06FDqCjFXq7c6AKeAnyxTqw2q8";
  const logoRodapeId = "10Z9irqXv1zmBlFJrMG7pYa06EdzDcfvn";
  let logoTopoBlob, logoRodapeBlob;
  try {
    logoTopoBlob = DriveApp.getFileById(logoTopoId).getBlob();
    logoRodapeBlob = DriveApp.getFileById(logoRodapeId).getBlob();
  } catch(e) {
    Logger.log("❌ CRÍTICO: Falha ao carregar imagens do Drive. Detalhe: " + e.message);
    return;
  }

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
      const corCelulaI = cores[i][8].toLowerCase(); // Coluna I
      const corColunaK = cores[i][10].toLowerCase(); // Coluna K
      const emailLogger = (dados[i][6] || "").toString().trim(); // Coluna G

      // 1. FILTROS SILENCIOSOS (Removida a validação da string "remoto")
      if (cargosBloqueados.includes(cargo)) continue;
      if (corCelulaI !== "#ffffff" && corCelulaI !== "white") continue;

      if (corColunaK !== "#ffffff" && corColunaK !== "white") {
        aba.getRange(i + 1, 24).setValue("Ignorado - coluna K colorida");
        continue;
      }

      if (!emailLogger) {
        Logger.log(`⚠️ Alerta: Linha ${i + 1} qualificada para Boas-vindas, mas sem e-mail.`);
        continue;
      }

      // 2. VERIFICAÇÃO DE DATAS (E > N)
      const dataE = dados[i][4];  // Coluna E
      const dataN = dados[i][13]; // Coluna N
      let dataBase;

      if (dataE) {
        dataBase = new Date(dataE);
      } else if (dataN) {
        dataBase = new Date(dataN);
      } else {
        Logger.log(`⚠️ Alerta [${emailLogger}]: Sem data de admissão mapeável.`);
        aba.getRange(i + 1, 24).setValue("Ignorado - sem data válida");
        continue;
      }

      dataBase.setHours(0, 0, 0, 0);

      // 📅 Envio programado para exatamente 7 dias antes da dataBase
      const dataEnvio = new Date(dataBase);
      dataEnvio.setDate(dataEnvio.getDate() - 7);

      // Se não for o dia exato do gatilho de Boas-vindas, avança em silêncio
      if (hoje.getTime() !== dataEnvio.getTime()) continue;

      // 🗓️ Geração das datas dinâmicas consumidas por este template específico
      const dataPreenchimento = new Date(hoje);
      dataPreenchimento.setDate(dataPreenchimento.getDate() + 1);
      const xboarding1 = new Date(dataBase);
      const xboarding2 = new Date(dataBase);
      xboarding2.setDate(xboarding2.getDate() + 1);

      // Formatação padrão DD/MM
      const formatarDataResumida = d => {
        const dd = String(d.getDate()).padStart(2, '0');
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        return `${dd}/${mm}`;
      };

      const formatarDiaSemana = d => {
        const diasSemana = ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"];
        return diasSemana[d.getDay()];
      };

      // 3. ENVIO DO EMAIL
      try {
        const assunto = "Logger, boas-vindas ao nosso time! 🐇✨";
        const corpoHtml = HtmlService.createTemplateFromFile('templateBoasVindasLary');
        
        // Atribuição de variáveis ao escopo do HTML
        corpoHtml.dataPreenchimento = formatarDataResumida(dataPreenchimento);
        corpoHtml.dataRastreio = formatarDataResumida(dataCodigoRastreio);
        corpoHtml.xboarding1 = formatarDataResumida(xboarding1);
        corpoHtml.xboarding2 = formatarDataResumida(xboarding2);
        corpoHtml.diaSemana1 = formatarDiaSemana(xboarding1);
        corpoHtml.diaSemana2 = formatarDiaSemana(xboarding2);
        
        const htmlFinal = corpoHtml.evaluate().getContent();

        MailApp.sendEmail({
          to: emailLogger,
          subject: assunto,
          htmlBody: htmlFinal,
          inlineImages: {
            logoTopo: logoTopoBlob,
            logoRodape: logoRodapeBlob
          }
        });

        aba.getRange(i + 1, 24).setValue('E-mail de boas-vindas enviado');
        Logger.log(`✅ Sucesso [${emailLogger}]: E-mail de Boas-Vindas enviado.`);
        enviosSucesso++;
      } catch(e) {
        Logger.log(`❌ Erro [${emailLogger}]: Falha ao processar e-mail de Boas-vindas. Detalhe: ` + e.message);
        aba.getRange(i + 1, 24).setValue("Erro ao enviar");
      }
    }
    Logger.log(`🏁 Resumo da aba ${nomeAba}: Total de e-mails de Boas-Vindas enviados: ${enviosSucesso}`);
  });
}
