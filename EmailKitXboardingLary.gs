function enviarKitXBoardingLary() {
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

  const logoTopoId = "118W37ZpXSk4QrmTxY9SoZKlfP62EEVl5";
  let logoTopoBlob, logoRodapeBlob;
  try {
    logoTopoBlob = DriveApp.getFileById(logoTopoId).getBlob();
    logoRodapeBlob = logoTopoBlob;
  } catch(e) {
    Logger.log("❌ CRÍTICO: Falha ao carregar imagens do Drive. Detalhe: " + e.message);
    return;
  }

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

    for (let i = 1; i < dados.length; i++) {
      const emailLogger = (dados[i][6] || "").toString().trim();   // Coluna G
      const cargo = (dados[i][15] || "").toString().trim().toUpperCase(); // Coluna P
      const tipoTrabalho = (dados[i][10] || "").toString().trim().toLowerCase(); // Coluna K
      const corCelulaI = cores[i][8].toLowerCase(); // Coluna I
      const corColunaK = cores[i][10].toLowerCase(); // Coluna K
      
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

      // Se não houver e-mail preenchido para esse remoto elegível, avisa no log
      if (!emailLogger) {
        Logger.log(`⚠️ Alerta: Linha ${i + 1} é remota mas está sem e-mail cadastrado.`);
        continue;
      }

      // 2. VERIFICAÇÃO DE DATAS
      const dataAlternativa = dados[i][4]; // Coluna E
      const dataAdmissaoOriginal = dados[i][13]; // Coluna N
      let dataAdmissao;

      if (dataAlternativa) {
        dataAdmissao = new Date(dataAlternativa);
      } else if (dataAdmissaoOriginal) {
        dataAdmissao = new Date(dataAdmissaoOriginal);
      } else {
        // Sem data nenhuma? Loga focado no e-mail
        Logger.log(`⚠️ Alerta [${emailLogger}]: Sem data de admissão preenchida.`);
        continue;
      }
      
      dataAdmissao = new Date(dataAdmissao.getFullYear(), dataAdmissao.getMonth(), dataAdmissao.getDate(), 0, 0, 0, 0);
      const dataEnvio = new Date(dataAdmissao);
      dataEnvio.setDate(dataEnvio.getDate() - 5);

      // Se não for o dia do envio desse Logger, pula silenciosamente
      if (hoje.getTime() !== dataEnvio.getTime()) continue;

      // 3. REGRA DO CÓDIGO DE RASTREIO (Última validação antes de disparar)
      const codigoRastreio = (dados[i][9] || "").toString().trim(); // Coluna J
      if (!codigoRastreio) {
        Logger.log(`⚠️ Pendência [${emailLogger}]: Dia de envio atingido, mas código de rastreio está em branco.`);
        aba.getRange(i + 1, 24).setValue("Erro - Sem código de rastreio");
        continue;
      }

      // 4. ENVIO DO EMAIL
      const formatarData = (d) => {
        const dd = String(d.getDate()).padStart(2, "0");
        const mm = String(d.getMonth() + 1).padStart(2, "0");
        const yyyy = d.getFullYear();
        return `${dd}/${mm}/${yyyy}`;
      };

      try {
        const assunto = "📦 Seu kit de XBoarding já está na rota!";
        const corpoHtml = HtmlService.createTemplateFromFile("templateKitXboardingLary");
        corpoHtml.codigoRastreio = codigoRastreio;
        corpoHtml.dataAdmissao = formatarData(dataAdmissao);
        corpoHtml.dataPrimeiroDia = formatarData(dataAdmissao);
        corpoHtml.linkEncontro = "https://meet.google.com/idi-jjad-wff?authuser=0";
        const htmlFinal = corpoHtml.evaluate().getContent();

        MailApp.sendEmail({
          to: emailLogger,
          subject: assunto,
          htmlBody: htmlFinal,
          inlineImages: { logoTopo: logoTopoBlob, logoRodape: logoRodapeBlob }
        });

        aba.getRange(i + 1, 24).setValue("Kit enviado");
        Logger.log(`✅ Sucesso [${emailLogger}]: E-mail do kit enviado.`);
        enviosSucesso++;
      } catch(e) {
        Logger.log(`❌ Erro [${emailLogger}]: Falha no envio. Detalhe: ` + e.message);
        aba.getRange(i + 1, 24).setValue("Erro ao enviar");
      }
    }
    Logger.log(`🏁 Resumo da aba ${nomeAba}: Total de e-mails enviados hoje: ${enviosSucesso}`);
  });
}
