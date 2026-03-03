/** CONFIGURAÇÕES **/
const SPREADSHEET_ID = "ALTERE PARA O ID DA PLANILHA";
const SHEET_NAME = "NOME DA PLANILHA";
const CONFIG_SHEET_NAME = "Configurações";
const SLIDES_TEMPLATE_ID = "ID DO TEMPLATE SLIDE";
const FOLDER_ID = "ID DA PASTA QUE VAI SALVAR AS CARTEIRINHAS";


/** USUÁRIOS AUTORIZADOS **/
const USUARIOS_AUTORIZADOS = {
  'COLOQUE O NOME DE USUARIO AQUI': 'COLOQUE A SENHA AQUI',
};

function testarPermissoes() {
  Logger.log(GmailApp.getAliases());
}

function quemEhODono() {
  Logger.log("Effective user: " + Session.getEffectiveUser().getEmail());
  Logger.log("Active user: " + Session.getActiveUser().getEmail());
}

/** INTERFACE WEB **/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Gerador de Carteirinha Escolar');
}

/** AUTENTICAÇÃO **/
function autenticar(usuario, senha) {
  try {
    if (USUARIOS_AUTORIZADOS[usuario] && USUARIOS_AUTORIZADOS[usuario] === senha) {
      return { 
        success: true, 
        message: 'Login realizado com sucesso!',
        usuario: usuario
      };
    } else {
      return { 
        success: false, 
        message: 'Usuário ou senha inválidos.'
      };
    }
  } catch (e) {
    Logger.log('Erro em autenticar: ' + e);
    return { 
      success: false, 
      message: 'Erro ao autenticar: ' + e.message 
    };
  }
}

/** INICIALIZA ABA DE CONFIGURAÇÕES **/
function initConfigSheet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    
    if (!configSheet) {
      configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
      
      // Cabeçalhos para datas de validade individuais
      configSheet.getRange('A1:C1').setValues([['Nº Carteirinha', 'Data Validade', 'Última Atualização']]);
      configSheet.getRange('A1:C1').setFontWeight('bold').setBackground('#667eea').setFontColor('#ffffff');
      
      // Cabeçalhos para trajetos adicionais
      configSheet.getRange('E1:G1').setValues([['Nº Carteirinha', 'Trajeto Adicional', 'Observações']]);
      configSheet.getRange('E1:G1').setFontWeight('bold').setBackground('#764ba2').setFontColor('#ffffff');
      
      configSheet.setColumnWidth(1, 150);
      configSheet.setColumnWidth(2, 150);
      configSheet.setColumnWidth(3, 200);
      configSheet.setColumnWidth(5, 150);
      configSheet.setColumnWidth(6, 300);
      configSheet.setColumnWidth(7, 300);
      
      Logger.log('Aba Configurações criada com sucesso');
    }
    
    return { success: true };
  } catch (e) {
    Logger.log('Erro em initConfigSheet: ' + e);
    return { success: false, message: e.message };
  }
}

/** BUSCA DATA DE VALIDADE DE UMA CARTEIRINHA **/
function getDataValidade(numeroCarteirinha) {
  try {
    initConfigSheet();
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    const data = configSheet.getDataRange().getValues();
    
    // Procura pela carteirinha específica
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(numeroCarteirinha)) {
        return String(data[i][1] || '31/12/2025');
      }
    }
    
    return '31/12/2025'; // Padrão
  } catch (e) {
    Logger.log('Erro em getDataValidade: ' + e);
    return '31/12/2025';
  }
}

/** ATUALIZA DATA DE VALIDADE DE UMA CARTEIRINHA **/
function setDataValidade(numeroCarteirinha, novaData) {
  try {
    initConfigSheet();
    
    // Valida formato da data
    if (!validarFormatoData(novaData)) {
      return { success: false, message: 'Formato de data inválido. Use DD/MM/AAAA' };
    }
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    const data = configSheet.getDataRange().getValues();
    
    // Procura se já existe
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(numeroCarteirinha)) {
        configSheet.getRange(i + 1, 2).setValue(novaData);
        configSheet.getRange(i + 1, 3).setValue(new Date());
        return { success: true, message: 'Data de validade atualizada!' };
      }
    }
    
    // Se não existe, adiciona nova linha
    const proximaLinha = data.length + 1;
    configSheet.getRange(proximaLinha, 1, 1, 3).setValues([[
      numeroCarteirinha,
      novaData,
      new Date()
    ]]);
    
    return { success: true, message: 'Data de validade definida!' };
    
  } catch (e) {
    Logger.log('Erro em setDataValidade: ' + e);
    return { success: false, message: 'Erro ao atualizar data: ' + e.message };
  }
}

/** VALIDA FORMATO DA DATA DD/MM/AAAA **/
function validarFormatoData(data) {
  const regex = /^(\d{2})\/(\d{2})\/(\d{4})$/;
  if (!regex.test(data)) return false;
  
  const partes = data.split('/');
  const dia = parseInt(partes[0]);
  const mes = parseInt(partes[1]);
  const ano = parseInt(partes[2]);
  
  if (mes < 1 || mes > 12) return false;
  if (dia < 1 || dia > 31) return false;
  if (ano < 2024 || ano > 2050) return false;
  
  return true;
}

/** BUSCA TRAJETO ADICIONAL **/
function getTrajetoAdicional(numeroCarteirinha) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    
    if (!configSheet) return null;
    
    const data = configSheet.getDataRange().getValues();
    
    // Coluna E (índice 4) contém os números das carteirinhas
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][4]) === String(numeroCarteirinha)) {
        return String(data[i][5] || '');
      }
    }
    
    return null;
  } catch (e) {
    Logger.log('Erro em getTrajetoAdicional: ' + e);
    return null;
  }
}

/** SALVA TRAJETO ADICIONAL **/
function setTrajetoAdicional(numeroCarteirinha, trajetoAdicional, observacoes) {
  try {
    initConfigSheet();
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    const data = configSheet.getDataRange().getValues();
    
    // Procura na coluna E (índice 4)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][4]) === String(numeroCarteirinha)) {
        configSheet.getRange(i + 1, 6).setValue(trajetoAdicional);
        if (observacoes) {
          configSheet.getRange(i + 1, 7).setValue(observacoes);
        }
        return { success: true, message: 'Trajeto adicional atualizado!' };
      }
    }
    
    // Se não existe, adiciona nova linha
    const proximaLinha = data.length + 1;
    configSheet.getRange(proximaLinha, 5, 1, 3).setValues([[
      numeroCarteirinha,
      trajetoAdicional,
      observacoes || ''
    ]]);
    
    return { success: true, message: 'Trajeto adicional cadastrado!' };
    
  } catch (e) {
    Logger.log('Erro em setTrajetoAdicional: ' + e);
    return { success: false, message: 'Erro ao salvar trajeto: ' + e.message };
  }
}

/** BUSCA DADOS DO ALUNO **/
function getStudent(cardNumber) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const backgrounds = sheet.getDataRange().getBackgrounds();

    const idxNumero = headers.indexOf('Documento CPF (somente números)');
    const idxNome = headers.indexOf('Nome Completo');
    const idxEscola = headers.indexOf('Nome da Instituição de Ensino');
    const idxTrajeto = headers.indexOf('Informar Trajeto - IDA e VOLTA');
    const idxFoto = headers.indexOf('FOTO 3x4');
    const idxEmail = headers.indexOf('Endereço de e-mail');
    const idxData = headers.indexOf('Carimbo de data/hora');

    if (idxNumero === -1)
      return { success: false, message: 'Coluna Nº Carteirinha não encontrada.' };

    // Suporte ao formato "CPF:rowIndex" para identificar linha específica quando há CPFs duplicados
    let cpfBusca = String(cardNumber);
    let rowIndexEspecifico = null;
    if (cpfBusca.includes(':')) {
      const partes = cpfBusca.split(':');
      cpfBusca = partes[0];
      rowIndexEspecifico = parseInt(partes[1]);
    }

    let row;
    if (rowIndexEspecifico !== null) {
      // Busca pela linha específica e valida que o CPF bate
      if (rowIndexEspecifico > 0 && rowIndexEspecifico < data.length) {
        const candidata = data[rowIndexEspecifico];
        if (String(candidata[idxNumero]) === cpfBusca){
          row = candidata;
        }
      }
      if (!row) {
        return { success: false, message: 'Cadastro específico não encontrado.' };
      }
    } else {
      const corAprovado = "#00ff00"; // coloque aqui a cor exata

      let encontrados = [];

      for (let i = data.length - 1; i > 0; i--) {
        if (
          String(data[i][idxNumero]) === cpfBusca &&
          backgrounds[i][idxData] === corAprovado
        ) {
          row = data[i];
          break; // para na primeira que encontrar (que será a última enviada)
        }
      }
    }

    if (!row)
      return { success: false, message: 'Carteirinha não encontrada.' };

    const trajetoAdicional = getTrajetoAdicional(cpfBusca);
    let trajetoCompleto = row[idxTrajeto] || '';
    
    if (trajetoAdicional) {
      trajetoCompleto = trajetoCompleto + ' | ' + trajetoAdicional;
    }

    const dataValidade = getDataValidade(cpfBusca);

    return {
      success: true,
      student: {
        nome: row[idxNome] || '',
        numero: String(cpfBusca),
        escola: row[idxEscola] || '',
        trajeto: trajetoCompleto,
        trajetoPrincipal: row[idxTrajeto] || '',
        trajetoAdicional: trajetoAdicional || '',
        foto: row[idxFoto] || '',
        email: idxEmail !== -1 ? row[idxEmail] : '',
        dataValidade: dataValidade
      }
    };
  } catch (e) {
    Logger.log('Erro em getStudent: ' + e);
    return { success: false, message: 'Erro ao buscar aluno: ' + e.message };
  }
}

/** LISTA TODOS OS ALUNOS - OTIMIZADO **/
function getAllStudents() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const backgrounds = sheet.getDataRange().getBackgrounds();

    const idxNumero       = headers.indexOf('Documento CPF (somente números)');
    const idxNome         = headers.indexOf('Nome Completo');
    const idxEscola       = headers.indexOf('Nome da Instituição de Ensino');
    const idxTrajeto      = headers.indexOf('Informar Trajeto - IDA e VOLTA');
    const idxEmail        = headers.indexOf('Endereço de e-mail');
    const idxFoto         = headers.indexOf('FOTO 3x4');
    // Coluna de status de reprova — tenta nome exato e variações comuns
    const idxReprova      = (() => {
      const candidatos = ['Status de reprova', 'Status de Reprova', 'Motivo Reprova', 'Reprova', 'Status Reprova'];
      for (const c of candidatos) {
        const idx = headers.indexOf(c);
        if (idx !== -1) return idx;
      }
      return -1;
    })();

    if (idxNumero === -1)
      return { success: false, message: 'Coluna Nº Carteirinha não encontrada.' };

    // Carrega trajetos e datas de uma vez
    const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    const trajetos = {};
    const datas = {};
    
    if (configSheet) {
      const configData = configSheet.getDataRange().getValues();
      for (let i = 1; i < configData.length; i++) {
        // Datas de validade (colunas A-C)
        if (configData[i][0]) {
          datas[String(configData[i][0])] = String(configData[i][1] || '31/12/2025');
        }
        // Trajetos adicionais (colunas E-G)
        if (configData[i][4]) {
          trajetos[String(configData[i][4])] = String(configData[i][5] || '');
        }
      }
    }

    // Carrega carteirinhas geradas
    const pasta = DriveApp.getFolderById(FOLDER_ID);
    const arquivos = pasta.getFiles();
    const carteirinhas = {};
    
    while (arquivos.hasNext()) {
      const file = arquivos.next();
      const fileName = file.getName();
      
      if (fileName.startsWith('Carteirinha_') && fileName.endsWith('.pdf')) {
        const numeroMatch = fileName.match(/Carteirinha_(\d+)\.pdf/);
        if (numeroMatch) {
          const numero = numeroMatch[1];
          if (!carteirinhas[numero] || file.getLastUpdated() > carteirinhas[numero].date) {
            carteirinhas[numero] = file.getUrl();
          }
        }
      }
    }

    const students = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const corLinha = backgrounds[i][0];
      if (row[idxNumero]) {
        const numeroCart = String(row[idxNumero] || '');
        const trajetoAdicional = trajetos[numeroCart] || '';
        let trajetoCompleto = String(row[idxTrajeto] || '');
        
        if (trajetoAdicional) {
          trajetoCompleto = trajetoCompleto + ' | ' + trajetoAdicional;
        }

        // Foto: guarda o valor bruto (link do Drive)
        const fotoRaw = idxFoto !== -1 ? String(row[idxFoto] || '') : '';

        // Status de reprova
        const statusReprova = idxReprova !== -1 ? String(row[idxReprova] || '') : '';

        if (corLinha !== '#00ff00') continue;
        
        students.push({
          numero: numeroCart,
          rowIndex: i, // índice real na planilha (começa em 1 para dados)
          nome: String(row[idxNome] || ''),
          escola: String(row[idxEscola] || ''),
          trajeto: trajetoCompleto,
          trajetoPrincipal: String(row[idxTrajeto] || ''),
          trajetoAdicional: trajetoAdicional,
          email: idxEmail !== -1 ? String(row[idxEmail] || '') : '',
          dataValidade: datas[numeroCart] || '31/12/2025',
          carteirinhaUrl: carteirinhas[numeroCart] || null,
          foto: fotoRaw,
          statusReprova: statusReprova
        });
      }
    }

    return { success: true, students: students};
  } catch (e) {
    Logger.log('Erro em getAllStudents: ' + e);
    return { success: false, message: 'Erro ao carregar lista: ' + e.message };
  }
}

/** GERA A CARTEIRINHA **/
function generateCard(cardNumber, dataValidadeCustom) {
  try {
    const studentResult = getStudent(cardNumber);
    if (!studentResult.success) return studentResult;
    const st = studentResult.student;

    // Usa data customizada ou a já cadastrada
    let dataValidade = dataValidadeCustom || st.dataValidade;
    
    // Valida formato
    if (!validarFormatoData(dataValidade)) {
      return { success: false, message: 'Data de validade inválida. Use DD/MM/AAAA' };
    }
    
    // Salva a data de validade
    setDataValidade(cardNumber, dataValidade);

    const copia = DriveApp.getFileById(SLIDES_TEMPLATE_ID)
      .makeCopy(`Carteirinha_${st.numero}_${new Date().getTime()}`);
    const pres = SlidesApp.openById(copia.getId());
    const slide = pres.getSlides()[0];

    // Substitui placeholders de texto
    const placeholders = {
      '{{NOME}}': String(st.nome || '').toUpperCase(),
      '{{NUMERO}}': String(st.numero || ''),
      '{{ESCOLA}}': String(st.escola || '').toUpperCase(),
      '{{TRAJETO}}': String(st.trajeto || '').toUpperCase(),
      '{{VALIDADE}}': dataValidade
    };
    
    for (let key in placeholders) {
      try {
        slide.replaceAllText(key, placeholders[key]);
      } catch (e) {
        Logger.log('Erro replaceAllText: ' + key + ' - ' + e);
      }
    }

    // FOTO 3x4
    if (st.foto) {
      try {
        const fotoBlob = fetchBlobFromReference(st.foto);
        if (fotoBlob) {
          let inserted = false;

          const elements = slide.getPageElements();
          for (let el of elements) {
            if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
              const shape = el.asShape();
              const text = shape.getText().asString();
              if (text.includes('{{FOTO}}')) {
                const left = shape.getLeft();
                const top = shape.getTop();
                const width = shape.getWidth();
                const height = shape.getHeight();
                shape.remove();
                
                const img = slide.insertImage(fotoBlob);
                const scaled = scaleImageToFit3x4(img, width, height);
                
                img.setLeft(left + (width - scaled.w) / 2);
                img.setTop(top + (height - scaled.h) / 2);
                img.setWidth(scaled.w).setHeight(scaled.h);
                
                const border = img.getBorder();
                border.setTransparent(false);
                border.getLineFill().setSolidFill('#000000');
                border.setWeight(3);
                
                inserted = true;
                break;
              }
            }
          }

          if (!inserted) {
            const elements2 = slide.getPageElements();
            for (let el of elements2) {
              if (el.getPageElementType() === SlidesApp.PageElementType.TABLE) {
                const table = el.asTable();
                const rows = table.getNumRows();
                const cols = table.getNumColumns();
                const tableLeft = table.getLeft();
                const tableTop = table.getTop();

                for (let r = 0; r < rows; r++) {
                  for (let c = 0; c < cols; c++) {
                    const cell = table.getCell(r, c);
                    const txt = cell.getText().asString();
                    if (txt.includes('{{FOTO}}')) {
                      cell.getText().clear();
                      
                      const cellWidth = cell.getWidth();
                      const cellHeight = cell.getHeight();
                      const left = tableLeft + cell.getLeft();
                      const top = tableTop + cell.getTop();
                      
                      const img = slide.insertImage(fotoBlob);
                      const scaled = scaleImageToFit3x4(img, cellWidth, cellHeight);
                      
                      img.setLeft(left + (cellWidth - scaled.w) / 2);
                      img.setTop(top + (cellHeight - scaled.h) / 2);
                      img.setWidth(scaled.w).setHeight(scaled.h);
                      
                      const border = img.getBorder();
                      border.setTransparent(false);
                      border.getLineFill().setSolidFill('#000000');
                      border.setWeight(3);
                      
                      inserted = true;
                      break;
                    }
                  }
                  if (inserted) break;
                }
                if (inserted) break;
              }
            }
          }
        }
      } catch (e) {
        Logger.log('Erro ao inserir foto: ' + e);
      }
    }

    pres.saveAndClose();

    const pdfBlob = DriveApp.getFileById(copia.getId())
      .getAs(MimeType.PDF)
      .setName(`Carteirinha_${st.numero}.pdf`);

    const pngBlob = exportSlideAsJPEG(copia.getId(), st.numero);

    const pasta = DriveApp.getFolderById(FOLDER_ID);
    
    // Remove carteirinhas antigas
    const arquivosAntigos = pasta.getFilesByName(`Carteirinha_${st.numero}.pdf`);
    while (arquivosAntigos.hasNext()) {
      arquivosAntigos.next().setTrashed(true);
    }
    
    const pngsAntigos = pasta.getFiles();
    while (pngsAntigos.hasNext()) {
      const file = pngsAntigos.next();
      if (file.getName().startsWith(`Carteirinha_${st.numero}`) && file.getName().endsWith('.png')) {
        file.setTrashed(true);
      }
    }
    
    const arquivoPdf = pasta.createFile(pdfBlob);
    
    if (pngBlob) {
      pasta.createFile(pngBlob);
    }

    DriveApp.getFileById(copia.getId()).setTrashed(true);

    return {
      success: true,
      fileUrl: arquivoPdf.getUrl(),
      fileName: arquivoPdf.getName(),
      studentData: st
    };
  } catch (e) {
    Logger.log('Erro em generateCard: ' + e);
    return { success: false, message: 'Erro ao gerar carteirinha: ' + e.message };
  }
}

function sendEmail(cardNumber) {
  try {
    const studentResult = getStudent(cardNumber);
    if (!studentResult.success) {
      return { success: false, message: studentResult.message };
    }
    
    const st = studentResult.student;
    
    if (!st.email) {
      return { success: false, message: 'Este aluno não possui e-mail cadastrado.' };
    }

    const pasta = DriveApp.getFolderById(FOLDER_ID);
    const arquivos = pasta.getFilesByName(`Carteirinha_${st.numero}.pdf`);
    
    if (!arquivos.hasNext()) {
      return { success: false, message: 'Carteirinha não encontrada. Gere a carteirinha primeiro.' };
    }

    const arquivoPdf = arquivos.next();
    const pdfBlob = arquivoPdf.getBlob();
    
    const attachments = [pdfBlob];
    const arquivosPng = pasta.getFiles();
    while (arquivosPng.hasNext()) {
      const file = arquivosPng.next();
      if (file.getName().startsWith(`Carteirinha_${st.numero}`) && file.getName().endsWith('.png')) {
        attachments.push(file.getBlob());
        break;
      }
    }

    const emailContent = `
      <!DOCTYPE html>
      <html lang="pt-BR">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
      </head>
      <body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f4f4;">
        <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f4f4f4; padding: 20px;">
          <tr>
            <td align="center">
              <table width="600" cellpadding="0" cellspacing="0" style="background-color: #ffffff; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                
                <tr>
                  <td style="background: linear-gradient(90deg, #E31E24 0%, #E31E24 85%, #FFD700 85%, #FFD700 100%); padding: 40px 30px; text-align: center;">
                    <h1 style="margin: 0; color: #ffffff; font-size: 28px; font-weight: 600;">
                      Expresso Itamarati
                    </h1>
                    <p style="margin: 10px 0 0 0; color: #ffffff; font-size: 14px; opacity: 0.9;">
                      Transporte Escolar de Qualidade
                    </p>
                  </td>
                </tr>
                
                <tr>
                  <td style="padding: 40px 30px;">
                    <p style="margin: 0 0 20px 0; color: #333333; font-size: 16px; line-height: 1.6;">
                      Olá <strong>${st.nome}</strong>,
                    </p>
                    
                    <p style="margin: 0 0 20px 0; color: #555555; font-size: 15px; line-height: 1.6;">
                      Sua carteirinha escolar está pronta! 🎉
                    </p>
                    
                    <p style="margin: 0 0 20px 0; color: #555555; font-size: 15px; line-height: 1.6;">
                      Em anexo você encontrará sua carteirinha nos formatos <strong>PDF</strong> e <strong>imagem</strong>.
                    </p>
                    
                    <div style="background-color: #fff5f5; border-left: 4px solid #E31E24; padding: 15px 20px; margin: 25px 0; border-radius: 4px;">
                      <p style="margin: 0; color: #666666; font-size: 14px;">
                        <strong>💡 Dica:</strong> Tenha sempre sua carteirinha em mãos no transporte escolar.
                      </p>
                    </div>
                    
                    <p style="margin: 30px 0 0 0; color: #555555; font-size: 15px;">
                      Qualquer dúvida, estamos à disposição.
                    </p>
                  </td>
                </tr>
                
                <tr>
                  <td style="background-color: #f8f9fa; padding: 30px; text-align: center;">
                    <p style="margin: 0 0 10px 0; color: #333333; font-size: 15px;">Atenciosamente,</p>
                    <p style="margin: 0 0 20px 0; color: #E31E24; font-size: 16px; font-weight: 600;">
                      Equipe Expresso Itamarati
                    </p>
                  </td>
                </tr>
                
              </table>
            </td>
          </tr>
        </table>
      </body>
      </html>
    `;

    GmailApp.sendEmail(
      st.email,
      'Sua Carteirinha Escolar - Expresso Itamarati',
      '',
      {
        htmlBody: emailContent,
        attachments: attachments,
        name: 'Expresso Itamarati',
        from: 'carteirinha@expressoitamarati.com.br'
      }
    );

    return { 
      success: true, 
      message: `E-mail enviado com sucesso para ${st.email}` 
    };
    
  } catch (e) {
    Logger.log('Erro em sendEmail: ' + e);
    return { 
      success: false, 
      message: 'Erro ao enviar e-mail: ' + e.message 
    };
  }
}

/** Exporta o slide como PNG **/
function exportSlideAsJPEG(presentationId, numeroCarteirinha) {
  try {
    const presentation = SlidesApp.openById(presentationId);
    const slides = presentation.getSlides();
    
    if (slides.length === 0) return null;
    
    const slideId = slides[0].getObjectId();
    const url = `https://docs.google.com/presentation/d/${presentationId}/export/png?id=${presentationId}&pageid=${slideId}`;
    
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const blob = response.getBlob();
      blob.setName(`Carteirinha_${numeroCarteirinha}.png`);
      return blob;
    }
    return null;
  } catch (e) {
    Logger.log('Erro em exportSlideAsJPEG: ' + e);
    return null;
  }
}

/** Ajusta imagem para proporção 3x4 **/
function scaleImageToFit3x4(img, maxWidth, maxHeight) {
  const ASPECT_RATIO_3X4 = 0.75;
  
  const originalWidth = img.getWidth();
  const originalHeight = img.getHeight();
  const originalRatio = originalWidth / originalHeight;
  
  let finalWidth, finalHeight;
  
  if (Math.abs(originalRatio - ASPECT_RATIO_3X4) < 0.1) {
    if (maxWidth / maxHeight <= ASPECT_RATIO_3X4) {
      finalWidth = maxWidth;
      finalHeight = finalWidth / ASPECT_RATIO_3X4;
    } else {
      finalHeight = maxHeight;
      finalWidth = finalHeight * ASPECT_RATIO_3X4;
    }
  } else {
    const widthBasedHeight = maxWidth / ASPECT_RATIO_3X4;
    const heightBasedWidth = maxHeight * ASPECT_RATIO_3X4;
    
    if (widthBasedHeight <= maxHeight) {
      finalWidth = maxWidth;
      finalHeight = widthBasedHeight;
    } else {
      finalWidth = heightBasedWidth;
      finalHeight = maxHeight;
    }
  }
  
  if (finalWidth > maxWidth) {
    finalWidth = maxWidth;
    finalHeight = finalWidth / ASPECT_RATIO_3X4;
  }
  
  if (finalHeight > maxHeight) {
    finalHeight = maxHeight;
    finalWidth = finalHeight * ASPECT_RATIO_3X4;
  }
  
  return { w: finalWidth, h: finalHeight };
}

/** Converte link/ID/base64 em Blob **/
function fetchBlobFromReference(ref) {
  if (!ref) return null;
  ref = String(ref).trim();
  try {
    if (ref.startsWith('data:')) {
      const parts = ref.split(',');
      if (parts.length === 2) {
        const meta = parts[0];
        const b64 = parts[1];
        const contentType = meta.split(';')[0].split(':')[1] || 'image/png';
        const bytes = Utilities.base64Decode(b64);
        return Utilities.newBlob(bytes, contentType, 'foto_base64');
      }
    }

    const driveIdMatch =
      ref.match(/\/d\/([a-zA-Z0-9_-]{10,})/) ||
      ref.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
    if (driveIdMatch) {
      const id = driveIdMatch[1];
      try {
        return DriveApp.getFileById(id).getBlob();
      } catch (e) {
        const altUrl = `https://drive.google.com/uc?export=view&id=${id}`;
        const resp = UrlFetchApp.fetch(altUrl, { muteHttpExceptions: true });
        if (resp.getResponseCode() === 200) return resp.getBlob();
        return null;
      }
    }

    if (/^[a-zA-Z0-9_-]{20,}$/.test(ref)) {
      try {
        return DriveApp.getFileById(ref).getBlob();
      } catch (e) {
        const altUrl = `https://drive.google.com/uc?export=view&id=${ref}`;
        const resp = UrlFetchApp.fetch(altUrl, { muteHttpExceptions: true });
        if (resp.getResponseCode() === 200) return resp.getBlob();
        return null;
      }
    }

    if (ref.startsWith('http')) {
      const resp = UrlFetchApp.fetch(ref, { muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) return resp.getBlob();
    }

    return null;
  } catch (e) {
    Logger.log('Erro em fetchBlobFromReference: ' + e);
    return null;
  }
}