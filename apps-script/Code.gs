// ===================== WASHLAB → GOOGLE SHEETS =====================
// Regista os cards do dashboard na folha "BD Cards" (1 linha por card, atualizada pelo ID).
// Publicar: Implementar > Nova implementação > Aplicação Web
//   Executar como: Eu | Quem tem acesso: Qualquer pessoa

const SPREADSHEET_ID = '1hMY74mhqyAPhtC3QJA5vitB5lCPnoILIp6Aw7KV-cVE';
const SHEET_NAME = 'BD Cards';

// [Cabeçalho na sheet, chave enviada pelo dashboard]
const COLUMNS = [
  ['ID Card', 'id'],
  ['Criado Em', 'createdAt'],
  ['Atualizado Em', 'updatedAt'],
  ['Última Ação', 'action'],
  ['Origem', 'origem'],
  ['Zona', 'zone'],
  ['Secção', 'section'],
  ['Estado', 'status'],
  ['Comercial', 'comercial'],
  ['Cliente/Revendedor', 'clienteNome'],
  ['País', 'pais'],
  ['Morada', 'morada'],
  ['VAT', 'vat'],
  ['Pessoa de Contacto', 'pessoaContacto'],
  ['Procedimento', 'procedimento'],
  ['Idioma', 'idioma'],
  ['Utensílio(s)', 'utensilios'],
  ['Qtd Utensílios', 'qtdUtensilios'],
  ['Trolley(s)', 'trolleys'],
  ['Qtd Trolleys', 'qtdTrolleys'],
  ['Tipo de Sujidade', 'sujidade'],
  ['Limpeza / Secagem', 'limpezaSecagem'],
  ['Outro', 'outro'],
  ['Observações', 'observacoes'],
  ['Temperatura Máxima', 'tempMax'],
  ['Data de Registo', 'dataRegisto'],
  ['Previsão de Conclusão', 'previsaoConclusao'],
  ['Nº Imagens', 'numImagens'],
  ['Detalhe Utensílios', 'utensiliosDetalhe'],
  ['Eliminado', 'eliminado']
];

// Correr uma vez à mão (ou é chamado automaticamente) para criar/repor os cabeçalhos
function setupHeaders() {
  const sheet = getSheet_();
  const headers = COLUMNS.map(c => c[0]);
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#00b8d4')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.getRange('A:A').setNumberFormat('@');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    if (sheet.getLastRow() === 0) setupHeaders();

    payload.updatedAt = new Date();
    payload.eliminado = payload.action === 'delete' ? 'Sim' : 'Não';

    const row = COLUMNS.map(([, key]) => toCell_(payload[key]));
    const rowIndex = findRowById_(sheet, payload.id);

    if (rowIndex) {
      // Mantém a data de criação original
      row[1] = sheet.getRange(rowIndex, 2).getValue() || row[1];
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return json_({ ok: true, row: rowIndex || sheet.getLastRow() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Teste rápido no browser: abre o URL da aplicação web
function doGet() {
  return json_({ ok: true, sheet: SHEET_NAME });
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function findRowById_(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return null;
  const match = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(id))
    .matchEntireCell(true)
    .findNext();
  return match ? match.getRow() : null;
}

function toCell_(value) {
  if (value === undefined || value === null) return '';
  if (Array.isArray(value)) {
    // Detalhe de utensílios vindo do PDF → uma linha por utensílio
    return value.map(u => [
      u.id, u.designation, u.type, (u.quantity || 0) + ' un.', u.dimensions,
      u.material, u.dirtType, u.washingObjective,
      'Lav/dia: ' + (u.washingNeeds || ''), 'Secagem: ' + (u.dryingNeeds || '')
    ].filter(Boolean).join(' | ')).join('\n').slice(0, 49000);
  }
  if (typeof value === 'object') return JSON.stringify(value).slice(0, 49000);
  return value;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
