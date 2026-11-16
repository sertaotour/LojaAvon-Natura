/**
 * JULIETA BARBOSA AVON & NATURA BELEZA
 * Aplicação completa Google Apps Script + Google Sheets, sem Cloudflare Worker.
 * Configure ADMIN_PASSWORD em Configurações do projeto > Propriedades do script.
 * Não publique chaves da InfinitePay no HTML ou em repositórios públicos.
 */
const SHEETS = {
  PRODUCTS: 'Produtos',
  ORDERS: 'Pedidos'
};
const PRODUCT_HEADERS = ['id','nome','marca','categoria','preco','estoque','imagem','descricao','ativo','criadoEm','atualizadoEm'];
const ORDER_HEADERS = ['id','criadoEm','cliente','itensJson','total','status','observacoes'];

function doGet(e) {
  // O parâmetro ?pagina=admin é lido no servidor para abrir a tela correta.
  const isAdmin = !!(e && e.parameter && e.parameter.pagina === 'admin');
  const template = HtmlService.createTemplateFromFile('index');
  template.isAdmin = isAdmin;
  return template.evaluate()
    .setTitle(isAdmin ? 'Administração | Julieta Barbosa' : 'Julieta Barbosa | Avon & Natura Beleza')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Ponto de entrada chamado pelo navegador via google.script.run.
 * O site é servido pelo próprio Apps Script, portanto não precisa de Worker/CORS.
 */
function apiCall(action, data) {
  try {
    const payload = Object.assign({}, data || {}, { action: String(action || '') });
    return route_(payload);
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return { ok: false, message: err.message || 'Erro interno.' };
  }
}

function route_(b) {
  switch (b.action) {
    case 'listProducts': return { ok: true, products: listProducts_(false) };
    case 'adminLogin': return adminLogin_(b.password);
    case 'adminListProducts': requireAdmin_(b.token); return { ok: true, products: listProducts_(true) };
    case 'saveProduct': requireAdmin_(b.token); return saveProduct_(b.product || {});
    case 'deleteProduct': requireAdmin_(b.token); return deleteProduct_(b.id);
    case 'createOrder': return createOrder_(b);
    case 'createCheckout': return { ok: false, code: 'INFINITEPAY_NOT_CONFIGURED', message: 'Checkout InfinitePay ainda não configurado. Cadastre as credenciais e conclua a ativação da integração.' };
    default: return { ok: false, message: 'Ação desconhecida.' };
  }
}

/** Execute manualmente uma vez no editor para criar as abas e cabeçalhos. */
function setupStore() {
  const ss = spreadsheet_();
  ensureSheet_(ss, SHEETS.PRODUCTS, PRODUCT_HEADERS);
  ensureSheet_(ss, SHEETS.ORDERS, ORDER_HEADERS);
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  Logger.log('Planilha preparada: ' + ss.getUrl());
}

function spreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) throw new Error('Crie uma planilha, abra Extensões > Apps Script e execute setupStore().');
  return active;
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(headers);
  else {
    const current = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    if (current.join('|') !== headers.join('|')) throw new Error('Cabeçalho inesperado na aba ' + name + '. Não altere a primeira linha.');
  }
  sh.setFrozenRows(1);
  return sh;
}

function productsSheet_() { return ensureSheet_(spreadsheet_(), SHEETS.PRODUCTS, PRODUCT_HEADERS); }
function ordersSheet_() { return ensureSheet_(spreadsheet_(), SHEETS.ORDERS, ORDER_HEADERS); }

function listProducts_(includeInactive) {
  const sh = productsSheet_();
  if (sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, PRODUCT_HEADERS.length).getValues();
  return rows.filter(r => includeInactive || String(r[8]).toLowerCase() === 'true')
    .map(r => ({ id: String(r[0]), name: String(r[1]), brand: String(r[2]), category: String(r[3]), price: Number(r[4]), stock: Number(r[5]), image: String(r[6] || ''), desc: String(r[7] || ''), active: String(r[8]).toLowerCase() === 'true' }));
}

function adminLogin_(password) {
  const saved = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!saved) return { ok: false, message: 'Configure ADMIN_PASSWORD nas propriedades do Apps Script antes de acessar o painel.' };
  if (!constantTimeEquals_(String(password || ''), saved)) return { ok: false, message: 'Senha incorreta.' };
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('admin_' + token, 'ok', 21600);
  return { ok: true, token: token };
}

function requireAdmin_(token) {
  if (!token || CacheService.getScriptCache().get('admin_' + String(token)) !== 'ok') throw new Error('Sessão administrativa expirada. Entre novamente.');
}

function saveProduct_(p) {
  const name = String(p.name || '').trim();
  const brand = String(p.brand || '');
  const category = String(p.category || '').trim();
  const price = Number(p.price);
  const stock = Number(p.stock);
  const image = String(p.image || '').trim();
  if (!name || name.length > 120) throw new Error('Informe um nome de produto válido (até 120 caracteres).');
  if (['Avon', 'Natura'].indexOf(brand) < 0) throw new Error('Marca inválida.');
  if (!category || !Number.isFinite(price) || price < 0 || !Number.isInteger(stock) || stock < 0) throw new Error('Confira categoria, preço e estoque.');
  if (image && !/^https:\/\//i.test(image)) throw new Error('A imagem deve usar um endereço HTTPS.');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = productsSheet_();
    const now = new Date();
    const id = String(p.id || Utilities.getUuid());
    let row = 0;
    if (p.id && sh.getLastRow() >= 2) {
      const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().map(String);
      const index = ids.indexOf(id); if (index >= 0) row = index + 2;
    }
    const created = row ? sh.getRange(row, 10).getValue() : now;
    const values = [id, name, brand, category, price, stock, image, String(p.desc || '').slice(0, 500), true, created, now];
    if (row) sh.getRange(row, 1, 1, values.length).setValues([values]); else sh.appendRow(values);
    return { ok: true, productId: id };
  } finally { lock.releaseLock(); }
}

function deleteProduct_(id) {
  const sh = productsSheet_();
  if (sh.getLastRow() < 2) return { ok: true };
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().map(String);
  const index = ids.indexOf(String(id));
  if (index >= 0) sh.getRange(index + 2, 9).setValue(false); // exclusão lógica
  return { ok: true };
}

function createOrder_(b) {
  const customer = String((b.customer && b.customer.name) || '').trim();
  const items = Array.isArray(b.items) ? b.items : [];
  if (!customer || !items.length) throw new Error('Informe seu nome e pelo menos um produto.');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const catalog = listProducts_(false);
    let total = 0;
    const verified = items.map(item => {
      const p = catalog.find(x => x.id === String(item.productId));
      const quantity = Number(item.quantity);
      if (!p || !Number.isInteger(quantity) || quantity < 1 || quantity > p.stock) throw new Error('Produto indisponível ou quantidade inválida: ' + item.productId);
      total += p.price * quantity;
      return { productId: p.id, name: p.name, brand: p.brand, quantity: quantity, unitPrice: p.price, subtotal: p.price * quantity };
    });
    const id = 'JB-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    ordersSheet_().appendRow([id, new Date(), customer, JSON.stringify(verified), total, 'AGUARDANDO PAGAMENTO', String(b.note || '').slice(0, 500)]);
    return { ok: true, orderId: id, total: total, status: 'AGUARDANDO PAGAMENTO' };
  } finally { lock.releaseLock(); }
}

function constantTimeEquals_(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0; for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
