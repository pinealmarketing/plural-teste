// Teste de polling de uma pasta do OneDrive via Microsoft Graph:
// lista, baixa, EXTRAI O CONTEÚDO (.xlsx/.pdf) e move o arquivo para a subpasta "Processados".
// Estado em memória (reiniciar = recomeça do zero).
import { createHash } from 'node:crypto';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { dirname } from 'node:path';
import ExcelJS from 'exceljs';
import { extractText, getDocumentProxy } from 'unpdf';

const {
  CLIENT_ID,
  FOLDER_PATH = 'DME-Entrada',   // pasta na raiz do OneDrive
  TENANT_ID = 'common',
  POLL_SECONDS = '30',           // de quanto em quanto tempo olha a pasta
  STABLE_SECONDS = '20',         // só processa se o arquivo estiver parado há tanto tempo
  EXTENSIONS = '.xlsx,.pdf',
  TOKEN_FILE = '/data/token.json', // onde guardar o login (use um Volume do Railway montado em /data)
  PROCESSED_FOLDER = 'Processados', // subpasta dentro de FOLDER_PATH (criada se não existir)
  MOVE_PROCESSED = 'true',          // 'false' = só extrai, não mexe nos arquivos
  PREVIEW_CHARS = '400',            // quanto do conteúdo extraído aparece no log
} = process.env;

const GRAPH = process.env.GRAPH_BASE || 'https://graph.microsoft.com/v1.0';
const LOGIN = process.env.LOGIN_BASE || `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0`;
const SCOPE = 'Files.ReadWrite offline_access'; // ReadWrite é necessário para mover
const exts = EXTENSIONS.split(',').map((e) => e.trim().toLowerCase());
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const log = (tag, msg) => console.log(`${new Date().toISOString().slice(11, 19)} [${tag}] ${msg}`);

if (!CLIENT_ID) {
  console.error('Falta a variável CLIENT_ID (ID do app registrado no Azure).');
  process.exit(1);
}

// ---------- login (código de dispositivo) ----------
let token = null, tokenExp = 0, refreshToken = null;

async function loadSavedLogin() {
  try {
    refreshToken = JSON.parse(await readFile(TOKEN_FILE, 'utf8')).refreshToken || null;
    if (refreshToken) log('SESSÃO', `login salvo encontrado em ${TOKEN_FILE}; tentando reutilizar`);
  } catch {
    log('SESSÃO', `nenhum login salvo em ${TOKEN_FILE}; será pedido um login novo`);
  }
}

let warnedSave = false;
async function saveLogin() {
  try {
    await mkdir(dirname(TOKEN_FILE), { recursive: true });
    await writeFile(TOKEN_FILE, JSON.stringify({ refreshToken, savedAt: new Date().toISOString() }), { mode: 0o600 });
  } catch (e) {
    if (!warnedSave) {
      warnedSave = true;
      log('AVISO', `não consegui salvar o login em ${TOKEN_FILE} (${e.message}). Sem um Volume, o login se perde a cada reinício.`);
    }
  }
}

const form = async (url, params) =>
  (await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(params),
    signal: AbortSignal.timeout(30_000),
  })).json();

async function deviceLogin() {
  const dc = await form(`${LOGIN}/devicecode`, { client_id: CLIENT_ID, scope: SCOPE });
  if (!dc.device_code) throw new Error('Falha ao iniciar login: ' + JSON.stringify(dc));
  console.log('\n=============== LOGIN NECESSÁRIO ===============');
  console.log(dc.message);
  console.log('================================================\n');
  const deadline = Date.now() + dc.expires_in * 1000;
  let wait = (dc.interval || 5) * 1000;
  let lastBeat = Date.now();
  while (Date.now() < deadline) {
    await sleep(wait);
    const j = await form(`${LOGIN}/token`, {
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      client_id: CLIENT_ID,
      device_code: dc.device_code,
    });
    if (j.access_token) { log('LOGIN OK', 'autenticado com sucesso'); return j; }
    if (Date.now() - lastBeat >= 30_000) {
      lastBeat = Date.now();
      const min = Math.max(0, Math.round((deadline - Date.now()) / 60_000));
      log('AGUARDANDO LOGIN', `código ${dc.user_code ?? '(veja acima)'} · expira em ~${min} min · ainda não foi digitado`);
    }
    if (j.error === 'slow_down') wait += 5000;
    else if (j.error !== 'authorization_pending') throw new Error('Login falhou: ' + JSON.stringify(j));
  }
  throw new Error('Código de login expirou. Reinicie o serviço.');
}

async function getToken() {
  if (token && Date.now() < tokenExp - 60_000) return token;
  let j;
  if (refreshToken) {
    j = await form(`${LOGIN}/token`, {
      grant_type: 'refresh_token', client_id: CLIENT_ID, refresh_token: refreshToken, scope: SCOPE,
    });
  }
  if (!j?.access_token) j = await deviceLogin();
  token = j.access_token;
  if (j.refresh_token && j.refresh_token !== refreshToken) {
    refreshToken = j.refresh_token; // a Microsoft troca o refresh token a cada uso; precisa salvar o novo
    await saveLogin();
  }
  tokenExp = Date.now() + j.expires_in * 1000;
  return token;
}

// ---------- Graph ----------
async function api(url, { method = 'GET', body, allow404 = false } = {}, tries = 0) {
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${await getToken()}`,
      ...(body && { 'Content-Type': 'application/json' }),
    },
    body: body && JSON.stringify(body),
    signal: AbortSignal.timeout(30_000),
  });
  if (res.status === 401 && tries < 1) { token = null; return api(url, { method, body, allow404 }, tries + 1); }
  if ((res.status === 429 || res.status >= 500) && tries < 3) {
    await sleep((Number(res.headers.get('retry-after')) || 2 ** tries * 2) * 1000);
    return api(url, { method, body, allow404 }, tries + 1);
  }
  if (allow404 && res.status === 404) return null;
  if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
  return res.json();
}

const folderUrl = (path) => `${GRAPH}/me/drive/root:/${encodeURI(path.replace(/^\/+|\/+$/g, ''))}`;

async function listFolder() {
  const items = [];
  let url = `${folderUrl(FOLDER_PATH)}:/children?$top=200`;
  while (url) {
    const j = await api(url);
    items.push(...j.value);
    url = j['@odata.nextLink'];
  }
  return items;
}

async function download(item) {
  const res = await fetch(item['@microsoft.graph.downloadUrl']); // URL pré-autenticada, sem header
  if (!res.ok) throw new Error(`Download ${res.status} de ${item.name}`);
  return Buffer.from(await res.arrayBuffer());
}

// ---------- extração de conteúdo ----------
async function extractXlsx(buf) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const parts = [];
  let rows = 0;
  wb.eachSheet((ws) => {
    const lines = [];
    ws.eachRow({ includeEmpty: false }, (row) => {
      const cells = [];
      row.eachCell({ includeEmpty: true }, (c, n) => { cells[n - 1] = c.text; });
      lines.push(Array.from(cells, (v) => v ?? '').join('\t'));
    });
    rows += lines.length;
    parts.push(`## Planilha: ${ws.name}\n${lines.join('\n')}`);
  });
  return { resumo: `${wb.worksheets.length} planilha(s), ${rows} linha(s)`, texto: parts.join('\n\n') };
}

async function extractPdf(buf) {
  const pdf = await getDocumentProxy(new Uint8Array(buf));
  const { totalPages, text } = await extractText(pdf, { mergePages: true });
  return { resumo: `${totalPages} página(s), ${text.length} caractere(s)`, texto: text };
}

const extract = (ext, buf) => (ext === '.pdf' ? extractPdf(buf) : extractXlsx(buf));

// ---------- mover para "Processados" ----------
let processedId = null; // id da subpasta, descoberto/criado uma vez

async function getProcessedFolderId(parentId) {
  if (processedId) return processedId;
  const path = `${FOLDER_PATH.replace(/\/+$/, '')}/${PROCESSED_FOLDER}`;
  let f = await api(folderUrl(path), { allow404: true });
  if (!f) {
    f = await api(`${GRAPH}/me/drive/items/${parentId}/children`, {
      method: 'POST',
      body: { name: PROCESSED_FOLDER, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' },
    });
    log('PASTA', `criada "${path}"`);
  }
  return (processedId = f.id);
}

async function moveToProcessed(item) {
  const destId = await getProcessedFolderId(item.parentReference.id);
  await api(`${GRAPH}/me/drive/items/${item.id}`, {
    method: 'PATCH',
    body: { parentReference: { id: destId }, '@microsoft.graph.conflictBehavior': 'rename' },
  });
}

// ---------- polling ----------
const seen = new Map();      // id -> { name, etag, sha, version, missing }
const ignored = new Set();   // para não repetir o aviso de "ignorado" a cada ciclo

async function cycle() {
  log('LENDO', `pasta "${FOLDER_PATH}" no OneDrive...`);
  const items = await listFolder();
  const ids = new Set();
  let count = 0;

  for (const it of items) {
    if (!it.file) continue; // subpastas não são lidas
    const ext = it.name.includes('.') ? it.name.slice(it.name.lastIndexOf('.')).toLowerCase() : '';
    if (it.name.startsWith('~$') || !exts.includes(ext)) {
      if (!ignored.has(it.id)) { ignored.add(it.id); log('IGNORADO', it.name); }
      continue;
    }
    count++;
    ids.add(it.id);

    const prev = seen.get(it.id);
    if (prev && prev.etag === it.eTag && !prev.missing && prev.processed) continue; // nada mudou

    const idade = (Date.now() - Date.parse(it.lastModifiedDateTime)) / 1000;
    if (idade < Number(STABLE_SECONDS)) {
      log('AGUARDANDO', `${it.name} mexido há ${Math.round(idade)}s; esperando estabilizar`);
      continue;
    }

    try {
      const buf = await download(it);
      const sha = createHash('sha256').update(buf).digest('hex');
      const emAndamento = prev && prev.sha === sha && !prev.processed; // extração/movimento falhou antes

      if (emAndamento) {
        log('RETENTATIVA', `${it.name}: reprocessando`);
      } else if (!prev) {
        const dup = [...seen.entries()].find(([id, v]) => id !== it.id && v.sha === sha);
        seen.set(it.id, { name: it.name, etag: it.eTag, sha, version: 1 });
        if (dup) log('DUPLICADO', `${it.name} tem o mesmo conteúdo de ${dup[1].name}`);
        else log('NOVO', `${it.name} (${buf.length} bytes)`);
      } else if (sha === prev.sha) {
        const voltou = prev.missing;
        Object.assign(prev, { name: it.name, etag: it.eTag, missing: false });
        if (voltou) log('VOLTOU', `${it.name} reapareceu na pasta com o mesmo conteúdo`);
        else log('SÓ METADADOS', `${it.name}: eTag mudou mas o conteúdo é idêntico`);
      } else {
        Object.assign(prev, { name: it.name, etag: it.eTag, sha, version: prev.version + 1, missing: false, processed: false });
        log('NOVA VERSÃO', `${it.name} agora na versão ${prev.version} (${buf.length} bytes)`);
      }

      const st = seen.get(it.id);
      if (!st.processed) {
        const { resumo, texto } = await extract(ext, buf);
        log('CONTEÚDO', `${it.name}: ${resumo}`);
        console.log(texto.slice(0, Number(PREVIEW_CHARS)).replace(/^/gm, '    | ') + (texto.length > Number(PREVIEW_CHARS) ? '\n    | …' : ''));
        if (MOVE_PROCESSED === 'true') {
          await moveToProcessed(it);
          log('MOVIDO', `${it.name} -> ${PROCESSED_FOLDER}/`);
        }
        st.processed = true;
      }
    } catch (e) {
      log('ERRO', `${it.name}: ${e.message}`); // não atualiza o estado -> tenta de novo no próximo ciclo
    }
  }

  for (const [id, v] of seen) {
    if (!ids.has(id) && !v.missing && !(v.processed && MOVE_PROCESSED === 'true')) { v.missing = true; log('SUMIU', `${v.name} não está mais na pasta`); }
  }
  log('CICLO', `${count} arquivo(s) elegíveis na pasta, ${seen.size} conhecido(s)`);
}

await loadSavedLogin();
log('INÍCIO', `pasta "${FOLDER_PATH}" · a cada ${POLL_SECONDS}s · estabilização ${STABLE_SECONDS}s · ${exts.join(', ')}`);
while (true) {
  try { await cycle(); } catch (e) { log('ERRO CICLO', e.message); }
  await sleep(Number(POLL_SECONDS) * 1000);
}
