/**
 * app.js — Google Sheets API Client
 * -----------------------------------
 * Conecta directamente con Google Sheets usando la cuenta de servicio
 * vía autenticación JWT (sin backend). Usa Web Crypto API nativa del navegador.
 *
 * Estructura del Sheet:
 *   - Múltiples hojas (una por semestre)
 *   - Fila 1, columnas D:AD → nombres de personas (encabezados)
 *   - Columna A, filas 2 en adelante → nombres de materias
 *   - Celda [fila_materia, col_persona] → valor del cupo (número)
 */

// ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────

const SPREADSHEET_ID = '1Xo1ElBJOOR07Gxzv-RbQnpPK6k9jKlg0_9YcZzdI6Lg';

const SERVICE_ACCOUNT = {
  type: "service_account",
  project_id: "registro-de-materias-por-cupo",
  private_key_id: "5559682c2088952dc0718ca399f824327f1b40fe",
  private_key: `-----BEGIN PRIVATE KEY-----
MIIEvwIBADANBgkqhkiG9w0BAQEFAASCBKkwggSlAgEAAoIBAQDh3V8PX9z9wO2i
Td1ss3NnM/PeU0iIa2171tQxAuRioLyP0GQd8AXvwZum0QZNkvj3Iykrr2j8Ol52
yD3cNP1oMinnz0ZgMNbEi35hJA1N+WXX73MlpB8EVxc9/QqMe5zvBUorE1RYtbLy
3SsHjWF9TdJCxRlvMG5gfU+W2uTfGj0zP4qaSD4yxgaUHdFjmxzAmdWGfZFUY1Qs
L29lQThbLIFSmXYeht7Gr1G7btEdCUs8eVQgoee/8nY4Zcj4P+72rGcF4tGujKKO
DHgUxpzW690UO3ukZ03LYOUSWfG0D1EP9EcudFpAzPSsbW5qu8iDSUetmBfFLvui
oj02ya+bAgMBAAECggEACaGnAWt/wP0OenKs6TW9y2jWaA9PyIF26buqJheO0FCK
ZJeTrV3O4v/JOL6GcPMTgCEVCwfah5GgSvBpsuTk6Xc6J5MZ19WSqE71KgSfjJ5V
+XSnNF5gnuQX2aIwu6I0eaeAF4DQ1/eSP0kEd2NQBPKMmiGi3rBHWgCopDDcEadi
7uBDuPeE751ywr6tlwoYyLnVKzwQg2IrhuPEL1xsCs7fSC/J0uFQUMMz+0DaRiLQ
vKJzYAa3GVbvDJtumfV1l6AfkEzZB7Npy11r65dk9KPkI/gsKPu+bd1ZsJupsdYX
oFgC25OjXRRsJ/mBZZ4umFXckpyUe7cpC0EPf2CCAQKBgQD5wpuenz9HDzayoRT0
mUtdLqT6z/7mYQm4WEygeByv6w+dF4SOpIFhxEp84uaBKPvoyeq/5Gwa1q/E+xae
cRaV913WgTVoZrxXJzJUFoH653PRF6ze+PCXJrRphfE+knhKAmK86txiLNV5E/Cq
tl4Y46SUtGOIaizQn9S04aHyJwKBgQDnge9osTYao1dqZ6qKrcKS2oK6ASFBeXSU
KHhEYVymENKvTzKFhbQYqp8F15nh4Mbkc7sJQp01dK1nofeqhc/784vzp7FxK1Cl
aCUyzRwIZ9GfyRvKfNDnmFurH/L+G78y6+RM3d5q0kbCRt6X0L2LXMrYTdzyv49p
Uvn7lxvjbQKBgQDvf0my5YnMHi6ZRBXQJ185T40sZV9Mkyi6+REhn4wCtSkXvoGC
NwVKNuwmrX3TxPUq2NSehe+UHOIXxZ5++Hpr+/SjyOrp3fokqJV+RMcHTgKlMkq2
1Yf/qUG/Ho2jLtjiPz9nYN9L1SovHIvfZ1j8DO65GfGH0ih/NYTGnsaoaQKBgQCV
pH2WFIY+bbrBgsTP40VUG35IsRZH9jQO2KH0wWJbzaABxZWIjUY+c3tbEWPch6jI
Xq5VbAOmXAcCZ8VpKhmoaGLcWlbuKet1H357+ezW2hS7zgjyt/9o1Cjc0kgFTPYn
+iaWMQvlzIoEZj7Xrwv2G0La0mmxV3VhxUrk/2X9eQKBgQCqDd76MZPvn3Cat3WY
yRrNnWH4Hbp2y+T/N1m0E6JKQOEwiB+IfoQ04/C0WYxb++EQEC2w711JYycbvcG/
+IFmkrwdldxRpRjEOUhl1JmOY2VLKY3Uh1HYXbq0VI3hKdeU/+m5Ymbd+MjduUTq
pYsdk66bcNy5aDqeOiaQUymP4g==
-----END PRIVATE KEY-----`,
  client_email: "registrocupos@registro-de-materias-por-cupo.iam.gserviceaccount.com",
  token_uri: "https://oauth2.googleapis.com/token"
};

const SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets';
const SHEETS_API   = 'https://sheets.googleapis.com/v4/spreadsheets';

// ─── CACHÉ DE TOKEN ───────────────────────────────────────────────────────────

let _tokenCache = { token: null, expiry: 0 };

// ─── UTILIDADES JWT (Web Crypto API) ─────────────────────────────────────────

/**
 * Convierte Base64URL a Base64 estándar
 */
function b64urlToB64(str) {
  return str.replace(/-/g, '+').replace(/_/g, '/') + '=='.slice(0, (4 - str.length % 4) % 4);
}

/**
 * Codifica un objeto a Base64URL
 */
function toB64url(obj) {
  const json = typeof obj === 'string' ? obj : JSON.stringify(obj);
  return btoa(unescape(encodeURIComponent(json)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}

/**
 * Importa la clave privada PEM como CryptoKey para firmar con RS256
 */
async function importPrivateKey(pem) {
  const pemContent = pem
    .replace('-----BEGIN PRIVATE KEY-----', '')
    .replace('-----END PRIVATE KEY-----', '')
    .replace(/\s+/g, '');
  const der = Uint8Array.from(atob(pemContent), c => c.charCodeAt(0));
  return crypto.subtle.importKey(
    'pkcs8',
    der.buffer,
    { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' },
    false,
    ['sign']
  );
}

/**
 * Crea y firma un JWT para la cuenta de servicio
 */
async function createJWT() {
  const now     = Math.floor(Date.now() / 1000);
  const header  = { alg: 'RS256', typ: 'JWT' };
  const payload = {
    iss  : SERVICE_ACCOUNT.client_email,
    scope: SHEETS_SCOPE,
    aud  : SERVICE_ACCOUNT.token_uri,
    iat  : now,
    exp  : now + 3600
  };

  const signingInput = `${toB64url(header)}.${toB64url(payload)}`;
  const key          = await importPrivateKey(SERVICE_ACCOUNT.private_key);
  const encoder      = new TextEncoder();
  const signature    = await crypto.subtle.sign(
    'RSASSA-PKCS1-v1_5',
    key,
    encoder.encode(signingInput)
  );

  const sigB64url = btoa(String.fromCharCode(...new Uint8Array(signature)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');

  return `${signingInput}.${sigB64url}`;
}

// ─── AUTENTICACIÓN ────────────────────────────────────────────────────────────

/**
 * Obtiene (o reutiliza si es válido) un access token de Google OAuth2
 * @returns {Promise<string>} Access token
 */
async function getAccessToken() {
  const now = Math.floor(Date.now() / 1000);
  if (_tokenCache.token && now < _tokenCache.expiry - 60) {
    return _tokenCache.token;
  }

  const jwt = await createJWT();
  const resp = await fetch(SERVICE_ACCOUNT.token_uri, {
    method : 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body   : `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`
  });

  if (!resp.ok) {
    const err = await resp.text();
    throw new Error(`Error al obtener token: ${resp.status} — ${err}`);
  }

  const data = await resp.json();
  _tokenCache = {
    token : data.access_token,
    expiry: now + (data.expires_in || 3600)
  };
  return _tokenCache.token;
}

// ─── HELPERS DE SHEETS API ────────────────────────────────────────────────────

/**
 * GET genérico a la Sheets API
 */
async function sheetsGet(path) {
  const token = await getAccessToken();
  const resp  = await fetch(`${SHEETS_API}/${SPREADSHEET_ID}${path}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!resp.ok) throw new Error(`Sheets API error ${resp.status}: ${await resp.text()}`);
  return resp.json();
}

/**
 * BATCH UPDATE genérico a la Sheets API (para escrituras)
 */
async function sheetsBatchUpdate(body) {
  const token = await getAccessToken();
  const resp  = await fetch(`${SHEETS_API}/${SPREADSHEET_ID}/values:batchUpdate`, {
    method : 'POST',
    headers: {
      Authorization : `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body)
  });
  if (!resp.ok) throw new Error(`Sheets API write error ${resp.status}: ${await resp.text()}`);
  return resp.json();
}

/**
 * Lectura de un rango específico (devuelve array 2D de valores)
 */
async function readRange(range) {
  const data = await sheetsGet(`/values/${encodeURIComponent(range)}`);
  return data.values || [];
}

// ─── FUNCIONES PRINCIPALES (equivalentes al Apps Script) ─────────────────────

/**
 * Obtiene los nombres de todas las hojas del spreadsheet
 * Equivale a: obtenerHojas()
 * @returns {Promise<string[]>}
 */
async function obtenerHojas() {
  const data   = await sheetsGet('?fields=sheets.properties.title');
  const sheets = data.sheets || [];
  return sheets.map(s => s.properties.title);
}

/**
 * Obtiene los nombres de personas desde el rango D1:AW1 de la primera hoja.
 * Equivale a: obtenerNombres()
 * @returns {Promise<string[]>}
 */
async function obtenerNombres() {
  // Usamos la primera hoja disponible para leer los enc cabezados
  const hojas = await obtenerHojas();
  if (!hojas.length) return [];
  const rango  = `'${hojas[0]}'!D1:AW1`;
  const values = await readRange(rango);
  if (!values.length) return [];
  return values[0].filter(v => v && String(v).trim() !== '');
}

/**
 * Obtiene la lista de materias (Columna A, filas 2+) de una hoja dada.
 * Equivale a: obtenerMaterias(nombreHoja)
 * @param {string} nombreHoja
 * @returns {Promise<string[]>}
 */
async function obtenerMaterias(nombreHoja) {
  const rango  = `'${nombreHoja}'!A2:A`;
  const values = await readRange(rango);
  return values.map(row => row[0]).filter(v => v && String(v).trim() !== '');
}

/**
 * Lee el valor actual de un cupo (celda intersección materia/persona).
 * @param {string} nombreHoja
 * @param {string} nombreMateria
 * @param {string} nombrePersona
 * @returns {Promise<{row: number, col: number, value: number}>}
 */
async function _encontrarCelda(nombreHoja, nombreMateria, nombrePersona) {
  // Leer materias (col A desde fila 2)
  const materiaValues = await readRange(`'${nombreHoja}'!A2:A`);
  let rowIndex = -1;
  for (let i = 0; i < materiaValues.length; i++) {
    if (materiaValues[i][0] === nombreMateria) {
      rowIndex = i + 2; // fila real en sheet (base 1, empieza en fila 2)
      break;
    }
  }
  if (rowIndex === -1) throw new Error(`Materia "${nombreMateria}" no encontrada.`);

  // Leer nombres (D1:AW1) — columnas de personas
  const nombresValues = await readRange(`'${nombreHoja}'!D1:AW1`);
  const nombres = nombresValues.length ? nombresValues[0] : [];
  let colIndex = -1;
  for (let j = 0; j < nombres.length; j++) {
    if (nombres[j] === nombrePersona) {
      colIndex = j + 4; // columna real: D=4, E=5, etc.
      break;
    }
  }
  if (colIndex === -1) throw new Error(`Persona "${nombrePersona}" no encontrada en encabezados (D1:AW1).`);

  return { row: rowIndex, col: colIndex };
}

/**
 * Convierte número de columna (1-based) a letra(s): 1→A, 4→D, 30→AD, etc.
 */
function colNumToLetter(n) {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Registra un cupo (+1) para la persona en la materia indicada.
 * Equivale a: agregarCupoAEstudiante(nombreHoja, nombreMateria, nombrePersona)
 * @param {string} nombreHoja
 * @param {string} nombreMateria
 * @param {string} nombrePersona
 * @returns {Promise<string>} Mensaje de resultado
 */
async function agregarCupoAEstudiante(nombreHoja, nombreMateria, nombrePersona) {
  const { row, col } = await _encontrarCelda(nombreHoja, nombreMateria, nombrePersona);

  // Leer valor actual de la celda de la persona
  const colLetter  = colNumToLetter(col);
  const cellRef    = `'${nombreHoja}'!${colLetter}${row}`;
  const currentVal = await readRange(cellRef);
  const oldValue   = (currentVal.length && currentVal[0].length) ? Number(currentVal[0][0]) || 0 : 0;
  const newValue   = oldValue + 1;

  // Preparar escritura: siempre actualiza la celda de la persona
  const batchData = [{ range: cellRef, values: [[newValue]] }];

  // Buscar columna "CONTEO DE CUPOS" en la fila 1 y actualizar también
  try {
    const headers = await _obtenerHeaders(nombreHoja);
    const conteoCol = headers['CONTEO DE CUPOS'];
    if (conteoCol) {
      const conteoRef   = `'${nombreHoja}'!${colNumToLetter(conteoCol)}${row}`;
      const conteoActual = await readRange(conteoRef);
      const conteoViejo  = (conteoActual.length && conteoActual[0].length) ? Number(conteoActual[0][0]) || 0 : 0;
      batchData.push({ range: conteoRef, values: [[conteoViejo + 1]] });
    }
  } catch(e) {
    console.warn('No se pudo actualizar CONTEO DE CUPOS:', e.message);
  }

  await sheetsBatchUpdate({ valueInputOption: 'RAW', data: batchData });

  return `✅ Cupo registrado (+1) para "${nombrePersona}" en "${nombreMateria}". Nuevo valor: ${newValue}`;
}

/**
 * Carga los usuarios desde la hoja "Usuarios" o, si no existe, desde usuarios.json local.
 * Columnas esperadas: A=NOMBRE, B=CODIGO
 * @returns {Promise<Array<{NOMBRE: string, CODIGO: number}>>}
 */
async function obtenerUsuarios() {
  try {
    const hojas = await obtenerHojas();
    const hojaUsuarios = hojas.find(h => h.toLowerCase() === 'usuarios');

    if (hojaUsuarios) {
      const values = await readRange(`'${hojaUsuarios}'!A2:B`);
      return values
        .filter(r => r[0] && r[1])
        .map(r => ({ NOMBRE: String(r[0]).trim(), CODIGO: Number(r[1]) }));
    }
  } catch (e) {
    console.warn('No se pudo leer usuarios del Sheet, usando usuarios.json:', e.message);
  }

  // Fallback: leer usuarios.json local
  const resp = await fetch('usuarios.json');
  return resp.json();
}

/**
 * Obtiene un mapa { HEADER_NAME: colNumber(1-based) } leyendo la fila 1 completa.
 * Cachea por hoja para no hacer múltiples lecturas.
 */
const _headerCache = {};
async function _obtenerHeaders(nombreHoja) {
  if (_headerCache[nombreHoja]) return _headerCache[nombreHoja];
  const data = await readRange(`'${nombreHoja}'!1:1`);
  const row  = (data[0] || []);
  const map  = {};
  row.forEach((h, i) => { if (h) map[String(h).trim().toUpperCase()] = i + 1; });
  _headerCache[nombreHoja] = map;
  return map;
}

/**
 * Obtiene las estadísticas de cupos para una materia específica.
 * Lee las columnas "CUPOS" (máximo) y "CONTEO DE CUPOS" (usados) desde los headers de fila 1.
 * @param {string} nombreHoja
 * @param {string} nombreMateria
 * @returns {Promise<{max: number|null, usado: number, restante: number|null, porPersona: Object}>}
 */
async function obtenerCuposMateria(nombreHoja, nombreMateria) {
  // Leer headers de fila 1 para ubicar columnas clave
  const headers   = await _obtenerHeaders(nombreHoja);
  const colCupos  = headers['CUPOS'];            // columna con el máximo
  const colConteo = headers['CONTEO DE CUPOS'];  // columna con el conteo actual

  // Leer todos los datos desde fila 2 (cols A hasta la última columna del sheet)
  const materiaValues = await readRange(`'${nombreHoja}'!A2:AZ`);
  let rowData = null;
  for (let i = 0; i < materiaValues.length; i++) {
    if (materiaValues[i][0] === nombreMateria) { rowData = materiaValues[i]; break; }
  }
  if (!rowData) throw new Error(`Materia "${nombreMateria}" no encontrada.`);

  // Obtener máximo desde columna "CUPOS"
  const maxCupos = (colCupos && rowData[colCupos - 1])
    ? Number(rowData[colCupos - 1]) || null
    : null;

  // Obtener conteo desde columna "CONTEO DE CUPOS"
  const usado = (colConteo && rowData[colConteo - 1])
    ? Number(rowData[colConteo - 1]) || 0
    : 0;

  const restante = maxCupos !== null ? maxCupos - usado : null;

  // Distribución por persona (cols D:AW)
  const porPersona = {};
  const nombresRow = await readRange(`'${nombreHoja}'!D1:AW1`);
  const nombres    = (nombresRow[0] || []);
  nombres.forEach((nombre, j) => {
    const val = Number(rowData[j + 3]) || 0;
    if (nombre && val > 0) porPersona[nombre] = val;
  });

  return { max: maxCupos, usado, restante, porPersona };
}

// ─── EXPORTAR PARA USO GLOBAL ─────────────────────────────────────────────────

/**
 * Lee la hoja completa para mostrar la tabla de visualización.
 * @param {string} nombreHoja
 * @returns {Promise<{headers: string[], filas: Array<{materia, cupos, conteo, personas: Object}>}>}
 */
async function obtenerTablaCompleta(nombreHoja) {
  // Leer headers fila 1
  const headers  = await _obtenerHeaders(nombreHoja);

  // Índices clave (0-based para el array de valores)
  const iCupos   = headers['CUPOS']          ? headers['CUPOS'] - 1          : null;
  const iConteo  = headers['CONTEO DE CUPOS'] ? headers['CONTEO DE CUPOS'] - 1 : null;

  // Leer nombres de personas D1:AW1
  const nombresRow = await readRange(`'${nombreHoja}'!D1:AW1`);
  const personas   = (nombresRow[0] || []).filter(n => n && String(n).trim() !== '');

  // Leer todas las filas de datos A2:AZ
  const rows = await readRange(`'${nombreHoja}'!A2:AZ`);

  const filas = rows
    .filter(r => r[0] && String(r[0]).trim() !== '')
    .map(r => {
      const materia = String(r[0]).trim();
      const cupos   = iCupos  !== null ? (Number(r[iCupos])  || 0) : null;
      const conteo  = iConteo !== null ? (Number(r[iConteo]) || 0) : null;
      const porPersona = {};
      personas.forEach((nombre, j) => {
        porPersona[nombre] = Number(r[j + 3]) || 0; // col D = índice 3
      });
      return { materia, cupos, conteo, personas: porPersona };
    });

  return { personas, filas };
}

// ─── EXPORTAR PARA USO GLOBAL ─────────────────────────────────────────────────

window.SheetsAPI = {
  obtenerHojas,
  obtenerNombres,
  obtenerMaterias,
  obtenerCuposMateria,
  obtenerTablaCompleta,
  agregarCupoAEstudiante,
  obtenerUsuarios,
  getAccessToken
};

console.log('[app.js] SheetsAPI cargado. Spreadsheet:', SPREADSHEET_ID);
