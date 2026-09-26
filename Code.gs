/**
 * EVALUACIONES FÍSICAS - BOMBEROS
 * Backend final para:
 * GitHub Pages -> Google Apps Script -> Google Sheets
 *
 * IMPORTANTE:
 * - Este proyecto NO necesita un archivo HTML "Index".
 * - doGet() devuelve JSON para comprobar que la API está activa.
 * - doPost() recibe las acciones enviadas por index.html.
 */

const DB_PROPERTY = 'BOMBEROS_SPREADSHEET_ID';

const DESTACAMENTOS = [
  'Minas',
  'Lascano',
  'J.P. Varela',
  'Treinta y Tres',
  'Vergara',
  'Río Branco',
  'Melo',
  'Fraile Muerto',
  'Santa Clara'
];


/* =========================================================
   API
   ========================================================= */


/**
 * La exportación ya no usa UrlFetchApp.
 * Esta función queda solo como comprobación simple.
 */
function autorizarExcel() {
  Logger.log('EXPORTACIÓN SIN URLFETCH OK');
  return 'EXPORTACIÓN SIN URLFETCH OK';
}

function doGet() {
  return json_({
    ok: true,
    message: 'API Evaluaciones Bomberos activa'
  });
}


function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Solicitud sin datos.');
    }

    const body = JSON.parse(e.postData.contents || '{}');
    const action = String(body.action || '').trim();
    const data = body.data || {};

    let result;

    switch (action) {
      case 'registrarFuncionario':
        result = registrarFuncionario_(data);
        break;

      case 'loginEvaluador':
        result = loginEvaluador_(data);
        break;

      case 'cargarPanel':
        validarSesionEvaluador_(body.token);
        result = cargarPanel_();
        break;

      case 'guardarResultado':
        validarSesionEvaluador_(body.token);
        result = guardarResultado_(data);
        break;

      case 'generarExcel':
        validarSesionEvaluador_(body.token);
        result = generarExcel_(data);
        break;

      default:
        throw new Error('Acción no válida: ' + action);
    }

    return json_({
      ok: true,
      data: result
    });

  } catch (err) {
    return json_({
      ok: false,
      error: err && err.message ? err.message : String(err)
    });
  }
}


function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* =========================================================
   BASE DE DATOS
   ========================================================= */

function loginEvaluador_(data) {
  const PASSWORD = 'IBARRASEGURO2026';
  if (String(data.password || '').trim() !== PASSWORD) {
    throw new Error('Contraseña incorrecta.');
  }

  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('eval_' + token, '1', 21600); // 6 horas
  return { token: token };
}

function validarSesionEvaluador_(token) {
  if (!token || CacheService.getScriptCache().get('eval_' + token) !== '1') {
    throw new Error('Acceso del evaluador no autorizado o sesión vencida.');
  }
}

function getDb_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(DB_PROPERTY);
  let ss;

  if (id) {
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (err) {
      id = '';
    }
  }

  if (!id) {
    ss = SpreadsheetApp.create('Evaluaciones Físicas Bomberos');
    props.setProperty(DB_PROPERTY, ss.getId());
  }

  setupDb_(ss);
  return ss;
}


function setupDb_(ss) {
  ensureSheet_(ss, 'Funcionarios', [
    'ID',
    'Fecha registro',
    'Grado',
    'Nombre',
    'CI',
    'Fecha nacimiento',
    'Género',
    'Teléfono',
    'Destacamento',
    'Fecha ingreso',
    'Carné salud',
    'Vencimiento carné',
    'Ergometría',
    'Fecha ergometría',
    'Antecedentes',
    'Lesiones',
    'Enfermedad crónica','Enfermedad crónica - otro','Antecedente enfermedad','Antecedente enfermedad - otro',
    'Presenta lesión','Lesión - cuál','Lesión afecta vida cotidiana/profesional','Tareas impedidas',
    'Rehabilitación','Rehabilitación - detalle','Peso kg','Altura cm'
  ]);

  ensureSheet_(ss, 'Resultados', [
    'Fecha',
    'Año',
    'Funcionario ID',
    'Prueba',
    'Segundos',
    'Repeticiones',
    'Metros',
    'Edad',
    'Nota',
    'Categoría'
  ]);
}


function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
  }

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  return sh;
}


/* =========================================================
   FUNCIONARIOS
   ========================================================= */

function registrarFuncionario_(data) {
  const nombre = valor_(data.nombre);
  const ci = valor_(data.ci);
  const nacimiento = valor_(data.nacimiento);
  const genero = valor_(data.genero);
  const destacamento = valor_(data.destacamento || data.dest);

  if (!nombre) throw new Error('Ingresá el nombre.');
  if (!ci) throw new Error('Ingresá la CI.');
  if (!nacimiento) throw new Error('Ingresá la fecha de nacimiento.');
  if (!genero) throw new Error('Seleccioná el género.');
  if (!destacamento) throw new Error('Seleccioná el destacamento.');

  if (DESTACAMENTOS.indexOf(destacamento) === -1) {
    throw new Error('Destacamento no válido.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getDb_();
    const sh = ss.getSheetByName('Funcionarios');

    const lastRow = sh.getLastRow();
    const normalizedCI = normalizarCi_(ci);

    if (lastRow > 1) {
      const existing = sh.getRange(2, 1, lastRow - 1, 16).getValues();

      const duplicate = existing.some(row =>
        normalizarCi_(row[4]) === normalizedCI
      );

      if (duplicate) {
        throw new Error('Ya existe un funcionario registrado con esa CI.');
      }
    }

    const id = Utilities.getUuid();

    sh.appendRow([
      id,
      new Date(),
      valor_(data.grado),
      nombre,
      ci,
      nacimiento,
      genero,
      valor_(data.telefono),
      destacamento,
      valor_(data.ingreso),
      valor_(data.carne),
      valor_(data.vencCarne),
      valor_(data.ergometria || data.ergo),
      valor_(data.fechaErgo),
      valor_(data.antecedentes),
      valor_(data.lesiones),
      valor_(data.enfermedadCronica),valor_(data.enfermedadCronicaOtro),
      valor_(data.enfermedadAntecedente),valor_(data.enfermedadAntecedenteOtro),
      valor_(data.presentaLesion),valor_(data.lesionCual),valor_(data.lesionAfecta),
      valor_(data.tareasImpedidas),valor_(data.rehabilitacion),valor_(data.rehabilitacionDetalle),
      numeroOVacio_(data.peso),numeroOVacio_(data.altura)
    ]);

    return {
      id: id,
      nombre: nombre,
      destacamento: destacamento
    };

  } finally {
    lock.releaseLock();
  }
}


/* =========================================================
   PANEL DEL EVALUADOR
   ========================================================= */

function cargarPanel_() {
  const ss = getDb_();

  const shF = ss.getSheetByName('Funcionarios');
  const shR = ss.getSheetByName('Resultados');

  const people = [];
  const results = {};

  if (shF.getLastRow() > 1) {
    const rows = shF
      .getRange(2, 1, shF.getLastRow() - 1, 28)
      .getValues();

    rows.forEach(row => {
      people.push({
        id: String(row[0] || ''),
        grado: String(row[2] || ''),
        nombre: String(row[3] || ''),
        ci: String(row[4] || ''),
        nacimiento: fechaIso_(row[5]),
        genero: String(row[6] || ''),
        telefono: String(row[7] || ''),
        destacamento: String(row[8] || ''),
        dest: String(row[8] || ''),
        ingreso: fechaIso_(row[9]),
        carne: String(row[10] || ''),
        vencCarne: fechaIso_(row[11]),
        ergometria: String(row[12] || ''),
        fechaErgo: fechaIso_(row[13]),
        antecedentes: String(row[14] || ''), lesiones: String(row[15] || ''),
        enfermedadCronica: String(row[16] || ''), enfermedadCronicaOtro: String(row[17] || ''),
        enfermedadAntecedente: String(row[18] || ''), enfermedadAntecedenteOtro: String(row[19] || ''),
        presentaLesion: String(row[20] || ''), lesionCual: String(row[21] || ''),
        lesionAfecta: String(row[22] || ''), tareasImpedidas: String(row[23] || ''),
        rehabilitacion: String(row[24] || ''), rehabilitacionDetalle: String(row[25] || ''),
        peso: row[26] === '' ? '' : Number(row[26]), altura: row[27] === '' ? '' : Number(row[27])
      });
    });
  }

  if (shR.getLastRow() > 1) {
    const rows = shR
      .getRange(2, 1, shR.getLastRow() - 1, 10)
      .getValues();

    rows.forEach(row => {
      const anio = String(row[1] || '');
      const funcionarioId = String(row[2] || '');
      const prueba = String(row[3] || '');

      if (!anio || !funcionarioId || !prueba) return;

      const key = anio + '|' + funcionarioId + '|' + prueba;

      const raw = {
        fecha: fechaIso_(row[0]),
        anio: Number(row[1]) || row[1],
        funcionarioId: funcionarioId,
        prueba: prueba,
        segundos: row[4] === '' ? '' : Number(row[4]),
        repeticiones: row[5] === '' ? '' : Number(row[5]),
        metros: row[6] === '' ? '' : Number(row[6]),
        edad: row[7] === '' ? '' : Number(row[7]),
        nota: row[8] === '' ? '' : row[8],
        categoria: String(row[9] || '')
      };

      // Para registros viejos que quedaron PENDIENTE_TABLA, calcular al cargar.
      if (!raw.categoria || raw.categoria === 'PENDIENTE_TABLA') {
        const persona = people.find(p => String(p.id) === funcionarioId);
        if (persona) {
          const calc = clasificarPrueba_(
            prueba,
            raw,
            persona.genero,
            raw.edad !== '' ? raw.edad : calcularEdad_(persona.nacimiento)
          );
          if (calc.categoria) raw.categoria = calc.categoria;
          if (calc.nota !== '') raw.nota = calc.nota;
        }
      }

      results[key] = raw;
    });
  }

  return {
    destacamentos: DESTACAMENTOS,
    people: people,
    results: results
  };
}


/* =========================================================
   RESULTADOS DE LAS PRUEBAS
   ========================================================= */

function guardarResultado_(data) {
  const funcionarioId = valor_(data.funcionarioId);
  const prueba = valor_(data.prueba);

  if (!funcionarioId) {
    throw new Error('Falta el funcionario.');
  }

  if (!prueba) {
    throw new Error('Falta la prueba.');
  }

  const pruebasValidas = [
    'Flexibilidad',
    'Core',
    'Flexiones',
    'Sentadilla',
    'Yo-Yo'
  ];

  if (pruebasValidas.indexOf(prueba) === -1) {
    throw new Error('Prueba no válida.');
  }

  const anio = Number(data.anio) || new Date().getFullYear();

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getDb_();
    const sh = ss.getSheetByName('Resultados');

    const edad =
      data.edad !== undefined &&
      data.edad !== null &&
      data.edad !== ''
        ? Number(data.edad)
        : obtenerEdadFuncionario_(ss, funcionarioId);

    const segundos = numeroOVacio_(data.segundos);
    const repeticiones = numeroOVacio_(data.repeticiones);
    const metros = numeroOVacio_(data.metros);

    // NO APTO siempre es manual. Para el resto, la tabla oficial calcula la categoría.
    let categoria = valor_(data.categoria);
    let nota = data.nota === undefined || data.nota === null ? '' : data.nota;

    if (categoria !== 'NO APTO' && prueba !== 'Flexibilidad') {
      const persona = obtenerFuncionario_(ss, funcionarioId);
      const calc = clasificarPrueba_(
        prueba,
        { segundos: segundos, repeticiones: repeticiones, metros: metros },
        persona ? persona.genero : '',
        edad
      );
      categoria = calc.categoria || 'PENDIENTE_TABLA';
      nota = calc.nota;
    }

    const rowData = [
      new Date(),
      anio,
      funcionarioId,
      prueba,
      segundos,
      repeticiones,
      metros,
      edad === '' ? '' : edad,
      nota,
      categoria || 'PENDIENTE_TABLA'
    ];

    let rowToUpdate = 0;

    if (sh.getLastRow() > 1) {
      const rows = sh
        .getRange(2, 1, sh.getLastRow() - 1, 4)
        .getValues();

      for (let i = 0; i < rows.length; i++) {
        const sameYear = Number(rows[i][1]) === Number(anio);
        const samePerson = String(rows[i][2]) === funcionarioId;
        const sameTest = String(rows[i][3]) === prueba;

        if (sameYear && samePerson && sameTest) {
          rowToUpdate = i + 2;
          break;
        }
      }
    }

    if (rowToUpdate) {
      sh.getRange(rowToUpdate, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sh.appendRow(rowData);
    }

    return {
      anio: anio,
      funcionarioId: funcionarioId,
      prueba: prueba,
      segundos: rowData[4],
      repeticiones: rowData[5],
      metros: rowData[6],
      edad: rowData[7],
      nota: rowData[8],
      categoria: rowData[9]
    };

  } finally {
    lock.releaseLock();
  }
}


/* =========================================================
   TABLAS OFICIALES DE CLASIFICACIÓN
   Fuente: tablas aportadas para Core, Flexiones, Sentadilla y Yo-Yo.
   A = Excelente (o nota 8-10)
   B = Muy Bueno/Bueno (o nota 5-7)
   C = Malo (o nota 1-4)
   NO APTO se mantiene manual.
   ========================================================= */

function clasificarPrueba_(prueba, r, genero, edad) {
  const g = String(genero || '').trim().toLowerCase();
  // Acepta las etiquetas del formulario y variantes habituales.
  const masculino = g === 'masculino' || g === 'hombre' || g === 'm';
  const femenino = g === 'femenino' || g === 'mujer' || g === 'f';

  if (prueba === 'Core') {
    if (!masculino && !femenino) return { categoria: '', nota: '' };
    const s = Number(r.segundos);
    if (!isFinite(s)) return { categoria: '', nota: '' };

    // Tabla: Excelente >=115" H / >=73" M.
    // Muy Bueno + Bueno = B. Malo = C.
    if (masculino) return { categoria: s >= 115 ? 'A' : (s >= 75 ? 'B' : 'C'), nota: '' };
    return { categoria: s >= 73 ? 'A' : (s >= 48 ? 'B' : 'C'), nota: '' };
  }

  if (prueba === 'Flexiones') {
    if (!masculino && !femenino) return { categoria: '', nota: '' };
    const n = Number(r.repeticiones);
    if (!isFinite(n)) return { categoria: '', nota: '' };

    // Tabla: Excelente >=52 H / >=34 M.
    // Muy Bueno + Bueno = B. Malo = C.
    if (masculino) return { categoria: n >= 52 ? 'A' : (n >= 20 ? 'B' : 'C'), nota: '' };
    return { categoria: n >= 34 ? 'A' : (n >= 10 ? 'B' : 'C'), nota: '' };
  }

  if (prueba === 'Sentadilla') {
    if (!masculino && !femenino) return { categoria: '', nota: '' };
    const s = Number(r.segundos);
    if (!isFinite(s)) return { categoria: '', nota: '' };

    // Masculino: A >=1'40" (100 s), B >=46", C por debajo.
    // Femenino: A >=1'20" (80 s), B >=36", C por debajo.
    if (masculino) return { categoria: s >= 100 ? 'A' : (s >= 46 ? 'B' : 'C'), nota: '' };
    return { categoria: s >= 80 ? 'A' : (s >= 36 ? 'B' : 'C'), nota: '' };
  }

  if (prueba === 'Yo-Yo') {
    const m = Number(r.metros);
    const e = Number(edad);
    if (!isFinite(m) || !isFinite(e) || e < 20) return { categoria: '', nota: '' };

    let minimos;
    if (e <= 29) {
      minimos = [20,180,380,600,780,980,1180,1380,1620,1860];
    } else if (e <= 39) {
      minimos = [20,140,300,480,640,800,980,1120,1280,1440];
    } else {
      minimos = [20,120,240,380,480,620,780,940,1100,1260];
    }

    let nota = 0;
    for (let i = 0; i < minimos.length; i++) {
      if (m >= minimos[i]) nota = i + 1;
    }
    if (nota === 0) return { categoria: '', nota: '' };

    const categoria = nota >= 8 ? 'A' : (nota >= 5 ? 'B' : 'C');
    return { categoria: categoria, nota: nota };
  }

  return { categoria: '', nota: '' };
}


function obtenerFuncionario_(ss, funcionarioId) {
  const sh = ss.getSheetByName('Funcionarios');
  if (sh.getLastRow() <= 1) return null;

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 16).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(funcionarioId)) {
      return {
        id: String(rows[i][0] || ''),
        nacimiento: fechaIso_(rows[i][5]),
        genero: String(rows[i][6] || '')
      };
    }
  }
  return null;
}


/**
 * Ejecutar UNA VEZ si querés escribir en Sheets las categorías/notas
 * de resultados anteriores que quedaron como PENDIENTE_TABLA.
 */
function recalcularResultadosExistentes() {
  const ss = getDb_();
  const sh = ss.getSheetByName('Resultados');
  if (sh.getLastRow() <= 1) return { actualizados: 0 };

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  let actualizados = 0;

  for (let i = 0; i < rows.length; i++) {
    const funcionarioId = String(rows[i][2] || '');
    const prueba = String(rows[i][3] || '');
    const categoriaActual = String(rows[i][9] || '');

    if (categoriaActual === 'NO APTO' || prueba === 'Flexibilidad') continue;

    const persona = obtenerFuncionario_(ss, funcionarioId);
    if (!persona) continue;

    const edad = rows[i][7] === '' ? calcularEdad_(persona.nacimiento) : Number(rows[i][7]);
    const calc = clasificarPrueba_(prueba, {
      segundos: rows[i][4],
      repeticiones: rows[i][5],
      metros: rows[i][6]
    }, persona.genero, edad);

    if (calc.categoria) {
      sh.getRange(i + 2, 8).setValue(edad);
      sh.getRange(i + 2, 9).setValue(calc.nota);
      sh.getRange(i + 2, 10).setValue(calc.categoria);
      actualizados++;
    }
  }

  return { actualizados: actualizados };
}


/* =========================================================
   EXPORTACIÓN EXCEL - RESUMEN ANUAL
   ========================================================= */

function generarExcel_(data) {
  const anio = Number(data.anio) || new Date().getFullYear();
  const destacamentoFiltro = valor_(data.destacamento);
  const ss = getDb_();
  const shF = ss.getSheetByName('Funcionarios');
  const shR = ss.getSheetByName('Resultados');

  const funcionarios = shF.getLastRow() > 1
    ? shF.getRange(2, 1, shF.getLastRow() - 1, 16).getValues()
    : [];

  const resultados = shR.getLastRow() > 1
    ? shR.getRange(2, 1, shR.getLastRow() - 1, 10).getValues()
    : [];

  const porPersona = {};
  resultados.forEach(r => {
    if (Number(r[1]) !== anio) return;
    const id = String(r[2] || '');
    const prueba = String(r[3] || '');
    if (!porPersona[id]) porPersona[id] = {};
    porPersona[id][prueba] = {
      segundos: r[4], repeticiones: r[5], metros: r[6],
      edad: r[7], nota: r[8], categoria: String(r[9] || '')
    };
  });

  const headers = [
    'NOMBRE','EDAD','GÉNERO','DESTACAMENTO',
    'FLEXIBILIDAD','CORE - TIEMPO','CORE - CATEGORÍA',
    'FLEXIONES - REPETICIONES','FLEXIONES - CATEGORÍA',
    'SENTADILLA - TIEMPO','SENTADILLA - CATEGORÍA',
    'YO-YO - METROS','YO-YO - NOTA','YO-YO - CATEGORÍA'
  ];

  const funcionariosFiltrados = destacamentoFiltro
    ? funcionarios.filter(f => String(f[8] || '') === destacamentoFiltro)
    : funcionarios;

  const rows = funcionariosFiltrados.map(f => {
    const id = String(f[0] || '');
    const r = porPersona[id] || {};
    return [
      String(f[3] || ''), calcularEdad_(f[5]), String(f[6] || ''), String(f[8] || ''),
      r['Flexibilidad'] ? r['Flexibilidad'].categoria : '',
      r['Core'] ? r['Core'].segundos : '', r['Core'] ? r['Core'].categoria : '',
      r['Flexiones'] ? r['Flexiones'].repeticiones : '', r['Flexiones'] ? r['Flexiones'].categoria : '',
      r['Sentadilla'] ? r['Sentadilla'].segundos : '', r['Sentadilla'] ? r['Sentadilla'].categoria : '',
      r['Yo-Yo'] ? r['Yo-Yo'].metros : '', r['Yo-Yo'] ? r['Yo-Yo'].nota : '', r['Yo-Yo'] ? r['Yo-Yo'].categoria : ''
    ];
  });

  // CSV UTF-8 con separador punto y coma: Excel lo abre directamente y
  // no requiere UrlFetchApp, Drive ni permisos externos.
  const csvEscape = v => {
    const t = String(v == null ? '' : v).replace(/"/g, '""');
    return /[;"\r\n]/.test(t) ? '"' + t + '"' : t;
  };
  const csv = '\uFEFF' + [headers].concat(rows)
    .map(row => row.map(csvEscape).join(';'))
    .join('\r\n');

  const etiqueta = destacamentoFiltro
    ? destacamentoFiltro.replace(/[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ]+/g, '_')
    : 'GENERAL';
  const nombre = 'Evaluaciones_Fisicas_' + anio + '_' + etiqueta + '.csv';
  const blob = Utilities.newBlob(csv, 'text/csv;charset=utf-8', nombre);

  return {
    nombre: nombre,
    mimeType: 'text/csv;charset=utf-8',
    base64: Utilities.base64Encode(blob.getBytes()),
    filas: rows.length,
    anio: anio,
    destacamento: destacamentoFiltro || 'GENERAL'
  };
}

/* =========================================================
   UTILIDADES
   ========================================================= */

function valor_(value) {
  if (value === undefined || value === null) return '';
  return String(value).trim();
}


function numeroOVacio_(value) {
  if (value === undefined || value === null || value === '') return '';

  const n = Number(value);

  if (!isFinite(n)) {
    throw new Error('Se recibió un valor numérico no válido.');
  }

  return n;
}


function normalizarCi_(ci) {
  return String(ci || '').replace(/\D/g, '');
}


function obtenerEdadFuncionario_(ss, funcionarioId) {
  const sh = ss.getSheetByName('Funcionarios');

  if (sh.getLastRow() <= 1) return '';

  const rows = sh
    .getRange(2, 1, sh.getLastRow() - 1, 6)
    .getValues();

  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === funcionarioId) {
      return calcularEdad_(rows[i][5]);
    }
  }

  return '';
}


function calcularEdad_(fechaNacimiento) {
  if (!fechaNacimiento) return '';

  let nacimiento;

  if (
    Object.prototype.toString.call(fechaNacimiento) === '[object Date]' &&
    !isNaN(fechaNacimiento)
  ) {
    nacimiento = fechaNacimiento;
  } else {
    const text = String(fechaNacimiento).trim();

    const parts = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);

    if (parts) {
      nacimiento = new Date(
        Number(parts[1]),
        Number(parts[2]) - 1,
        Number(parts[3])
      );
    } else {
      nacimiento = new Date(text);
    }
  }

  if (isNaN(nacimiento.getTime())) return '';

  const hoy = new Date();

  let edad = hoy.getFullYear() - nacimiento.getFullYear();

  const mes = hoy.getMonth() - nacimiento.getMonth();

  if (
    mes < 0 ||
    (mes === 0 && hoy.getDate() < nacimiento.getDate())
  ) {
    edad--;
  }

  return edad >= 0 ? edad : '';
}


function fechaIso_(value) {
  if (value === undefined || value === null || value === '') return '';

  if (
    Object.prototype.toString.call(value) === '[object Date]' &&
    !isNaN(value)
  ) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  return String(value);
}


/* =========================================================
   FUNCIONES PARA CONFIGURAR / COMPROBAR
   ========================================================= */

function inicializarAplicacion() {
  const ss = getDb_();

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl()
  };
}


function getSpreadsheetUrl() {
  return getDb_().getUrl();
}
