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
const YOYO_AUDIO_START_SECONDS = 147.22;
const RESULTADOS_RESET_PROPERTY = 'BOMBEROS_RESULTADOS_RESET_MS';
const REGISTROS_RESET_PROPERTY = 'BOMBEROS_REGISTROS_RESET_MS';

function resetDependenciaKey_(comando, destacamento) {
  return 'BOMBEROS_DEP_RESET_' + encodeURIComponent(comando + '|' + destacamento);
}



// Comandos y dependencias tomados de la planilla oficial entregada.
// Se conservaron las denominaciones de la planilla y se unificaron solamente
// variantes evidentes de escritura de una misma dependencia.
const COMANDOS_DEPENDENCIAS = {
  'COBAM': [
    'Ayudantía Dirección','Comando Zona 3','Dep. Flota','Dep. Logística',
    'Departamento I','Departamento II','Depto. Sanidad',
    'Dest. Aer. Laguna del Sauce','Dsto. Aero. Melilla','Desto. Belvedere',
    'Desto. Biarritz','Desto. Canelones','Desto. Carrasco','Desto. Casavalle',
    'Dsto. Centro Cordón','Desto. Ciudad del Plata','Dsto. Las Piedras',
    'Desto. Libertad','Desto. Maroñas','Desto. Pando','Desto. Parque del Plata',
    'Dsto. San Jacinto','Desto. San José','Desto. San Ramón','Desto. Santa Lucía',
    'Desto. Solymar','Desto. Tala','Haz Mat','Jefe Región I CoBAM','Las Brujas',
    'O.N.C','Planificación y Est.','Prevención','REGION 1','REGION 2 COBAM',
    'REGION 3 COBAM','RR.PP','SYSO','Sala de Prot. Res.','Secretaria COBAM',
    'Servicios Generales','Sumario'
  ],
  'CBI': [
    'Desto. Aigua','Desto. Carmelo','Desto. Castillos','Desto. Chuy','Desto. Colonia',
    'Desto. Dolores','Desto. Durazno','Desto. Florida','Desto. Fray Bentos',
    'Desto. Jose Ignacio','Dest. José Enrique Rodó','Desto. Juan Lacaze',
    'Desto. La Paloma','Desto. Maldonado','Desto. Minas de Corrales',
    'Desto. Nueva Helvecia','Desto. Nueva Palmira','Desto. Palmita',
    'Desto. Pan de Azucar','Desto. Paso de los Toros','Desto. Piriapolis',
    'Desto. Rivera','Desto. Rocha','Desto. Rosario','Desto. San Carlos',
    'Desto. San Gregorio de Polanco','Desto. Santa Teresa','Desto. Sarandi Grande',
    'Desto. Sarandi del Yi','Desto. Tacuarembó','Desto. Tarariras','Desto. Tranqueras',
    'Desto. Trinidad','Desto. Vichadero',
    // Dependencias ya operativas en esta aplicación, bajo el comando CBI.
    'Desto. Minas','Desto. Lascano','Desto. J.P. Varela','Desto. Treinta y Tres',
    'Desto. Vergara','Desto. Río Branco','Desto. Melo','Desto. Fraile Muerto',
    'Desto. Santa Clara'
  ],
  'CBE': ['A.I.C','Dpto. IDS','Haz Mat'],
  'CD': ['Departamento I','Electricista','RR.PP','Sumario','TIC'],
  'CSD': ['Bienes y Usos Generales','CCU','Departamento II','Depto. Sanidad','O.N.C']
};

const DESTACAMENTOS = Object.keys(COMANDOS_DEPENDENCIAS).reduce(function(lista, comando) {
  return lista.concat(COMANDOS_DEPENDENCIAS[comando]);
}, []);

const DESTACAMENTOS_ANTERIORES = {
  'Minas':'Desto. Minas','Lascano':'Desto. Lascano','J.P. Varela':'Desto. J.P. Varela',
  'Treinta y Tres':'Desto. Treinta y Tres','Vergara':'Desto. Vergara',
  'Río Branco':'Desto. Río Branco','Melo':'Desto. Melo',
  'Fraile Muerto':'Desto. Fraile Muerto','Santa Clara':'Desto. Santa Clara'
};

function normalizarDependencia_(dependencia) {
  const valor = valor_(dependencia);
  return DESTACAMENTOS_ANTERIORES[valor] || valor;
}

function comandoDependencia_(dependencia) {
  const dep = normalizarDependencia_(dependencia);
  const comandos = Object.keys(COMANDOS_DEPENDENCIAS);
  for (let i = 0; i < comandos.length; i++) {
    if (COMANDOS_DEPENDENCIAS[comandos[i]].indexOf(dep) !== -1) return comandos[i];
  }
  return '';
}

function dependenciaValida_(comando, dependencia) {
  const cmd = valor_(comando);
  const dep = normalizarDependencia_(dependencia);
  return !!(COMANDOS_DEPENDENCIAS[cmd] && COMANDOS_DEPENDENCIAS[cmd].indexOf(dep) !== -1);
}





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

    message: 'API Evaluaciones Bomberos activa',
    version: 'produccion-panel-acordeones-y-sede-evaluacion-2026-09-27',
    acciones: ['actualizarFuncionario', 'confirmarEvaluadorSesionCompartida', 'estadoSesionCompartida', 'asignarFuncionarioSesionCompartida', 'iniciarCronometroSesionCompartida', 'omitirIntroduccionYoyoCompartida', 'guardarTiempoSesionCompartida', 'registrarFaltaYoyoCompartida', 'anularFaltaYoyoCompartida', 'finalizarSesionCompartida', 'eliminarEvaluacion', 'eliminarEvaluacionesFuncionario', 'eliminarEvaluacionesDestacamento', 'eliminarRegistrosDestacamento', 'reiniciarTodasEvaluaciones', 'eliminarTodosLosRegistros']

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

      case 'actualizarFuncionario':

        validarSesionEvaluador_(body.token);

        result = actualizarFuncionario_(data);

        break;

      case 'asignarSedeEvaluacion':

        validarSesionEvaluador_(body.token);

        result = asignarSedeEvaluacion_(data);

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
        bloquearGuardadoCompartido_(data);

        result = guardarResultado_(data);

        break;

      case 'eliminarEvaluacion':
        validarSesionEvaluador_(body.token);
        result = eliminarEvaluacion_(data);
        break;

      case 'eliminarEvaluacionesFuncionario':
        validarSesionEvaluador_(body.token);
        result = eliminarEvaluacionesFuncionario_(data);
        break;

      case 'eliminarEvaluacionesDestacamento':
        validarSesionEvaluador_(body.token);
        result = eliminarEvaluacionesDestacamento_(data);
        break;

      case 'eliminarRegistrosDestacamento':
        validarSesionEvaluador_(body.token);
        result = eliminarRegistrosDestacamento_(data);
        break;

      case 'reiniciarTodasEvaluaciones':
        validarSesionEvaluador_(body.token);
        result = reiniciarTodasEvaluaciones_(data);
        break;

      case 'eliminarTodosLosRegistros':
        validarSesionEvaluador_(body.token);
        result = eliminarTodosLosRegistros_(data);
        break;



      case 'generarExcel':

        validarSesionEvaluador_(body.token);

        result = generarExcel_(data);

        break;



      case 'iniciarSesionCompartida':

        validarSesionEvaluador_(body.token);

        result = iniciarSesionCompartida_(data);

        break;

      case 'confirmarEvaluadorSesionCompartida':

        validarSesionEvaluador_(body.token);

        result = confirmarEvaluadorSesionCompartida_(data);

        break;

      case 'iniciarCronometroSesionCompartida':
        validarSesionEvaluador_(body.token);
        result = iniciarCronometroSesionCompartida_(data);
        break;



      case 'obtenerSesionCompartida':
      case 'estadoSesionCompartida':

        validarSesionEvaluador_(body.token);

        result = estadoSesionCompartida_(data);

        break;



      case 'finalizarSesionCompartida':

        validarSesionEvaluador_(body.token);

        result = finalizarSesionCompartida_(data);

        break;



      case 'asignarFuncionarioSesionCompartida':
      case 'pausarTiempoSesionCompartida':
        validarSesionEvaluador_(body.token);
        result = modificarFuncionarioCompartido_(data, action === 'asignarFuncionarioSesionCompartida' ? 'asignar' : 'pausar');
        break;

      case 'guardarTiempoSesionCompartida':
        validarSesionEvaluador_(body.token);
        result = guardarTiempoCompartido_(data);
        break;

      case 'registrarFaltaYoyoCompartida':
        validarSesionEvaluador_(body.token);
        result = registrarFaltaYoyoCompartida_(data);
        break;

      case 'anularFaltaYoyoCompartida':
        validarSesionEvaluador_(body.token);
        result = anularFaltaYoyoCompartida_(data);
        break;

      case 'omitirIntroduccionYoyoCompartida':
        validarSesionEvaluador_(body.token);
        result = omitirIntroduccionYoyoCompartida_(data);
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

   SESIONES COMPARTIDAS DE CRONÓMETRO

   Un único inicio en servidor para varios dispositivos.

   ========================================================= */

function claveSesionCompartida_(data) {

  const dest = normalizarDependencia_(data.destacamento);
  const comando = valor_(data.comando) || comandoDependencia_(dest);

  const prueba = valor_(data.prueba);

  const anio = Number(data.anio) || new Date().getFullYear();

  if (!dependenciaValida_(comando, dest)) throw new Error('Comando o dependencia no válidos.');

  if (['Core','Sentadilla','Yo-Yo'].indexOf(prueba) === -1) throw new Error('Esta prueba no admite sesión compartida.');

  return 'SESION_COMPARTIDA_' + anio + '_' + comando + '_' + dest + '_' + prueba;

}



function iniciarSesionCompartida_(data) {
  return iniciarCronometroSesionCompartida_(data);
}

function confirmarEvaluadorSesionCompartida_(data) {

  const key = claveSesionCompartida_(data);

  const rol = valor_(data.evaluador);

  const clienteId = valor_(data.clienteId);

  if (['Evaluador 1','Evaluador 2'].indexOf(rol) === -1) throw new Error('Seleccioná Evaluador 1 o Evaluador 2.');

  if (!clienteId || clienteId.length < 12) throw new Error('No se pudo identificar este dispositivo. Recargá la página.');

  const lock = LockService.getScriptLock(); lock.waitLock(10000);

  try {

    const props = PropertiesService.getScriptProperties();

    const existente = props.getProperty(key);

    let ses = existente ? JSON.parse(existente) : null;

    if (!ses || ses.estado === 'FINALIZADA' || !ses.evaluadores) {

      ses = {

        id: Utilities.getUuid(), comando: valor_(data.comando) || comandoDependencia_(data.destacamento),
        destacamento: normalizarDependencia_(data.destacamento),

        prueba: valor_(data.prueba), anio: Number(data.anio) || new Date().getFullYear(),

        estado: 'ESPERANDO', inicioMs: null,

        evaluadores: {'Evaluador 1': null, 'Evaluador 2': null}, funcionarios: {}

      };

    }

    const otroRol = rol === 'Evaluador 1' ? 'Evaluador 2' : 'Evaluador 1';

    if (ses.evaluadores[otroRol] && ses.evaluadores[otroRol].clienteId === clienteId) {

      throw new Error('Este dispositivo ya confirmó como ' + otroRol + '. El otro evaluador debe entrar desde su dispositivo.');

    }

    if (ses.evaluadores[rol] && ses.evaluadores[rol].clienteId !== clienteId) {

      throw new Error(rol + ' ya está ocupado. Seleccioná el otro rol o finalizá esta sesión.');

    }

    if (!ses.evaluadores[rol]) ses.evaluadores[rol] = {clienteId: clienteId, confirmadoMs: Date.now()};

    if (ses.estado === 'ESPERANDO' && ses.evaluadores['Evaluador 1'] && ses.evaluadores['Evaluador 2']) ses.estado = 'LISTA';

    if (!ses.funcionarios) ses.funcionarios = {};
    props.setProperty(key, JSON.stringify(ses));

    return Object.assign({}, ses, {serverNow: Date.now(), miRol: rol});

  } finally { lock.releaseLock(); }

}

function iniciarCronometroSesionCompartida_(data) {
  const key = claveSesionCompartida_(data);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties(), raw = props.getProperty(key);
    if (!raw) throw new Error('Primero deben confirmar ambos evaluadores.');
    const ses = JSON.parse(raw);
    validarDueñoCompartido_(ses, data);
    if (ses.estado === 'ACTIVA') return Object.assign({}, ses, {serverNow:Date.now()});
    if (ses.estado !== 'LISTA') throw new Error('Falta confirmar uno de los evaluadores.');
    const registros = Object.values(ses.funcionarios || {});
    if (!registros.some(r => r.evaluador === 'Evaluador 1') || !registros.some(r => r.evaluador === 'Evaluador 2')) {
      throw new Error('Cada evaluador debe seleccionar al menos un funcionario antes de iniciar.');
    }
    registros.forEach(r => { if (r.estado === 'ASIGNADO') r.estado = 'EN_CURSO'; });
    const ahora = Date.now();
    ses.inicioMs = valor_(ses.prueba) === 'Yo-Yo' && ses.introduccionOmitida
      ? ahora - Math.round(YOYO_AUDIO_START_SECONDS * 1000)
      : ahora;
    ses.estado = 'ACTIVA'; ses.iniciadoPor = valor_(data.evaluador);
    props.setProperty(key, JSON.stringify(ses));
    return Object.assign({}, ses, {serverNow:Date.now()});
  } finally { lock.releaseLock(); }
}



function estadoSesionCompartida_(data) {

  const key = claveSesionCompartida_(data);

  const raw = PropertiesService.getScriptProperties().getProperty(key);

  if (!raw) return {estado:'SIN_SESION', serverNow:Date.now()};

  const ses = JSON.parse(raw);

  return Object.assign({}, ses, {serverNow:Date.now()});

}



function validarDueñoCompartido_(ses, data) {
  const rol = valor_(data.evaluador), clienteId = valor_(data.clienteId);
  if (['ESPERANDO','LISTA','ACTIVA'].indexOf(ses.estado) === -1) throw new Error('La sesión compartida no está disponible.');
  if (!ses.evaluadores || !ses.evaluadores[rol] || ses.evaluadores[rol].clienteId !== clienteId) {
    throw new Error('Confirmá tu puesto en este dispositivo antes de controlar funcionarios.');
  }
  return rol;
}

function bloquearGuardadoCompartido_(data) {
  if (['Core','Sentadilla','Yo-Yo'].indexOf(valor_(data.prueba)) === -1) return;
  const funcionarioId = valor_(data.funcionarioId);
  const dest = normalizarDependencia_(data.destacamento || destacamentoEvaluacionFuncionario_(funcionarioId));
  const key = claveSesionCompartida_({comando:valor_(data.comando) || comandoEvaluacionFuncionario_(funcionarioId) || comandoDependencia_(dest), destacamento:dest, prueba:data.prueba, anio:data.anio});
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return;
  const ses = JSON.parse(raw), registro = (ses.funcionarios || {})[valor_(data.funcionarioId)];
  if (ses.estado === 'ACTIVA' && registro) throw new Error('Este funcionario está en sesión compartida. Solo su evaluador responsable puede finalizar el tiempo.');
}

function destacamentoFuncionario_(id) {
  const persona = obtenerFuncionario_(getDb_(), id);
  return persona ? valor_(persona.destacamento || persona.dest) : '';
}

function comandoFuncionario_(id) {
  const persona = obtenerFuncionario_(getDb_(), id);
  return persona ? valor_(persona.comando) || comandoDependencia_(persona.destacamento || persona.dest) : '';
}

function destacamentoEvaluacionFuncionario_(id) {
  const persona = obtenerFuncionario_(getDb_(), id);
  return persona ? valor_(persona.destacamentoEvaluacion) || valor_(persona.destacamento || persona.dest) : '';
}

function comandoEvaluacionFuncionario_(id) {
  const persona = obtenerFuncionario_(getDb_(), id);
  if (!persona) return '';
  const dest = valor_(persona.destacamentoEvaluacion) || valor_(persona.destacamento || persona.dest);
  return valor_(persona.comandoEvaluacion) || valor_(persona.comando) || comandoDependencia_(dest);
}

function modificarFuncionarioCompartido_(data, accion) {
  const key = claveSesionCompartida_(data), id = valor_(data.funcionarioId);
  if (!id) throw new Error('Falta el funcionario.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties(), raw = props.getProperty(key);
    if (!raw) throw new Error('Primero iniciá la sesión compartida.');
    const ses = JSON.parse(raw), rol = validarDueñoCompartido_(ses, data);
    if (destacamentoEvaluacionFuncionario_(id) !== ses.destacamento || comandoEvaluacionFuncionario_(id) !== ses.comando) {
      throw new Error('El funcionario no está asignado a esta sede de evaluación.');
    }
    ses.funcionarios = ses.funcionarios || {};
    const actual = ses.funcionarios[id];
    if (accion === 'asignar') {
      if (data.asignado) {
        if (actual && actual.evaluador !== rol) throw new Error('Este funcionario lo controla ' + actual.evaluador + '.');
        if (ses.estado === 'ACTIVA') throw new Error('El cronómetro ya comenzó; no se pueden cambiar las asignaciones.');
        if (!actual) ses.funcionarios[id] = {evaluador:rol, estado:'ASIGNADO'};
      } else {
        if (actual && actual.evaluador !== rol) throw new Error('Solo el evaluador responsable puede quitarlo.');
        if (ses.estado === 'ACTIVA') throw new Error('El cronómetro ya comenzó; no se pueden cambiar las asignaciones.');
        if (actual && actual.estado !== 'ASIGNADO') throw new Error('Este tiempo ya fue registrado.');
        delete ses.funcionarios[id];
      }
    } else {
      if (ses.estado !== 'ACTIVA') throw new Error('Esperá el inicio sincronizado.');
      if (!actual || actual.evaluador !== rol) throw new Error('Solo el evaluador responsable puede pausar este tiempo.');
      if (actual.estado !== 'EN_CURSO') throw new Error('El tiempo ya está pausado.');
      actual.segundos = Number(((Date.now() - Number(ses.inicioMs)) / 1000).toFixed(1));
      actual.estado = 'PAUSADO'; actual.pausadoMs = Date.now();
    }
    props.setProperty(key, JSON.stringify(ses));
    return Object.assign({}, ses, {serverNow:Date.now()});
  } finally { lock.releaseLock(); }
}

function guardarTiempoCompartido_(data) {
  const key = claveSesionCompartida_(data), id = valor_(data.funcionarioId);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  let segundos, sesId;
  try {
    const props = PropertiesService.getScriptProperties(), raw = props.getProperty(key);
    if (!raw) throw new Error('No hay sesión compartida.');
    const ses = JSON.parse(raw), rol = validarDueñoCompartido_(ses, data);
    const registro = (ses.funcionarios || {})[id];
    if (!registro || registro.evaluador !== rol) throw new Error('Solo el evaluador responsable puede finalizar este tiempo.');
    if (ses.estado !== 'ACTIVA') throw new Error('El cronómetro todavía no comenzó.');
    if (registro.estado !== 'EN_CURSO') throw new Error('Este tiempo ya fue registrado.');
    const calculadoServidor = Number(((Date.now() - Number(ses.inicioMs)) / 1000).toFixed(1));
    const capturado = Number(data.segundosCapturados);
    segundos = Number.isFinite(capturado) && capturado >= 0 && capturado <= calculadoServidor + 2
      ? Number(capturado.toFixed(1))
      : calculadoServidor;
    sesId = ses.id;
    registro.segundos = segundos; registro.pausadoMs = Date.now();
    registro.estado = 'GUARDANDO';
    props.setProperty(key, JSON.stringify(ses));
  } finally { lock.releaseLock(); }
  let guardado;
  try {
    guardado = guardarResultado_({funcionarioId:id, prueba:valor_(data.prueba), anio:Number(data.anio), segundos:segundos});
  } catch (err) {
    const retryLock = LockService.getScriptLock(); retryLock.waitLock(10000);
    try {
      const props = PropertiesService.getScriptProperties(), ses = JSON.parse(props.getProperty(key) || 'null');
      if (ses && ses.id === sesId && ses.funcionarios[id]?.estado === 'GUARDANDO') {
        ses.funcionarios[id].estado = 'EN_CURSO';
        delete ses.funcionarios[id].segundos;
        delete ses.funcionarios[id].pausadoMs;
        props.setProperty(key, JSON.stringify(ses));
      }
    } finally { retryLock.releaseLock(); }
    throw err;
  }
  const finishLock = LockService.getScriptLock(); finishLock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties(), ses = JSON.parse(props.getProperty(key) || 'null');
    if (ses && ses.id === sesId && ses.funcionarios[id]) {
      ses.funcionarios[id].estado = 'GUARDADO';
      ses.funcionarios[id].guardadoMs = Date.now();
      ses.funcionarios[id].categoria = guardado.categoria;
      ses.funcionarios[id].nota = guardado.nota;
      props.setProperty(key, JSON.stringify(ses));
    }
    return {sesion:ses, resultado:guardado, serverNow:Date.now()};
  } finally { finishLock.releaseLock(); }
}

function registrarFaltaYoyoCompartida_(data) {
  if (valor_(data.prueba) !== 'Yo-Yo') throw new Error('Esta acción corresponde únicamente al Yo-Yo Test.');
  const key = claveSesionCompartida_(data), id = valor_(data.funcionarioId);
  if (!id) throw new Error('Falta el funcionario.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  let sesId, segundosAudio, metros;
  try {
    const props = PropertiesService.getScriptProperties(), raw = props.getProperty(key);
    if (!raw) throw new Error('No hay una sesión compartida del Yo-Yo.');
    const ses = JSON.parse(raw), rol = validarDueñoCompartido_(ses, data);
    const registro = (ses.funcionarios || {})[id];
    if (!registro || registro.evaluador !== rol) throw new Error('Solo podés registrar faltas de tus funcionarios.');
    if (ses.estado !== 'ACTIVA') throw new Error('El Yo-Yo todavía no comenzó.');
    if (registro.estado !== 'EN_CURSO') throw new Error('Este funcionario ya finalizó la prueba.');
    const calculadoServidor = Math.max(0, (Date.now() - Number(ses.inicioMs)) / 1000);
    if (calculadoServidor < YOYO_AUDIO_START_SECONDS) throw new Error('La carrera todavía no comenzó. Esperá el final de la introducción.');
    const faltas = Number(registro.faltas || 0) + 1;
    registro.faltas = faltas;
    registro.ultimaFaltaMs = Date.now();
    if (faltas < 2) {
      props.setProperty(key, JSON.stringify(ses));
      return {sesion:Object.assign({}, ses, {serverNow:Date.now()}), segundaFalta:false};
    }
    const capturado = Number(data.segundosCapturados);
    segundosAudio = Number.isFinite(capturado) && capturado >= YOYO_AUDIO_START_SECONDS && capturado <= calculadoServidor + 2
      ? capturado : calculadoServidor;
    metros = Math.max(0, Math.min(4420, Math.floor(Number(data.metrosCapturados || 0) / 20) * 20));
    registro.estado = 'GUARDANDO';
    registro.segundosAudio = Number(segundosAudio.toFixed(1));
    registro.segundos = Number(Math.max(0, segundosAudio - YOYO_AUDIO_START_SECONDS).toFixed(1));
    registro.metros = metros;
    sesId = ses.id;
    props.setProperty(key, JSON.stringify(ses));
  } finally { lock.releaseLock(); }

  let guardado;
  try {
    guardado = guardarResultado_({
      funcionarioId:id, prueba:'Yo-Yo', anio:Number(data.anio),
      segundos:Number(Math.max(0, segundosAudio - YOYO_AUDIO_START_SECONDS).toFixed(1)), metros:metros
    });
  } catch (err) {
    const retryLock = LockService.getScriptLock(); retryLock.waitLock(10000);
    try {
      const props = PropertiesService.getScriptProperties(), ses = JSON.parse(props.getProperty(key) || 'null');
      if (ses && ses.id === sesId && ses.funcionarios[id]?.estado === 'GUARDANDO') {
        ses.funcionarios[id].estado = 'EN_CURSO'; ses.funcionarios[id].faltas = 1;
        delete ses.funcionarios[id].segundos; delete ses.funcionarios[id].segundosAudio; delete ses.funcionarios[id].metros;
        props.setProperty(key, JSON.stringify(ses));
      }
    } finally { retryLock.releaseLock(); }
    throw err;
  }

  const finishLock = LockService.getScriptLock(); finishLock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties(), ses = JSON.parse(props.getProperty(key) || 'null');
    if (ses && ses.id === sesId && ses.funcionarios[id]) {
      const registro = ses.funcionarios[id];
      registro.estado = 'GUARDADO'; registro.guardadoMs = Date.now();
      registro.categoria = guardado.categoria; registro.nota = guardado.nota;
      props.setProperty(key, JSON.stringify(ses));
    }
    return {sesion:Object.assign({}, ses, {serverNow:Date.now()}), resultado:guardado, segundaFalta:true};
  } finally { finishLock.releaseLock(); }
}

function anularFaltaYoyoCompartida_(data) {
  if (valor_(data.prueba) !== 'Yo-Yo') throw new Error('Esta acción corresponde únicamente al Yo-Yo Test.');
  const key = claveSesionCompartida_(data), id = valor_(data.funcionarioId);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties(), raw = props.getProperty(key);
    if (!raw) throw new Error('No hay una sesión compartida del Yo-Yo.');
    const ses = JSON.parse(raw), rol = validarDueñoCompartido_(ses, data);
    const registro = (ses.funcionarios || {})[id];
    if (!registro || registro.evaluador !== rol) throw new Error('Solo podés anular faltas de tus funcionarios.');
    if (ses.estado !== 'ACTIVA' || registro.estado !== 'EN_CURSO') throw new Error('La falta ya no puede modificarse.');
    registro.faltas = Math.max(0, Number(registro.faltas || 0) - 1);
    props.setProperty(key, JSON.stringify(ses));
    return Object.assign({}, ses, {serverNow:Date.now()});
  } finally { lock.releaseLock(); }
}

function omitirIntroduccionYoyoCompartida_(data) {
  if (valor_(data.prueba) !== 'Yo-Yo') throw new Error('Esta acción corresponde únicamente al Yo-Yo Test.');
  const key = claveSesionCompartida_(data);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties(), raw = props.getProperty(key);
    if (!raw) throw new Error('No hay sesión compartida del Yo-Yo.');
    const ses = JSON.parse(raw);
    validarDueñoCompartido_(ses, data);
    if (ses.estado === 'ACTIVA') throw new Error('La prueba ya fue iniciada; la introducción se omite antes de comenzar.');
    if (ses.estado !== 'LISTA') throw new Error('Primero deben confirmar ambos evaluadores.');
    const now = Date.now();
    ses.introduccionOmitida = true;
    ses.introduccionOmitidaMs = now;
    ses.introduccionOmitidaPor = valor_(data.evaluador);
    props.setProperty(key, JSON.stringify(ses));
    return Object.assign({}, ses, {serverNow:Date.now()});
  } finally { lock.releaseLock(); }
}

function finalizarSesionCompartida_(data) {

  const key = claveSesionCompartida_(data);

  const lock = LockService.getScriptLock(); lock.waitLock(10000);

  try {

    const props = PropertiesService.getScriptProperties();

    const raw = props.getProperty(key);

    if (!raw) return {estado:'SIN_SESION', serverNow:Date.now()};

    const ses = JSON.parse(raw);
    validarDueñoCompartido_(ses, data);
    if (Object.values(ses.funcionarios || {}).some(r => r.estado === 'PAUSADO' || r.estado === 'GUARDANDO')) {
      throw new Error('Hay tiempos pausados sin guardar. Cada evaluador debe finalizar sus funcionarios antes de reiniciar la sesión.');
    }

    ses.estado='FINALIZADA'; ses.finMs=Date.now(); ses.finalizadaPor=valor_(data.evaluador)||'Evaluador';

    props.setProperty(key, JSON.stringify(ses));

    return Object.assign({}, ses, {serverNow:Date.now()});

  } finally { lock.releaseLock(); }

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

    'Rehabilitación','Rehabilitación - detalle','Peso kg','Altura cm','Comando','Apellido',

    'Comando sede evaluación','Sede de evaluación'

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

  if (sh.getMaxColumns() < headers.length) {

    sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());

  }



  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  sh.setFrozenRows(1);

  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  return sh;

}





/* =========================================================

   FUNCIONARIOS

   ========================================================= */

function validarFichaFuncionario_(data) {
  const f = {
    grado: valor_(data.grado), nombre: valor_(data.nombre), apellido: valor_(data.apellido),
    ci: valor_(data.ci), nacimiento: valor_(data.nacimiento), genero: valor_(data.genero),
    telefono: valor_(data.telefono), comando: valor_(data.comando),
    destacamento: normalizarDependencia_(data.destacamento || data.dest), ingreso: valor_(data.ingreso),
    carne: valor_(data.carne), vencCarne: valor_(data.vencCarne),
    ergometria: valor_(data.ergometria || data.ergo), fechaErgo: valor_(data.fechaErgo),
    enfermedadCronica: valor_(data.enfermedadCronica), enfermedadCronicaOtro: valor_(data.enfermedadCronicaOtro),
    enfermedadAntecedente: valor_(data.enfermedadAntecedente), enfermedadAntecedenteOtro: valor_(data.enfermedadAntecedenteOtro),
    presentaLesion: valor_(data.presentaLesion), lesionCual: valor_(data.lesionCual),
    lesionAfecta: valor_(data.lesionAfecta), tareasImpedidas: valor_(data.tareasImpedidas),
    rehabilitacion: valor_(data.rehabilitacion), rehabilitacionDetalle: valor_(data.rehabilitacionDetalle),
    peso: numeroOVacio_(data.peso), altura: numeroOVacio_(data.altura),
    antecedentes: valor_(data.antecedentes), lesiones: valor_(data.lesiones)
  };

  if (!f.grado) throw new Error('Ingresá el grado.');
  if (!f.nombre) throw new Error('Ingresá el nombre.');
  if (!f.apellido) throw new Error('Ingresá el apellido.');
  if (!f.ci) throw new Error('Ingresá la CI.');
  if (!f.nacimiento) throw new Error('Ingresá la fecha de nacimiento.');
  const partes = /^(\d{4})-(\d{2})-(\d{2})$/.exec(f.nacimiento);
  if (!partes) throw new Error('Ingresá día, mes y año de nacimiento válidos.');
  const fechaNac = new Date(Number(partes[1]), Number(partes[2]) - 1, Number(partes[3]));
  if (fechaNac.getFullYear() !== Number(partes[1]) || fechaNac.getMonth() !== Number(partes[2]) - 1 || fechaNac.getDate() !== Number(partes[3]) || fechaNac > new Date()) throw new Error('Fecha de nacimiento no válida.');
  if (!f.genero) throw new Error('Seleccioná el género.');
  if (!f.telefono) throw new Error('Ingresá el teléfono.');
  if (!f.comando) throw new Error('Seleccioná el comando.');
  if (!f.destacamento) throw new Error('Seleccioná la dependencia.');
  if (!dependenciaValida_(f.comando, f.destacamento)) throw new Error('La dependencia no pertenece al comando seleccionado.');
  if (!f.ingreso) throw new Error('Ingresá la fecha de ingreso.');
  if (!f.carne) throw new Error('Indicá si el carné de salud está vigente.');
  if (f.carne === 'Sí' && !f.vencCarne) throw new Error('Ingresá el vencimiento del carné de salud.');

  const edad = calcularEdad_(f.nacimiento);
  if (Number(edad) >= 35 && !f.ergometria) throw new Error('La ergometría es obligatoria desde los 35 años inclusive.');
  if (f.ergometria === 'Sí' && !f.fechaErgo) throw new Error('Ingresá la fecha de realización de la ergometría.');

  if (!f.enfermedadCronica) throw new Error('Completá enfermedades crónicas.');
  if (f.enfermedadCronica === 'Otros' && !f.enfermedadCronicaOtro) throw new Error('Especificá la enfermedad crónica.');
  if (!f.enfermedadAntecedente) throw new Error('Completá los antecedentes de enfermedad.');
  if (f.enfermedadAntecedente === 'Otros' && !f.enfermedadAntecedenteOtro) throw new Error('Especificá el antecedente de enfermedad.');
  if (!f.presentaLesion) throw new Error('Indicá si presenta alguna lesión.');
  if (f.presentaLesion === 'Sí') {
    if (!f.lesionCual) throw new Error('Especificá cuál es la lesión.');
    if (!f.lesionAfecta) throw new Error('Indicá si la lesión afecta su vida cotidiana o profesional.');
    if (f.lesionAfecta === 'Sí' && !f.tareasImpedidas) throw new Error('Especificá qué tareas le impide realizar.');
    if (!f.rehabilitacion) throw new Error('Indicá si realiza rehabilitación.');
    if (f.rehabilitacion === 'Sí' && !f.rehabilitacionDetalle) throw new Error('Detallá el tratamiento de rehabilitación.');
  }
  if (!(Number(f.peso) > 0)) throw new Error('Ingresá un peso válido.');
  if (!(Number(f.altura) > 0)) throw new Error('Ingresá una altura válida.');
  return f;
}

function filaFuncionario_(id, fechaRegistro, f, sedeEvaluacion) {
  sedeEvaluacion = sedeEvaluacion || {};
  return [
    id, fechaRegistro, f.grado, f.nombre, f.ci, f.nacimiento, f.genero, f.telefono,
    f.destacamento, f.ingreso, f.carne, f.vencCarne, f.ergometria, f.fechaErgo,
    f.antecedentes, f.lesiones, f.enfermedadCronica, f.enfermedadCronicaOtro,
    f.enfermedadAntecedente, f.enfermedadAntecedenteOtro, f.presentaLesion,
    f.lesionCual, f.lesionAfecta, f.tareasImpedidas, f.rehabilitacion,
    f.rehabilitacionDetalle, f.peso, f.altura, f.comando, f.apellido,
    valor_(sedeEvaluacion.comandoEvaluacion), normalizarDependencia_(sedeEvaluacion.destacamentoEvaluacion)
  ];
}

function actualizarFuncionario_(data) {
  const id = valor_(data.funcionarioId || data.id);
  if (!id) throw new Error('Funcionario no identificado.');
  const ficha = validarFichaFuncionario_(data);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = getDb_().getSheetByName('Funcionarios');
    const lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error('No se encontró el funcionario.');
    const rows = sh.getRange(2, 1, lastRow - 1, 32).getValues();
    const index = rows.findIndex(row => String(row[0] || '') === id);
    if (index < 0) throw new Error('No se encontró el funcionario.');
    const normalizedCI = normalizarCi_(ficha.ci);
    const duplicate = rows.some((row, i) => i !== index && normalizarCi_(row[4]) === normalizedCI);
    if (duplicate) throw new Error('La CI ingresada pertenece a otro funcionario.');
    const sheetRow = index + 2;
    const fechaRegistro = rows[index][1] || new Date();
    const sedeEvaluacion = {comandoEvaluacion:String(rows[index][30] || ''), destacamentoEvaluacion:normalizarDependencia_(rows[index][31])};
    sh.getRange(sheetRow, 1, 1, 32).setValues([filaFuncionario_(id, fechaRegistro, ficha, sedeEvaluacion)]);
    return {funcionario: Object.assign({id:id, nombreSolo:ficha.nombre, nombre:[ficha.nombre,ficha.apellido].filter(Boolean).join(' '), dest:ficha.destacamento}, ficha, sedeEvaluacion)};
  } finally {
    lock.releaseLock();
  }
}

/** Asigna una sede operativa sin modificar la dependencia de origen. */
function asignarSedeEvaluacion_(data) {
  const id = valor_(data.funcionarioId || data.id);
  if (!id) throw new Error('Funcionario no identificado.');
  let comandoEvaluacion = valor_(data.comandoEvaluacion);
  let destacamentoEvaluacion = normalizarDependencia_(data.destacamentoEvaluacion);
  if (!!comandoEvaluacion !== !!destacamentoEvaluacion) throw new Error('Seleccioná el comando y la sede de evaluación.');
  if (destacamentoEvaluacion && !dependenciaValida_(comandoEvaluacion, destacamentoEvaluacion)) {
    throw new Error('La sede de evaluación no pertenece al comando seleccionado.');
  }

  asegurarSinSesionActiva_({funcionarioId:id});
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = getDb_().getSheetByName('Funcionarios');
    if (!sh || sh.getLastRow() < 2) throw new Error('No se encontró el funcionario.');
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 32).getValues();
    const index = rows.findIndex(row => String(row[0] || '') === id);
    if (index < 0) throw new Error('No se encontró el funcionario.');
    const origen = normalizarDependencia_(rows[index][8]);
    const comandoOrigen = String(rows[index][28] || '') || comandoDependencia_(origen);
    if (comandoEvaluacion === comandoOrigen && destacamentoEvaluacion === origen) {
      comandoEvaluacion = '';
      destacamentoEvaluacion = '';
    }
    sh.getRange(index + 2, 31, 1, 2).setValues([[comandoEvaluacion, destacamentoEvaluacion]]);
    limpiarSesionesCompartidas_({funcionarioId:id});
    return {funcionarioId:id, comandoOrigen:comandoOrigen, destacamentoOrigen:origen,
      comandoEvaluacion:comandoEvaluacion, destacamentoEvaluacion:destacamentoEvaluacion};
  } finally { lock.releaseLock(); }
}



function registrarFuncionario_(data) {

  const resetRegistrosMs = Number(PropertiesService.getScriptProperties().getProperty(REGISTROS_RESET_PROPERTY) || 0);

  const registroCreadoMs = Number(data._clientCreatedMs || 0);

  if (resetRegistrosMs && (!registroCreadoMs || registroCreadoMs < resetRegistrosMs)) {

    throw new Error('Esta ficha pertenece a una versión anterior al reinicio general. Recargá la página y registrala nuevamente.');

  }

  let nombre = valor_(data.nombre);
  let apellido = valor_(data.apellido);
  if (!apellido && nombre) {
    const anterior = separarNombre_(nombre);
    nombre = anterior.nombre;
    apellido = anterior.apellido;
  }

  const ci = valor_(data.ci);

  const nacimiento = valor_(data.nacimiento);

  const genero = valor_(data.genero);

  const destacamento = normalizarDependencia_(data.destacamento || data.dest);
  const comando = valor_(data.comando) || comandoDependencia_(destacamento);
  const resetDependenciaMs = Number(PropertiesService.getScriptProperties().getProperty(resetDependenciaKey_(comando, destacamento)) || 0);
  if (resetDependenciaMs && (!registroCreadoMs || registroCreadoMs < resetDependenciaMs)) {
    throw new Error('Esta ficha es anterior al borrado de la dependencia. Recargá la página y registrala nuevamente.');
  }



  if (!nombre) throw new Error('Ingresá el nombre.');
  if (!apellido) throw new Error('Ingresá el apellido.');

  if (!ci) throw new Error('Ingresá la CI.');

  if (!nacimiento) throw new Error('Ingresá la fecha de nacimiento.');
  const partesNacimiento = /^(\d{4})-(\d{2})-(\d{2})$/.exec(nacimiento);
  if (!partesNacimiento) throw new Error('Ingresá día, mes y año de nacimiento válidos.');
  const fechaNac = new Date(Number(partesNacimiento[1]), Number(partesNacimiento[2]) - 1, Number(partesNacimiento[3]));
  if (fechaNac.getFullYear() !== Number(partesNacimiento[1]) || fechaNac.getMonth() !== Number(partesNacimiento[2]) - 1 || fechaNac.getDate() !== Number(partesNacimiento[3]) || fechaNac > new Date()) throw new Error('Fecha de nacimiento no válida.');

  if (!genero) throw new Error('Seleccioná el género.');

  if (!comando) throw new Error('Seleccioná el comando.');
  if (!destacamento) throw new Error('Seleccioná la dependencia.');



  if (!dependenciaValida_(comando, destacamento)) {

    throw new Error('La dependencia no pertenece al comando seleccionado.');

  }

  const ficha = validarFichaFuncionario_(data);



  const lock = LockService.getScriptLock();

  lock.waitLock(30000);



  try {

    const ss = getDb_();

    const sh = ss.getSheetByName('Funcionarios');



    const lastRow = sh.getLastRow();

    const normalizedCI = normalizarCi_(ficha.ci);



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



    sh.appendRow(filaFuncionario_(id, new Date(), ficha));



    return {

      id: id,

      nombre: [ficha.nombre, ficha.apellido].filter(Boolean).join(' '),

      destacamento: ficha.destacamento

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

      .getRange(2, 1, shF.getLastRow() - 1, 32)

      .getValues();



    rows.forEach(row => {
      const dependencia = normalizarDependencia_(row[8]);
      const comando = String(row[28] || '') || comandoDependencia_(dependencia);
      const nombreRegistrado = String(row[3] || '');
      const apellidoRegistrado = String(row[29] || '');
      const nombreAnterior = apellidoRegistrado ? {nombre:nombreRegistrado, apellido:apellidoRegistrado} : separarNombre_(nombreRegistrado);

      people.push({

        id: String(row[0] || ''),

        grado: String(row[2] || ''),

        nombre: [nombreAnterior.nombre, nombreAnterior.apellido].filter(Boolean).join(' '),
        nombreSolo: nombreAnterior.nombre,
        apellido: nombreAnterior.apellido,

        ci: String(row[4] || ''),

        nacimiento: fechaIso_(row[5]),

        genero: String(row[6] || ''),

        telefono: String(row[7] || ''),

        comando: comando,
        destacamento: dependencia,

        comandoEvaluacion: String(row[30] || ''),
        destacamentoEvaluacion: normalizarDependencia_(row[31]),

        dest: dependencia,

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

    comandos: COMANDOS_DEPENDENCIAS,
    destacamentos: DESTACAMENTOS,

    people: people,

    results: results

  };

}





/* =========================================================

   RESULTADOS DE LAS PRUEBAS

   ========================================================= */

function eliminarEvaluacion_(data) {
  const funcionarioId = valor_(data.funcionarioId);
  const prueba = valor_(data.prueba);
  const anio = Number(data.anio) || new Date().getFullYear();
  const pruebasValidas = ['Flexibilidad','Core','Flexiones','Sentadilla','Yo-Yo'];
  if (!funcionarioId) throw new Error('Falta el funcionario.');
  if (pruebasValidas.indexOf(prueba) === -1) throw new Error('Prueba no válida.');

  if (['Core','Sentadilla','Yo-Yo'].indexOf(prueba) !== -1) {
    const destacamento = normalizarDependencia_(data.destacamento || destacamentoEvaluacionFuncionario_(funcionarioId));
    const comando = valor_(data.comando) || comandoEvaluacionFuncionario_(funcionarioId) || comandoDependencia_(destacamento);
    const key = claveSesionCompartida_({comando:comando, destacamento:destacamento, prueba:prueba, anio:anio});
    const raw = PropertiesService.getScriptProperties().getProperty(key);
    if (raw) {
      const ses = JSON.parse(raw);
      if (ses.estado === 'ACTIVA') throw new Error('Finalizá la sesión compartida antes de eliminar este resultado.');
    }
  }

  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = getDb_().getSheetByName('Resultados');
    if (!sh || sh.getLastRow() <= 1) return {eliminado:false, anio:anio, funcionarioId:funcionarioId, prueba:prueba};
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    let eliminados = 0;
    for (let i = rows.length - 1; i >= 0; i--) {
      if (Number(rows[i][1]) === anio && String(rows[i][2]) === funcionarioId && String(rows[i][3]) === prueba) {
        sh.deleteRow(i + 2); eliminados++;
      }
    }
    return {eliminado:eliminados > 0, eliminados:eliminados, anio:anio, funcionarioId:funcionarioId, prueba:prueba};
  } finally { lock.releaseLock(); }
}

/** Elimina todas las evaluaciones históricas de un funcionario, pero conserva su ficha. */
function eliminarEvaluacionesFuncionario_(data) {
  const funcionarioId = valor_(data.funcionarioId);
  if (!funcionarioId) throw new Error('Falta el funcionario.');
  if (valor_(data.confirmacion) !== 'ELIMINAR FUNCIONARIO') throw new Error('Confirmación de seguridad no válida.');

  const ss = getDb_();
  const persona = obtenerFuncionario_(ss, funcionarioId);
  if (!persona) throw new Error('No se encontró el funcionario.');
  asegurarSinSesionActiva_({funcionarioId:funcionarioId});

  const eliminados = filtrarResultados_(ss, function(row) {
    return String(row[2] || '') !== funcionarioId;
  });
  limpiarSesionesCompartidas_({funcionarioId:funcionarioId});
  return {eliminados:eliminados, funcionarioId:funcionarioId, nombre:persona.nombre || ''};
}

/** Elimina todas las evaluaciones históricas de un destacamento, sin borrar funcionarios. */
function eliminarEvaluacionesDestacamento_(data) {
  const destacamento = normalizarDependencia_(data.destacamento);
  const comando = valor_(data.comando) || comandoDependencia_(destacamento);
  if (!dependenciaValida_(comando, destacamento)) throw new Error('Comando o dependencia no válidos.');
  if (valor_(data.confirmacion) !== destacamento) throw new Error('Confirmación de seguridad no válida.');

  const ss = getDb_();
  const ids = idsFuncionariosDestacamento_(ss, destacamento, comando);
  const eliminados = filtrarResultados_(ss, function(row) {
    return !ids[String(row[2] || '')];
  });
  limpiarSesionesCompartidas_({comando:comando, destacamento:destacamento});
  return {eliminados:eliminados, comando:comando, destacamento:destacamento, funcionarios:Object.keys(ids).length};
}

/** Borra las fichas y el historial de una sola dependencia dentro de su comando. */
function eliminarRegistrosDestacamento_(data) {
  const destacamento = normalizarDependencia_(data.destacamento);
  const comando = valor_(data.comando);
  if (!dependenciaValida_(comando, destacamento)) throw new Error('Comando o dependencia no válidos.');
  if (valor_(data.confirmacion) !== 'ELIMINAR ' + comando + ' / ' + destacamento) {
    throw new Error('Confirmación de seguridad no válida.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    asegurarSinSesionActiva_({comando:comando, destacamento:destacamento});
    const ss = getDb_();
    const shF = ss.getSheetByName('Funcionarios');
    const shR = ss.getSheetByName('Resultados');
    const colsF = Math.max(30, shF.getLastColumn());
    const filasF = shF.getLastRow() > 1 ? shF.getRange(2, 1, shF.getLastRow() - 1, colsF).getValues() : [];
    const ids = {};
    const conservarF = filasF.filter(function(row) {
      const dep = normalizarDependencia_(row[8]);
      const cmd = String(row[28] || '') || comandoDependencia_(dep);
      if (dep !== destacamento || cmd !== comando) return true;
      ids[String(row[0] || '')] = true;
      return false;
    });
    const colsR = Math.max(10, shR.getLastColumn());
    const filasR = shR.getLastRow() > 1 ? shR.getRange(2, 1, shR.getLastRow() - 1, colsR).getValues() : [];
    const conservarR = filasR.filter(function(row) { return !ids[String(row[2] || '')]; });
    if (filasR.length) {
      shR.getRange(2, 1, filasR.length, colsR).clearContent();
      if (conservarR.length) shR.getRange(2, 1, conservarR.length, colsR).setValues(conservarR);
    }
    if (filasF.length) {
      shF.getRange(2, 1, filasF.length, colsF).clearContent();
      if (conservarF.length) shF.getRange(2, 1, conservarF.length, colsF).setValues(conservarF);
    }
    limpiarSesionesCompartidas_({comando:comando, destacamento:destacamento});
    PropertiesService.getScriptProperties().setProperty(resetDependenciaKey_(comando, destacamento), String(Date.now()));
    SpreadsheetApp.flush();
    return {comando:comando, destacamento:destacamento,
      funcionariosEliminados:filasF.length - conservarF.length,
      resultadosEliminados:filasR.length - conservarR.length};
  } finally { lock.releaseLock(); }
}

/** Reinicio general para pasar de pruebas a operación real. Conserva todas las fichas. */
function reiniciarTodasEvaluaciones_(data) {
  if (valor_(data.confirmacion) !== 'REINICIAR TODO') throw new Error('Confirmación de seguridad no válida.');
  const ss = getDb_();
  const eliminados = filtrarResultados_(ss, function() { return false; });
  limpiarSesionesCompartidas_({todas:true});
  PropertiesService.getScriptProperties().setProperty(RESULTADOS_RESET_PROPERTY, String(Date.now()));
  return {eliminados:eliminados, reiniciado:true};
}

/**
 * Reinicio absoluto: elimina funcionarios, evaluaciones y sesiones compartidas.
 * Conserva las hojas, encabezados y configuración de la aplicación.
 */
function eliminarTodosLosRegistros_(data) {
  if (valor_(data.confirmacion) !== 'ELIMINAR TODOS LOS REGISTROS') {
    throw new Error('Confirmación de seguridad no válida.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getDb_();
    const shF = ss.getSheetByName('Funcionarios');
    const shR = ss.getSheetByName('Resultados');
    const funcionariosEliminados = shF && shF.getLastRow() > 1 ? shF.getLastRow() - 1 : 0;
    const resultadosEliminados = shR && shR.getLastRow() > 1 ? shR.getLastRow() - 1 : 0;

    if (funcionariosEliminados) {
      shF.getRange(2, 1, funcionariosEliminados, Math.max(30, shF.getLastColumn())).clearContent();
    }
    if (resultadosEliminados) {
      shR.getRange(2, 1, resultadosEliminados, Math.max(10, shR.getLastColumn())).clearContent();
    }

    limpiarSesionesCompartidas_({todas:true});
    const ahora = String(Date.now());
    const props = PropertiesService.getScriptProperties();
    props.setProperty(RESULTADOS_RESET_PROPERTY, ahora);
    props.setProperty(REGISTROS_RESET_PROPERTY, ahora);
    SpreadsheetApp.flush();
    return {reiniciado:true, funcionariosEliminados:funcionariosEliminados, resultadosEliminados:resultadosEliminados};
  } finally {
    lock.releaseLock();
  }
}

/**
 * Conserva las filas para las que filtro(row) devuelve true y retorna cuántas eliminó.
 * Se reescribe el bloque bajo el encabezado para que la operación sea rápida y atómica.
 */
function filtrarResultados_(ss, filtro) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = ss.getSheetByName('Resultados');
    if (!sh || sh.getLastRow() <= 1) return 0;
    const lastRow = sh.getLastRow();
    const lastCol = Math.max(10, sh.getLastColumn());
    const rows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const conservar = rows.filter(filtro);
    sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    if (conservar.length) sh.getRange(2, 1, conservar.length, lastCol).setValues(conservar);
    SpreadsheetApp.flush();
    return rows.length - conservar.length;
  } finally {
    lock.releaseLock();
  }
}

function idsFuncionariosDestacamento_(ss, destacamento, comando) {
  const ids = {};
  const sh = ss.getSheetByName('Funcionarios');
  if (!sh || sh.getLastRow() <= 1) return ids;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 32).getValues();
  rows.forEach(function(row) {
    const dep = normalizarDependencia_(row[8]);
    const cmd = String(row[28] || '') || comandoDependencia_(dep);
    if (dep === destacamento && (!comando || cmd === comando)) ids[String(row[0] || '')] = true;
  });
  return ids;
}

function asegurarSinSesionActiva_(scope) {
  const props = PropertiesService.getScriptProperties().getProperties();
  Object.keys(props).forEach(function(key) {
    if (key.indexOf('SESION_COMPARTIDA_') !== 0) return;
    let ses;
    try { ses = JSON.parse(props[key]); } catch (err) { return; }
    if (!ses || ses.estado !== 'ACTIVA') return;
    const afecta = scope.funcionarioId
      ? !!(ses.funcionarios && ses.funcionarios[scope.funcionarioId])
      : scope.destacamento ? ses.destacamento === scope.destacamento && (!scope.comando || ses.comando === scope.comando) : true;
    if (afecta) throw new Error('Hay una sesión compartida activa. Finalizala antes de eliminar estos registros.');
  });
}

function limpiarSesionesCompartidas_(scope) {
  const service = PropertiesService.getScriptProperties();
  const props = service.getProperties();
  Object.keys(props).forEach(function(key) {
    if (key.indexOf('SESION_COMPARTIDA_') !== 0) return;
    let ses;
    try { ses = JSON.parse(props[key]); } catch (err) { ses = null; }
    const borrar = scope.todas ||
      (scope.destacamento && ses && ses.destacamento === scope.destacamento && (!scope.comando || ses.comando === scope.comando)) ||
      (scope.funcionarioId && ses && ses.funcionarios && ses.funcionarios[scope.funcionarioId]);
    if (borrar) service.deleteProperty(key);
  });
}



function guardarResultado_(data) {

  const funcionarioId = valor_(data.funcionarioId);

  const prueba = valor_(data.prueba);

  const resetMs = Number(PropertiesService.getScriptProperties().getProperty(RESULTADOS_RESET_PROPERTY) || 0);
  const clientCreatedMs = Number(data._clientCreatedMs || 0);
  if (resetMs && clientCreatedMs && clientCreatedMs < resetMs) {
    throw new Error('Este resultado pendiente pertenece a la etapa anterior al reinicio y no se guardó.');
  }



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

   NO APTO se aplica cuando la tabla oficial define un mínimo.

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



    // Fórmulas de la planilla: H 115/75/35; M 73/48/20.
    if (masculino) return { categoria: s >= 115 ? 'A' : (s >= 75 ? 'B' : (s >= 35 ? 'C' : 'NO APTO')), nota: '' };

    return { categoria: s >= 73 ? 'A' : (s >= 48 ? 'B' : (s >= 20 ? 'C' : 'NO APTO')), nota: '' };

  }



  if (prueba === 'Flexiones') {

    if (!masculino && !femenino) return { categoria: '', nota: '' };

    const n = Number(r.repeticiones);

    if (!isFinite(n)) return { categoria: '', nota: '' };



    // Fórmulas de la planilla: Excelente >=49 H / >=34 M.

    // Muy Bueno + Bueno = B. Malo = C.

    if (masculino) return { categoria: n >= 49 ? 'A' : (n >= 20 ? 'B' : 'C'), nota: '' };

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

    if (!isFinite(m) || !isFinite(e) || e < 18) return { categoria: '', nota: '' };



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

    if (nota === 0) return { categoria: 'C', nota: 0 };



    const categoria = nota >= 8 ? 'A' : (nota >= 5 ? 'B' : 'C');

    return { categoria: categoria, nota: nota };

  }



  return { categoria: '', nota: '' };

}





function obtenerFuncionario_(ss, funcionarioId) {

  const sh = ss.getSheetByName('Funcionarios');

  if (sh.getLastRow() <= 1) return null;



  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 32).getValues();

  for (let i = 0; i < rows.length; i++) {

    if (String(rows[i][0]) === String(funcionarioId)) {

      const partesNombre = String(rows[i][29] || '')
        ? {nombre:String(rows[i][3] || ''), apellido:String(rows[i][29] || '')}
        : separarNombre_(String(rows[i][3] || ''));
      return {

        id: String(rows[i][0] || ''),

        nombre: [partesNombre.nombre, partesNombre.apellido].filter(Boolean).join(' '),
        nombreSolo: partesNombre.nombre,
        apellido: partesNombre.apellido,

        nacimiento: fechaIso_(rows[i][5]),

        genero: String(rows[i][6] || ''),
        destacamento: normalizarDependencia_(rows[i][8]),
        comando: String(rows[i][28] || '') || comandoDependencia_(rows[i][8]),
        comandoEvaluacion: String(rows[i][30] || ''),
        destacamentoEvaluacion: normalizarDependencia_(rows[i][31])

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
  const comandoFiltro = valor_(data.comando);
  const destacamentoFiltro = normalizarDependencia_(data.destacamento);

  const ss = getDb_();

  const shF = ss.getSheetByName('Funcionarios');

  const shR = ss.getSheetByName('Resultados');



  const funcionarios = shF.getLastRow() > 1

    ? shF.getRange(2, 1, shF.getLastRow() - 1, 30).getValues()

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
      fecha: r[0],
      segundos: r[4], repeticiones: r[5], metros: r[6],

      edad: r[7], nota: r[8], categoria: String(r[9] || '')

    };

  });



  const funcionariosFiltrados = funcionarios.filter(function(f) {
    const dep = normalizarDependencia_(f[8]);
    const cmd = String(f[28] || '') || comandoDependencia_(dep);
    return (!destacamentoFiltro || dep === destacamentoFiltro) && (!comandoFiltro || cmd === comandoFiltro);
  });



  const rows = funcionariosFiltrados.map(f => {

    const id = String(f[0] || '');

    const r = porPersona[id] || {};
    const dependencia = normalizarDependencia_(f[8]);
    const comando = String(f[28] || '') || comandoDependencia_(dependencia);
    const nombre = String(f[29] || '')
      ? {nombre:String(f[3] || ''), apellido:String(f[29] || '')}
      : separarNombre_(String(f[3] || ''));
    const edad = edadResultado_(r, f[5]);
    const altura = Number(f[27]) > 3 ? Number(f[27]) / 100 : numeroOVacio_(f[27]);
    const peso = numeroOVacio_(f[26]);
    const altura2 = altura === '' ? '' : Math.round(altura * altura * 10000) / 10000;
    const imc = peso !== '' && altura2 ? Math.round((peso / altura2) * 100) / 100 : '';
    const fechaPrueba = fechaUltimoResultado_(r);

    return [
      fechaPlanilla_(fechaPrueba), comando, dependencia, String(f[2] || ''), nombre.nombre, nombre.apellido,
      String(f[4] || ''), String(f[7] || ''), edad, fechaPlanilla_(f[5]), fechaPlanilla_(f[9]),
      f[9] ? calcularEdad_(f[9]) : '', peso, altura, altura2, imc, clasificarImc_(imc), '', '', String(f[6] || ''),
      valorResultado_(r, 'Core', 'segundos'), valorResultado_(r, 'Core', 'categoria'),
      valorResultado_(r, 'Flexibilidad', 'categoria'),
      valorResultado_(r, 'Flexiones', 'repeticiones'), valorResultado_(r, 'Flexiones', 'categoria'),
      valorResultado_(r, 'Sentadilla', 'segundos'), valorResultado_(r, 'Sentadilla', 'categoria'),
      valorResultado_(r, 'Yo-Yo', 'metros'), valorResultado_(r, 'Yo-Yo', 'categoria'),
      comorbilidadFuncionario_(f), estadoDocumento_(f[10]), estadoDocumento_(f[12])
    ];

  }).sort(function(a, b) {
    return String(a[1]).localeCompare(String(b[1]), 'es') ||
      String(a[2]).localeCompare(String(b[2]), 'es') ||
      String(a[5]).localeCompare(String(b[5]), 'es') ||
      String(a[4]).localeCompare(String(b[4]), 'es');
  });



  const etiqueta = destacamentoFiltro

    ? destacamentoFiltro.replace(/[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ]+/g, '_')

    : 'GENERAL';

  const nombre = 'Evaluaciones_Fisicas_' + anio + '_' + etiqueta + '.xlsx';
  const blob = crearXlsxResultados_(rows, nombre, anio);



  return {

    nombre: nombre,

    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',

    base64: Utilities.base64Encode(blob.getBytes()),

    filas: rows.length,

    anio: anio,

    comando: comandoFiltro || 'TODOS',
    destacamento: destacamentoFiltro || 'GENERAL'

  };

}

function valorResultado_(resultados, prueba, campo) {
  return resultados[prueba] && resultados[prueba][campo] !== undefined ? resultados[prueba][campo] : '';
}

function edadResultado_(resultados, nacimiento) {
  const pruebas = Object.keys(resultados);
  for (let i = 0; i < pruebas.length; i++) {
    const edad = resultados[pruebas[i]].edad;
    if (edad !== '' && edad !== undefined) return Number(edad);
  }
  return calcularEdad_(nacimiento);
}

function fechaUltimoResultado_(resultados) {
  let ultima = null;
  Object.keys(resultados).forEach(function(prueba) {
    const fecha = resultados[prueba].fecha;
    if (!fecha) return;
    const d = fecha instanceof Date ? fecha : new Date(fecha);
    if (!isNaN(d.getTime()) && (!ultima || d > ultima)) ultima = d;
  });
  return ultima;
}

function separarNombre_(texto) {
  const partes = valor_(texto).split(/\s+/).filter(Boolean);
  if (partes.length < 2) return {nombre:partes[0] || '', apellido:''};
  return {nombre:partes.slice(0, -1).join(' '), apellido:partes[partes.length - 1]};
}

function fechaPlanilla_(valor) {
  if (!valor) return '';
  const d = valor instanceof Date ? valor : new Date(valor);
  if (isNaN(d.getTime())) return String(valor);
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'America/Montevideo', 'dd/MM/yyyy');
}

function clasificarImc_(imc) {
  const n = Number(imc);
  if (!isFinite(n) || n <= 0) return '';
  if (n < 18.5) return 'Bajo peso';
  if (n < 25) return 'Normal';
  if (n < 30) return 'Sobrepeso';
  if (n < 35) return 'Obesidad I';
  if (n < 40) return 'Obesidad II';
  return 'Obesidad III';
}

function estadoDocumento_(valor) {
  const t = valor_(valor).toLowerCase();
  if (t === 'sí' || t === 'si' || t === 'vigente') return 'VIGENTE';
  if (t === 'no' || t === 'no vigente') return 'NO VIGENTE';
  return valor_(valor);
}

function comorbilidadFuncionario_(f) {
  const datos = [];
  if (/^(sí|si)$/i.test(valor_(f[16]))) datos.push(valor_(f[17]) || 'Enfermedad crónica');
  if (/^(sí|si)$/i.test(valor_(f[18]))) datos.push(valor_(f[19]) || 'Antecedente de enfermedad');
  if (/^(sí|si)$/i.test(valor_(f[20]))) datos.push(valor_(f[21]) || 'Lesión');
  return datos.length ? datos.join(' · ') : 'No presenta';
}

function columnaExcel_(numero) {
  let n = numero, texto = '';
  while (n > 0) { n--; texto = String.fromCharCode(65 + (n % 26)) + texto; n = Math.floor(n / 26); }
  return texto;
}

function xmlEscape_(valor) {
  return String(valor == null ? '' : valor).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
}

function celdaXlsx_(fila, columna, valor, estilo) {
  const ref = columnaExcel_(columna) + fila;
  const s = estilo ? ' s="' + estilo + '"' : '';
  if (valor === '' || valor === null || valor === undefined) return '<c r="' + ref + '"' + s + '/>';
  if (typeof valor === 'number' && isFinite(valor)) return '<c r="' + ref + '"' + s + '><v>' + valor + '</v></c>';
  return '<c r="' + ref + '" t="inlineStr"' + s + '><is><t xml:space="preserve">' + xmlEscape_(valor) + '</t></is></c>';
}

function estiloCategoriaXlsx_(valor) {
  const t = valor_(valor).toUpperCase();
  return t === 'A' ? 5 : (t === 'B' ? 6 : (t === 'C' ? 7 : (t === 'NO APTO' ? 8 : 4)));
}

function crearXlsxResultados_(rows, nombre, anio) {
  // Orden A–AF idéntico a la hoja TABLA RESULTADOS de la planilla oficial.
  const headers = ['Fecha Prueba','Comando','Dependencia','Grado','Nombre','Apellido','Cédula','CELULAR','Edad de evaluación','F. Nacimiento','Fecha Ingreso','Años de servicio','Peso','Altura (m)','Altura²','IMC','Clasificación IMC','REPOSO','Máxima','SEXO','CORE (seg.)','','Flexibilidad / OFD','Flexiones (MMSS)','Categoría MMSS','Sentadilla isométrica (MMII)','Categoría MMII','Yo-Yo (m)','Categoría Yo-Yo','Comorbilidad','Carné de salud','Ergometría'];
  const grupos = {17:'ESTADO FÍSICO',18:'FRECUENCIA CARDIACA',21:'CORE',23:'Flexibilidad',24:'R. Muscular (MMSS)',26:'R. Muscular (MMII)',28:'CAPACIDAD AERÓBICA'};
  const categorias = {22:true,23:true,25:true,27:true,29:true};
  let sheetRows = '<row r="1" ht="30" customHeight="1">' + celdaXlsx_(1,1,'',1) + '</row>';
  let grupoCeldas = '';
  for (let c = 1; c <= 32; c++) grupoCeldas += celdaXlsx_(2,c,grupos[c] || '',2);
  sheetRows += '<row r="2" ht="25" customHeight="1">' + grupoCeldas + '</row>';
  let headerCeldas = '';
  headers.forEach(function(h, i) { headerCeldas += celdaXlsx_(3,i + 1,h,3); });
  sheetRows += '<row r="3" ht="42" customHeight="1">' + headerCeldas + '</row>';
  rows.forEach(function(row, i) {
    const numeroFila = i + 4;
    let celdas = '';
    row.forEach(function(valor, j) { celdas += celdaXlsx_(numeroFila,j + 1,valor,categorias[j + 1] ? estiloCategoriaXlsx_(valor) : 4); });
    sheetRows += '<row r="' + numeroFila + '">' + celdas + '</row>';
  });
  const lastRow = Math.max(3, rows.length + 3);
  const widths = [13,12,25,14,22,18,13,14,18,14,14,16,10,12,11,10,18,11,11,12,13,17,18,18,17,24,17,13,17,28,18,16];
  let cols = '<cols>';
  widths.forEach(function(w,i){ cols += '<col min="' + (i+1) + '" max="' + (i+1) + '" width="' + w + '" customWidth="1"/>'; });
  cols += '</cols>';
  const sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="18"/>' + cols + '<sheetData>' + sheetRows + '</sheetData>' +
    '<autoFilter ref="A3:AF' + lastRow + '"/><mergeCells count="6"><mergeCell ref="A1:AF1"/><mergeCell ref="R2:S2"/><mergeCell ref="U2:V2"/><mergeCell ref="X2:Y2"/><mergeCell ref="Z2:AA2"/><mergeCell ref="AB2:AC2"/></mergeCells></worksheet>';
  const stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="16"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Calibri"/></font></fonts>' +
    '<fills count="8"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF7A1823"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF202020"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFE699"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill></fills>' +
    '<borders count="2"><border/><border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="9">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0"/><xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0"/><xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0"/><xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0"/></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  const files = [
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>','application/xml','[Content_Types].xml'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>','application/xml','_rels/.rels'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="TABLA RESULTADOS" sheetId="1" r:id="rId1"/></sheets></workbook>','application/xml','xl/workbook.xml'),
    Utilities.newBlob('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>','application/xml','xl/_rels/workbook.xml.rels'),
    Utilities.newBlob(stylesXml,'application/xml','xl/styles.xml'),
    Utilities.newBlob(sheetXml,'application/xml','xl/worksheets/sheet1.xml')
  ];
  return Utilities.zip(files, nombre).setContentType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
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
