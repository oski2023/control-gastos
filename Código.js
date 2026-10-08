const SS = SpreadsheetApp.getActiveSpreadsheet();

function doGet(e) {
  const isDemo = (e && e.parameter && (e.parameter.demo === '1' || e.parameter.demo === 'true' || e.parameter.modo === 'demo')) ? true : false;
  const template = HtmlService.createTemplateFromFile('Index');
  template.isDemoParam = isDemo;
  return template.evaluate()
    .setTitle('Control de Gastos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function verificarEstadoDemo() {
  try {
    const sh = getSheet_('Configuración');
    const celdaTitulo = String(sh.getRange('E1').getValue() || '').trim();
    const celdaValor = String(sh.getRange('E2').getValue() || '').trim().toUpperCase();

    if (!celdaTitulo) {
      sh.getRange('E1').setValue('Acceso Demo');
      sh.getRange('E2').setValue('ACTIVO');
      return { activo: true };
    }

    if (celdaValor === 'BLOQUEADO' || celdaValor === 'NO' || celdaValor === 'PAUSADO' || celdaValor === 'INACTIVO') {
      return { activo: false };
    }

    return { activo: true };
  } catch(e) {
    return { activo: true };
  }
}

function getSheet_(name) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    const allSheets = SS.getSheets();
    const target = name.toLowerCase().trim();
    for (let i = 0; i < allSheets.length; i++) {
      const s = allSheets[i];
      const sName = s.getName().toLowerCase().trim();
      if (sName === target || sName === target.replace(/s$/, '') || (sName + 's') === target) {
        return s;
      }
    }
    if (target === 'tarjetas' || target === 'tarjeta') {
      sh = SS.insertSheet('Tarjetas');
      sh.appendRow(['Fecha', 'Descripción', 'Entidad', 'Categoría', 'Importe Total', 'Cuotas', 'Cuota Mensual', 'Primera Cuota', 'Fecha de Registro']);
      return sh;
    }
    if (target === 'presupuestos' || target === 'presupuesto') {
      sh = SS.insertSheet('Presupuestos');
      sh.appendRow(['Categoría', 'Monto Presupuesto']);
      return sh;
    }
    if (target === 'metasahorro' || target === 'metas' || target === 'meta') {
      sh = SS.insertSheet('MetasAhorro');
      sh.appendRow(['ID', 'Nombre', 'Monto Objetivo', 'Monto Actual', 'Fecha Límite', 'Color']);
      return sh;
    }
    if (target === 'gastosfijos' || target === 'gastofijo') {
      sh = SS.insertSheet('GastosFijos');
      sh.appendRow(['Categoría', 'Día Vencimiento', 'Activo', 'Es Préstamo', 'Cuota Base', 'Total Cuotas', 'Mes Base', 'Año Base', 'Monto Estimado']);
      return sh;
    }
    throw new Error('No existe la hoja: ' + name);
  }
  return sh;
}

function parseFechaLocal_(fechaStr) {
  if (!fechaStr) return new Date();
  if (fechaStr instanceof Date) return fechaStr;
  const partes = String(fechaStr).split('-');
  if (partes.length < 3) return new Date();
  return new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));
}

function getConfig() {
  const sh = getSheet_('Configuración');
  const last = Math.max(sh.getLastRow(), 2);
  const values = sh.getRange(2, 1, last - 1, 3).getValues();
  const limpiar = columna => columna
    .map(v => String(v || '').trim())
    .filter(v => v && !/agreg[aá].*(debajo|lista|aquí)/i.test(v));

  const ordenarAZ = lista => lista.sort((a, b) => String(a).localeCompare(String(b), 'es', { sensitivity: 'base', numeric: true }));

  const cats = limpiar(values.map(r => r[0]));
  if (!cats.some(c => c.toLowerCase() === 'peajes' || c.toLowerCase() === 'peaje')) {
    cats.push('Peajes');
  }
  const ents = limpiar(values.map(r => r[2]));
  if (!ents.some(e => e.toLowerCase().includes('rio') || e.toLowerCase().includes('santander'))) {
    ents.push('Visa Banco Rio');
  }

  return {
    categorias: ordenarAZ(cats),
    medios: ordenarAZ(limpiar(values.map(r => r[1]))),
    entidades: ordenarAZ(ents)
  };
}

function addConfigItem(tipo, nombre) {
  nombre = String(nombre || '').trim();

  if (!nombre) {
    throw new Error('El nombre no puede estar vacío.');
  }

  const columnas = {
    categoria: 1,
    medio: 2,
    entidad: 3
  };

  const columna = columnas[tipo];

  if (!columna) {
    throw new Error('Tipo de configuración no válido.');
  }

  const sh = getSheet_('Configuración');
  const ultimaFila = Math.max(sh.getLastRow(), 2);

  const valores = sh
    .getRange(2, columna, Math.max(ultimaFila - 1, 1), 1)
    .getValues()
    .map(r => String(r[0] || '').trim().toLowerCase());

  if (valores.includes(nombre.toLowerCase())) {
    throw new Error('Ese elemento ya existe.');
  }

  let fila = 2;

  while (
    fila <= sh.getMaxRows() &&
    String(sh.getRange(fila, columna).getValue() || '').trim() !== ''
  ) {
    fila++;
  }

  if (fila > sh.getMaxRows()) {
    sh.insertRowAfter(sh.getMaxRows());
  }

  sh.getRange(fila, columna).setValue(nombre);

  return getConfig();
}

function removeConfigItem(tipo, nombre) {
  nombre = String(nombre || '').trim().toLowerCase();
  const columnas = { categoria: 1, medio: 2, entidad: 3 };
  const col = columnas[tipo];
  if (!col) throw new Error('Tipo de configuración no válido.');

  const sh = getSheet_('Configuración');
  const last = Math.max(sh.getLastRow(), 2);
  const vals = sh.getRange(2, col, Math.max(last - 1, 1), 1).getValues();

  const filtrados = vals
    .map(r => String(r[0] || '').trim())
    .filter(v => v && v.toLowerCase() !== nombre && !/agreg[aá].*(debajo|lista|aquí)/i.test(v));

  // Limpiar columna desde la fila 2
  sh.getRange(2, col, Math.max(last - 1, 1), 1).clearContent();

  if (filtrados.length > 0) {
    const nuevosVals = filtrados.map(v => [v]);
    sh.getRange(2, col, nuevosVals.length, 1).setValues(nuevosVals);
  }

  return getConfig();
}

function asegurarColumnasGastosFijos_(sh) {
  if (!sh) return;
  const headers = ['Categoría', 'Día Vencimiento', 'Activo', 'Es Préstamo', 'Cuota Base', 'Total Cuotas', 'Mes Base', 'Año Base', 'Monto Estimado'];
  const maxCols = sh.getMaxColumns();
  if (maxCols < headers.length) {
    sh.insertColumnsAfter(maxCols, headers.length - maxCols);
  }
  const lastCol = Math.max(sh.getLastColumn(), headers.length);
  const fila1 = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  let cambioHeaders = false;
  for (let c = 0; c < headers.length; c++) {
    if (!fila1[c]) {
      sh.getRange(1, c + 1).setValue(headers[c]);
      cambioHeaders = true;
    }
  }
}

function getGastosFijosConfig(mes, anio) {
  const sh = getSheet_('GastosFijos');
  asegurarColumnasGastosFijos_(sh);
  const last = sh.getLastRow();
  if (last < 2) return [];

  const hoy = new Date();
  const mesConsulta = mes ? Number(mes) : (hoy.getMonth() + 1);
  const anioConsulta = anio ? Number(anio) : hoy.getFullYear();

  const numCols = Math.max(sh.getLastColumn(), 9);
  const vals = sh.getRange(2, 1, last - 1, numCols).getValues();
  const lista = [];

  for (let i = 0; i < vals.length; i++) {
    const cat = String(vals[i][0] || '').trim();
    if (!cat) continue;
    const dia = Math.max(1, Math.min(31, Number(vals[i][1]) || 1));
    const activoVal = String(vals[i][2] || 'SI').trim().toUpperCase();
    const activo = (activoVal !== 'NO' && activoVal !== 'PAUSADO' && activoVal !== 'INACTIVO');

    const esPrestamoVal = String(vals[i][3] || '').trim().toUpperCase();
    const cuotaBase = Number(vals[i][4]) || 0;
    const totalCuotas = Number(vals[i][5]) || 0;
    const esPrestamo = (esPrestamoVal === 'SI' || esPrestamoVal === 'TRUE' || totalCuotas > 0);

    const mesBase = Number(vals[i][6]) || mesConsulta;
    const anioBase = Number(vals[i][7]) || anioConsulta;
    const montoEstimado = Number(vals[i][8]) || 0;

    let cuotaActual = 0;
    let cuotaTexto = '';
    let finalizado = false;
    let noIniciado = false;
    let esUltimaCuota = false;
    let cuotasRestantes = 0;
    let porcentaje = 0;

    if (esPrestamo && totalCuotas > 0) {
      const diffMeses = (anioConsulta - anioBase) * 12 + (mesConsulta - mesBase);
      cuotaActual = (cuotaBase || 1) + diffMeses;
      finalizado = cuotaActual > totalCuotas;
      noIniciado = cuotaActual < 1;
      esUltimaCuota = (cuotaActual === totalCuotas);
      cuotasRestantes = Math.max(0, totalCuotas - cuotaActual);
      cuotaTexto = 'Cuota ' + Math.max(1, Math.min(cuotaActual, totalCuotas)) + '/' + totalCuotas;
      porcentaje = Math.min(100, Math.max(0, Math.round((Math.max(0, cuotaActual) / totalCuotas) * 100)));
    }

    lista.push({
      fila: i + 2,
      categoria: cat,
      diaVencimiento: dia,
      activo: activo,
      esPrestamo: esPrestamo,
      cuotaBase: cuotaBase,
      totalCuotas: totalCuotas,
      mesBase: mesBase,
      anioBase: anioBase,
      montoEstimado: montoEstimado,
      cuotaActual: cuotaActual,
      cuotaTexto: cuotaTexto,
      finalizado: finalizado,
      noIniciado: noIniciado,
      esUltimaCuota: esUltimaCuota,
      cuotasRestantes: cuotasRestantes,
      porcentaje: porcentaje
    });
  }
  return lista.sort((a, b) => a.categoria.localeCompare(b.categoria, 'es', { sensitivity: 'base' }));
}

function addGastoFijo(param1, diaVencimiento, esPrestamo, cuotaBase, totalCuotas, mesBase, anioBase, montoEstimado) {
  let cat = '';
  let dia = 10;
  let esPrest = false;
  let cBase = 1;
  let tCuotas = 0;
  let mBase = 0;
  let aBase = 0;
  let mEstimado = 0;

  if (param1 && typeof param1 === 'object') {
    cat = String(param1.categoria || '').trim();
    dia = Number(param1.diaVencimiento || param1.dia || 10);
    esPrest = !!param1.esPrestamo;
    cBase = Number(param1.cuotaBase || param1.cuotaActual || 1) || 1;
    tCuotas = Number(param1.totalCuotas || 0) || 0;
    mBase = Number(param1.mesBase || 0);
    aBase = Number(param1.anioBase || 0);
    mEstimado = Number(param1.montoEstimado || param1.monto || 0);
  } else {
    cat = String(param1 || '').trim();
    dia = Number(diaVencimiento) || 10;
    esPrest = !!esPrestamo;
    cBase = Number(cuotaBase) || 1;
    tCuotas = Number(totalCuotas) || 0;
    mBase = Number(mesBase) || 0;
    aBase = Number(anioBase) || 0;
    mEstimado = Number(montoEstimado) || 0;
  }

  if (!cat) throw new Error('Ingresá el nombre o categoría del gasto fijo.');
  dia = Math.max(1, Math.min(31, dia || 1));

  const hoy = new Date();
  if (!mBase) mBase = hoy.getMonth() + 1;
  if (!aBase) aBase = hoy.getFullYear();

  const sh = getSheet_('GastosFijos');
  asegurarColumnasGastosFijos_(sh);

  const tieneCuotas = esPrest || tCuotas > 0;
  sh.appendRow([
    cat,
    dia,
    'SI',
    tieneCuotas ? 'SI' : 'NO',
    tieneCuotas ? cBase : '',
    tieneCuotas ? tCuotas : '',
    tieneCuotas ? mBase : '',
    tieneCuotas ? aBase : '',
    mEstimado || ''
  ]);

  return getGastosFijosConfig(mBase, aBase);
}

function editarGastoFijo(fila, data) {
  fila = Number(fila);
  const sh = getSheet_('GastosFijos');
  asegurarColumnasGastosFijos_(sh);
  if (fila < 2 || fila > sh.getLastRow()) {
    throw new Error('Fila de gasto fijo no válida.');
  }

  const cat = String(data.categoria || '').trim();
  if (!cat) throw new Error('El nombre o categoría no puede estar vacío.');
  const dia = Math.max(1, Math.min(31, Number(data.diaVencimiento || 10)));
  const esPrest = !!data.esPrestamo || Number(data.totalCuotas) > 0;
  const cBase = Number(data.cuotaBase || data.cuotaActual || 1) || 1;
  const tCuotas = Number(data.totalCuotas || 0) || 0;
  const hoy = new Date();
  const mBase = Number(data.mesBase) || (hoy.getMonth() + 1);
  const aBase = Number(data.anioBase) || hoy.getFullYear();
  const mEstimado = Number(data.montoEstimado || data.monto || 0);

  const estadoActual = String(sh.getRange(fila, 3).getValue() || 'SI').trim().toUpperCase();
  const activo = (estadoActual !== 'NO' && estadoActual !== 'PAUSADO' && estadoActual !== 'INACTIVO') ? 'SI' : 'NO';

  sh.getRange(fila, 1, 1, 9).setValues([[
    cat,
    dia,
    activo,
    esPrest ? 'SI' : 'NO',
    esPrest ? cBase : '',
    esPrest ? tCuotas : '',
    esPrest ? mBase : '',
    esPrest ? aBase : '',
    mEstimado || ''
  ]]);

  return getGastosFijosConfig(mBase, aBase);
}

function toggleGastoFijo(fila, activo) {
  const sh = getSheet_('GastosFijos');
  fila = Number(fila);
  if (fila >= 2 && fila <= sh.getLastRow()) {
    sh.getRange(fila, 3).setValue(activo ? 'SI' : 'NO');
  }
  return getGastosFijosConfig();
}

function eliminarGastoFijo(fila, categoriaVerif) {
  const sh = getSheet_('GastosFijos');
  fila = Number(fila);
  let filaBorrada = false;

  if (fila >= 2 && fila <= sh.getLastRow()) {
    if (categoriaVerif) {
      const catHoja = String(sh.getRange(fila, 1).getValue() || '').trim().toLowerCase();
      if (catHoja === String(categoriaVerif).trim().toLowerCase()) {
        sh.deleteRow(fila);
        filaBorrada = true;
      }
    } else {
      sh.deleteRow(fila);
      filaBorrada = true;
    }
  }

  if (!filaBorrada && categoriaVerif) {
    const last = sh.getLastRow();
    if (last >= 2) {
      const vals = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][0] || '').trim().toLowerCase() === String(categoriaVerif).trim().toLowerCase()) {
          sh.deleteRow(i + 2);
          break;
        }
      }
    }
  }

  return getGastosFijosConfig();
}

function addIngreso(data) {
  const sh = getSheet_('Ingresos');
  const f = parseFechaLocal_(data.fecha);
  sh.appendRow([
    f,
    data.descripcion || '',
    data.tipo || 'Ingreso',
    Number(data.importe) || 0,
    data.observaciones || '',
    new Date()
  ]);
  const mes = data.mesSeleccionado ? Number(data.mesSeleccionado) : (f.getMonth() + 1);
  const anio = data.anioSeleccionado ? Number(data.anioSeleccionado) : f.getFullYear();
  return getResumen(mes, anio);
}

function addGasto(data) {
  const sh = getSheet_('Gastos');
  const f = parseFechaLocal_(data.fecha);
  sh.appendRow([
    f,
    data.descripcion || '',
    data.categoria || '',
    data.entidad || '',
    data.medio || '',
    Number(data.importe) || 0,
    new Date()
  ]);
  const mes = data.mesSeleccionado ? Number(data.mesSeleccionado) : (f.getMonth() + 1);
  const anio = data.anioSeleccionado ? Number(data.anioSeleccionado) : f.getFullYear();
  return getResumen(mes, anio);
}

function getGastosFijosPendientes(mes, anio) {
  const shFijos = getSheet_('GastosFijos');
  asegurarColumnasGastosFijos_(shFijos);
  const fijos = shFijos.getDataRange().getValues();
  const gastos = getSheet_('Gastos').getDataRange().getValues();
  const tarjetas = getSheet_('Tarjetas').getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  const hoy = new Date();
  const mesActual = mes ? Number(mes) : (hoy.getMonth() + 1);
  const anioActual = anio ? Number(anio) : hoy.getFullYear();
  const esMesActual = (anioActual === hoy.getFullYear() && mesActual === hoy.getMonth() + 1);
  const diaActual = esMesActual ? hoy.getDate() : 31;

  let mesAnt = mesActual - 1;
  let anioAnt = anioActual;
  if (mesAnt === 0) {
    mesAnt = 12;
    anioAnt = anioActual - 1;
  }

  const nombresMeses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const getNombreMes = m => nombresMeses[(Number(m) - 1 + 12) % 12];
  const normCat = s => String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');

  const parseFechaFila = val => {
    if (val instanceof Date) return val;
    if (!val) return null;
    if (typeof val === 'string') {
      if (/^\d{4}-\d{2}-\d{2}/.test(val)) {
        const p = val.split('-');
        return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2].substring(0, 2)));
      }
      if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(val)) {
        const p = val.split('/');
        return new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
      }
      const d = new Date(val);
      return isNaN(d.getTime()) ? null : d;
    }
    return null;
  };

  // Contamos la cantidad y detalle de pagos por categoría para mes anterior y actual
  const pagosMesAnt = {};
  const pagosMesActual = {};
  const pagosInfoMesActual = {};

  const registrarPago = (fRaw, catRaw, importe, medio, desc) => {
    if (!catRaw) return;
    const f = parseFechaFila(fRaw);
    if (!f) return;
    const cat = normCat(catRaw);
    const a = f.getFullYear();
    const m = f.getMonth() + 1;

    if (a === anioAnt && m === mesAnt) {
      pagosMesAnt[cat] = (pagosMesAnt[cat] || 0) + 1;
    } else if (a === anioActual && m === mesActual) {
      pagosMesActual[cat] = (pagosMesActual[cat] || 0) + 1;
      if (!pagosInfoMesActual[cat]) pagosInfoMesActual[cat] = [];
      pagosInfoMesActual[cat].push({
        fecha: f,
        importe: Number(importe) || 0,
        medio: String(medio || ''),
        descripcion: String(desc || ''),
        usado: false
      });
    }
  };

  // Revisar pagos en Gastos (columna 0: fecha, 1: desc, 2: categoría, 4: medio, 5: importe)
  for (let i = 1; i < gastos.length; i++) {
    registrarPago(gastos[i][0], gastos[i][2], gastos[i][5], gastos[i][4], gastos[i][1]);
  }

  // Revisar pagos en Tarjetas (columna 0: fecha, 1: desc, 3: categoría, 4: importe)
  for (let i = 1; i < tarjetas.length; i++) {
    registrarPago(tarjetas[i][0], tarjetas[i][3], tarjetas[i][4], 'Tarjeta', tarjetas[i][1]);
  }

  const pendientes = [];
  const alDia = [];
  const proximos = [];

  for (let i = 1; i < fijos.length; i++) {
    const categoria = String(fijos[i][0] || '').trim();
    const diaVencimiento = Number(fijos[i][1]) || 0;
    const activoVal = String(fijos[i][2] || 'SI').trim().toUpperCase();
    const activo = (activoVal !== 'NO' && activoVal !== 'PAUSADO' && activoVal !== 'INACTIVO');
    if (!categoria || !activo) continue;

    // Campos de préstamo / cuotas
    const esPrestamoVal = String(fijos[i][3] || '').trim().toUpperCase();
    const cuotaBase = Number(fijos[i][4]) || 0;
    const totalCuotas = Number(fijos[i][5]) || 0;
    const esPrestamo = (esPrestamoVal === 'SI' || esPrestamoVal === 'TRUE' || totalCuotas > 0);
    const mesBase = Number(fijos[i][6]) || mesActual;
    const anioBase = Number(fijos[i][7]) || anioActual;
    const montoEstimado = Number(fijos[i][8]) || 0;

    let cuotaActual = 0;
    let cuotaTexto = '';
    let finalizado = false;
    let esUltimaCuota = false;
    let cuotasRestantes = 0;

    if (esPrestamo && totalCuotas > 0) {
      const diffMeses = (anioActual - anioBase) * 12 + (mesActual - mesBase);
      cuotaActual = (cuotaBase || 1) + diffMeses;
      finalizado = cuotaActual > totalCuotas;
      esUltimaCuota = (cuotaActual === totalCuotas);
      cuotasRestantes = Math.max(0, totalCuotas - cuotaActual);
      cuotaTexto = 'Cuota ' + cuotaActual + '/' + totalCuotas;

      // Si el préstamo ya terminó antes de este mes, no se procesa como deuda pendiente
      if (finalizado) {
        continue;
      }
    }

    const catKey = normCat(categoria);
    const listaPagosEsteMes = pagosInfoMesActual[catKey] || [];

    // Buscar si hay un pago no usado que corresponda a este gasto fijo / préstamo
    let pagoMatcheado = null;

    if (esPrestamo && totalCuotas > 0) {
      for (let pIdx = 0; pIdx < listaPagosEsteMes.length; pIdx++) {
        const pItem = listaPagosEsteMes[pIdx];
        if (pItem.usado) continue;
        const descNorm = normCat(pItem.descripcion);
        if (descNorm.includes(String(cuotaActual)) || descNorm.includes(String(totalCuotas)) || descNorm.includes(catKey)) {
          pagoMatcheado = pItem;
          pItem.usado = true;
          break;
        }
      }
    }

    if (!pagoMatcheado) {
      for (let pIdx = 0; pIdx < listaPagosEsteMes.length; pIdx++) {
        const pItem = listaPagosEsteMes[pIdx];
        if (!pItem.usado) {
          pagoMatcheado = pItem;
          pItem.usado = true;
          break;
        }
      }
    }

    // 1. ¿Tiene pago registrado este mes?
    if (pagoMatcheado) {
      alDia.push({
        categoria,
        diaVencimiento,
        monto: pagoMatcheado.importe || 0,
        fechaPago: pagoMatcheado.fecha ? Utilities.formatDate(pagoMatcheado.fecha, tz, 'dd/MM/yyyy') : '',
        medio: pagoMatcheado.medio || '',
        cantidadPagos: 1,
        esPrestamo,
        cuotaActual,
        totalCuotas,
        cuotaTexto,
        cuotasRestantes,
        esUltimaCuota,
        montoEstimado
      });
    } else {
      // No se ha pagado aún este mes
      // 1.a. Si tampoco se pagó el mes anterior, arrastra la deuda del mes anterior
      let pagosAnt = pagosMesAnt[catKey] || 0;
      if (pagosAnt === 0) {
        const fechaVtoAnt = new Date(anioAnt, mesAnt - 1, diaVencimiento || 1);
        const diffMs = hoy.getTime() - fechaVtoAnt.getTime();
        const diasVencidoAnt = Math.max(1, Math.floor(diffMs / (1000 * 60 * 60 * 24)));
        const cuotaAnt = esPrestamo ? Math.max(1, cuotaActual - 1) : 0;
        const cuotaTextoAnt = esPrestamo ? ('Cuota ' + cuotaAnt + '/' + totalCuotas) : '';

        pendientes.push({
          categoria,
          diaVencimiento,
          mesNombre: getNombreMes(mesAnt),
          mesNum: mesAnt,
          anio: anioAnt,
          periodo: 'anterior',
          esMesAnterior: true,
          diasVencido: diasVencidoAnt,
          esPrestamo,
          cuotaActual: cuotaAnt,
          totalCuotas,
          cuotaTexto: cuotaTextoAnt,
          cuotasRestantes: esPrestamo ? Math.max(0, totalCuotas - cuotaAnt) : 0,
          esUltimaCuota: esPrestamo && (cuotaAnt === totalCuotas),
          montoEstimado
        });
      }

      // 1.b. Pendiente o próximo en el mes actual
      if (diaActual >= diaVencimiento) {
        pendientes.push({
          categoria,
          diaVencimiento,
          mesNombre: getNombreMes(mesActual),
          mesNum: mesActual,
          anio: anioActual,
          periodo: 'actual',
          esMesAnterior: false,
          diasVencido: diaActual - diaVencimiento,
          esPrestamo,
          cuotaActual,
          totalCuotas,
          cuotaTexto,
          cuotasRestantes,
          esUltimaCuota,
          montoEstimado
        });
      } else {
        proximos.push({
          categoria,
          diaVencimiento,
          diasFaltan: diaVencimiento - diaActual,
          esPrestamo,
          cuotaActual,
          totalCuotas,
          cuotaTexto,
          cuotasRestantes,
          esUltimaCuota,
          montoEstimado
        });
      }
    }
  }

  // Ordenar pendientes: mes anterior primero, luego por días vencidos descendente
  pendientes.sort((a, b) => {
    if (a.esMesAnterior !== b.esMesAnterior) {
      return a.esMesAnterior ? -1 : 1;
    }
    return b.diasVencido - a.diasVencido;
  });

  // Ordenar al día por categoría
  alDia.sort((a, b) => a.categoria.localeCompare(b.categoria));

  // Ordenar próximos por día de vencimiento ascendente
  proximos.sort((a, b) => a.diaVencimiento - b.diaVencimiento);

  return {
    pendientes,
    alDia,
    proximos,
    totalFijos: fijos.length - 1
  };
}

function addTarjeta(data) {
  const sh = getSheet_('Tarjetas');
  const importe = Number(data.importe) || 0;
  const cuotas = Number(data.cuotas) || 1;
  const f = parseFechaLocal_(data.fecha);
  sh.appendRow([
    f,
    data.descripcion || '',
    data.entidad || '',
    data.categoria || '',
    importe,
    cuotas,
    cuotas ? importe / cuotas : importe,
    data.primeraCuota ? parseFechaLocal_(data.primeraCuota + '-01') : '',
    new Date()
  ]);
  const mes = data.mesSeleccionado ? Number(data.mesSeleccionado) : (f.getMonth() + 1);
  const anio = data.anioSeleccionado ? Number(data.anioSeleccionado) : f.getFullYear();
  return getResumen(mes, anio);
}

function getResumen(mes, anio) {
  const hoy = new Date();
  mes = mes ? Number(mes) : (hoy.getMonth() + 1);
  anio = anio ? Number(anio) : hoy.getFullYear();

  let mesAnt = mes - 1;
  let anioAnt = anio;
  if (mesAnt === 0) {
    mesAnt = 12;
    anioAnt = anio - 1;
  }

  const gastos = getSheet_('Gastos').getDataRange().getValues();
  const ingresos = getSheet_('Ingresos').getDataRange().getValues();
  const tarjetas = getSheet_('Tarjetas').getDataRange().getValues();

  let ingresoBruto = 0;
  for (let i = 1; i < ingresos.length; i++) {
    const f = ingresos[i][0];
    if (!(f instanceof Date)) continue;
    if (f.getFullYear() === anio && f.getMonth() + 1 === mes) {
      ingresoBruto += Number(ingresos[i][3]) || 0;
    }
  }

  let gastoTotal = 0, transferencia = 0, debito = 0;
  const porCategoria = {};

  for (let i = 1; i < gastos.length; i++) {
    const f = gastos[i][0];
    if (!(f instanceof Date)) continue;
    if (f.getFullYear() === anio && f.getMonth() + 1 === mes) {
      const categoria = String(gastos[i][2] || '');
      const medio = String(gastos[i][4] || '');
      const importe = Number(gastos[i][5]) || 0;

      gastoTotal += importe;
      if (medio === 'Transferencia') transferencia += importe;
      if (medio === 'Débito') debito += importe;
      if (categoria) porCategoria[categoria] = (porCategoria[categoria] || 0) + importe;
    }
  }

  // Tarjeta a pagar este mes: compras del mes anterior o cuotas que caen en este mes
  let tarjetaMesAnterior = 0;
  // Nuevas compras realizadas con tarjeta en este mes
  let tarjetaConsumoMes = 0;

  for (let i = 1; i < tarjetas.length; i++) {
    const f = tarjetas[i][0];
    if (!(f instanceof Date)) continue;

    const importe = Number(tarjetas[i][4]) || 0;
    const cuotas = Number(tarjetas[i][5]) || 1;
    const cuotaMensual = Number(tarjetas[i][6]) || (cuotas ? importe / cuotas : importe);
    const primeraCuotaRaw = tarjetas[i][7];
    const primeraCuota = primeraCuotaRaw instanceof Date ? primeraCuotaRaw : (primeraCuotaRaw ? parseFechaLocal_(primeraCuotaRaw + '-01') : null);

    // Consumos realizados en el mes seleccionado
    if (f.getFullYear() === anio && f.getMonth() + 1 === mes) {
      tarjetaConsumoMes += importe;
    }

    // Cuotas o pagos de tarjeta que vencen en este mes seleccionado
    let caeEnEsteMes = false;
    if (cuotas > 1) {
      let pAnio, pMes;
      if (primeraCuota) {
        pAnio = primeraCuota.getFullYear();
        pMes = primeraCuota.getMonth() + 1;
      } else {
        pMes = f.getMonth() + 2;
        pAnio = f.getFullYear();
        if (pMes > 12) { pMes -= 12; pAnio += 1; }
      }
      const diff = (anio - pAnio) * 12 + (mes - pMes);
      if (diff >= 0 && diff < cuotas) {
        caeEnEsteMes = true;
      }
    } else {
      if (primeraCuota) {
        if (primeraCuota.getFullYear() === anio && primeraCuota.getMonth() + 1 === mes) {
          caeEnEsteMes = true;
        }
      } else {
        // Por defecto: se paga el mes siguiente a la compra (mes anterior)
        if (f.getFullYear() === anioAnt && f.getMonth() + 1 === mesAnt) {
          caeEnEsteMes = true;
        }
      }
    }

    if (caeEnEsteMes) {
      tarjetaMesAnterior += (cuotas > 1 ? cuotaMensual : importe);
    }
  }

  // Descuento solicitado: el ingreso disponible resta el monto total de tarjeta del mes pasado
  const ingresoNeto = ingresoBruto - tarjetaMesAnterior;
  const disponible = ingresoNeto - gastoTotal;

  return {
    mes,
    anio,
    mesAnt,
    anioAnt,
    ingresoBruto,
    ingresoNeto,
    ingresoTotal: ingresoNeto, // Mantiene compatibilidad con el resto de la interfaz
    tarjetaMesAnterior,
    tarjetaConsumoMes,
    tarjetaTotal: tarjetaConsumoMes, // Consumo con crédito del mes seleccionado
    gastoTotal,
    transferencia,
    debito,
    disponible,
    porCategoria
  };
}

function getInitialData(mes, anio) {
  const hoy = new Date();
  mes = mes ? Number(mes) : (hoy.getMonth() + 1);
  anio = anio ? Number(anio) : hoy.getFullYear();
  return {
    config: getConfig(),
    resumen: getResumen(mes, anio),
    gastosFijosConfig: getGastosFijosConfig(mes, anio),
    presupuestos: getPresupuestosConfig(),
    metasAhorro: getMetasAhorro()
  };
}

function getHistorial(filtros) {
  filtros = filtros || {};
  const tipo = filtros.tipo === 'ingreso' ? 'ingreso' : 'gasto';
  const sheetName = tipo === 'ingreso' ? 'Ingresos' : 'Gastos';
  const data = getSheet_(sheetName).getDataRange().getValues();

  const filas = [];
  for (let i = 1; i < data.length; i++) {
    const fecha = data[i][0];
    if (!(fecha instanceof Date)) continue;
    const mes = fecha.getMonth() + 1;
    const anio = fecha.getFullYear();
    if (filtros.mes && Number(filtros.mes) !== mes) continue;
    if (filtros.anio && Number(filtros.anio) !== anio) continue;

    if (tipo === 'gasto') {
      if (filtros.categoria && filtros.categoria !== String(data[i][2])) continue;
      filas.push({
        fila: i + 1,
        fechaObj: fecha,
        descripcion: String(data[i][1] || ''),
        categoria: String(data[i][2] || ''),
        entidad: String(data[i][3] || ''),
        medio: String(data[i][4] || ''),
        importe: Number(data[i][5]) || 0
      });
    } else {
      filas.push({
        fila: i + 1,
        fechaObj: fecha,
        descripcion: String(data[i][1] || ''),
        tipo: String(data[i][2] || ''),
        importe: Number(data[i][3]) || 0
      });
    }
  }

  filas.sort((a, b) => b.fechaObj - a.fechaObj);
  const tz = Session.getScriptTimeZone();
  const resultado = filas.map(f => {
    const base = {
      fila: f.fila,
      fecha: Utilities.formatDate(f.fechaObj, tz, 'dd/MM/yyyy'),
      descripcion: f.descripcion,
      importe: f.importe
    };
    if (tipo === 'gasto') { base.categoria = f.categoria; base.entidad = f.entidad; base.medio = f.medio; }
    else { base.tipo = f.tipo; }
    return base;
  });

  return resultado;
}

// -------------- FUNCION PARA CARDS INTERACTIVA -------------------
function obtenerDetalleV2(tipo, mes, anio) {
  const hoy = new Date();
  mes = mes ? Number(mes) : (hoy.getMonth() + 1);
  anio = anio ? Number(anio) : hoy.getFullYear();

  let mesAnt = mes - 1;
  let anioAnt = anio;
  if (mesAnt === 0) {
    mesAnt = 12;
    anioAnt = anio - 1;
  }

  const tz = Session.getScriptTimeZone();
  const formatF = f => (f instanceof Date ? Utilities.formatDate(f, tz, 'dd/MM/yyyy') : String(f || ''));

  if (tipo === 'ingresos') {
    const ingresos = getSheet_('Ingresos').getDataRange().getValues();
    const resultado = [];
    for (let i = 1; i < ingresos.length; i++) {
      const r = ingresos[i];
      const f = r[0];
      if (!(f instanceof Date)) continue;
      if (f.getFullYear() === anio && f.getMonth() + 1 === mes && (r[1] || r[3])) {
        resultado.push({
          fila: i + 1,
          fecha: formatF(r[0]),
          descripcion: r[1],
          tipoIngreso: r[2],
          monto: Number(r[3]) || 0,
          observaciones: r[4]
        });
      }
    }
    return resultado;
  }

  if (tipo === 'tarjetas' || tipo === 'tarjetas_a_pagar') {
    const tarjetas = getSheet_('Tarjetas').getDataRange().getValues();
    const resultado = [];
    for (let i = 1; i < tarjetas.length; i++) {
      const r = tarjetas[i];
      const f = r[0];
      if (!(f instanceof Date)) continue;
      const importe = Number(r[4]) || 0;
      const cuotas = Number(r[5]) || 1;
      const cuotaMensual = Number(r[6]) || (cuotas ? importe / cuotas : importe);
      const primeraCuotaRaw = r[7];
      const primeraCuota = primeraCuotaRaw instanceof Date ? primeraCuotaRaw : (primeraCuotaRaw ? parseFechaLocal_(primeraCuotaRaw + '-01') : null);

      let caeEnEsteMes = false;
      let cuotaTexto = '1 cuota (Mes anterior)';

      if (cuotas > 1) {
        let pAnio, pMes;
        if (primeraCuota) {
          pAnio = primeraCuota.getFullYear();
          pMes = primeraCuota.getMonth() + 1;
        } else {
          pMes = f.getMonth() + 2;
          pAnio = f.getFullYear();
          if (pMes > 12) { pMes -= 12; pAnio += 1; }
        }
        const diff = (anio - pAnio) * 12 + (mes - pMes);
        if (diff >= 0 && diff < cuotas) {
          caeEnEsteMes = true;
          cuotaTexto = 'Cuota ' + (diff + 1) + '/' + cuotas;
        }
      } else {
        if (primeraCuota) {
          if (primeraCuota.getFullYear() === anio && primeraCuota.getMonth() + 1 === mes) {
            caeEnEsteMes = true;
            cuotaTexto = '1 cuota (Programada)';
          }
        } else {
          if (f.getFullYear() === anioAnt && f.getMonth() + 1 === mesAnt) {
            caeEnEsteMes = true;
            cuotaTexto = 'Compra mes anterior';
          }
        }
      }

      if (caeEnEsteMes) {
        resultado.push({
          fila: i + 1,
          fecha: formatF(f),
          descripcion: r[1],
          entidad: r[2],
          categoria: r[3],
          monto: cuotas > 1 ? cuotaMensual : importe,
          cuotas: cuotaTexto
        });
      }
    }
    return resultado;
  }

  if (tipo === 'tarjetas_consumo') {
    const tarjetas = getSheet_('Tarjetas').getDataRange().getValues();
    const resultado = [];
    for (let i = 1; i < tarjetas.length; i++) {
      const r = tarjetas[i];
      const f = r[0];
      if (!(f instanceof Date)) continue;
      if (f.getFullYear() === anio && f.getMonth() + 1 === mes && (r[1] || r[4])) {
        resultado.push({
          fila: i + 1,
          fecha: formatF(r[0]),
          descripcion: r[1],
          entidad: r[2],
          categoria: r[3],
          monto: Number(r[4]) || 0,
          cuotas: (r[5] || 1) + ' cuota(s)'
        });
      }
    }
    return resultado;
  }

  if (tipo === 'transferencia' || tipo === 'debito') {
    const medioBuscado = tipo === 'transferencia' ? 'Transferencia' : 'Débito';
    const gastos = getSheet_('Gastos').getDataRange().getValues();
    const resultado = [];
    for (let i = 1; i < gastos.length; i++) {
      const r = gastos[i];
      const f = r[0];
      if (!(f instanceof Date)) continue;
      if (f.getFullYear() === anio && f.getMonth() + 1 === mes && String(r[4] || '') === medioBuscado) {
        resultado.push({
          fila: i + 1,
          fecha: formatF(r[0]),
          categoria: r[2],
          entidad: r[3],
          descripcion: r[1],
          monto: Number(r[5]) || 0
        });
      }
    }
    return resultado;
  }

  if (tipo === 'gastos') {
    const gastos = getSheet_('Gastos').getDataRange().getValues();
    const resultado = [];
    for (let i = 1; i < gastos.length; i++) {
      const r = gastos[i];
      const f = r[0];
      if (!(f instanceof Date)) continue;
      if (f.getFullYear() === anio && f.getMonth() + 1 === mes && (r[1] || r[5])) {
        resultado.push({
          fila: i + 1,
          fecha: formatF(r[0]),
          categoria: r[2],
          entidad: r[3],
          descripcion: r[1],
          monto: Number(r[5]) || 0
        });
      }
    }
    return resultado;
  }

  return [];
}

  

// --------------FIN CARD INTERACTIVA -------------------

function getEvolucionMensual() {
  const tz = Session.getScriptTimeZone();
  const hoy = new Date();
  const meses = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1);
    meses.push({
      anio: d.getFullYear(),
      mes: d.getMonth() + 1,
      label: Utilities.formatDate(d, tz, 'MMM'),
      ingresos: 0,
      gastos: 0,
      catGastos: {}
    });
  }
  const clave = (anio, mes) => anio + '-' + mes;
  const mapa = {};
  meses.forEach(m => mapa[clave(m.anio, m.mes)] = m);

  // 1. Ingresos
  const ingresos = getSheet_('Ingresos').getDataRange().getValues();
  for (let i = 1; i < ingresos.length; i++) {
    const f = ingresos[i][0];
    if (!(f instanceof Date)) continue;
    const k = clave(f.getFullYear(), f.getMonth() + 1);
    if (mapa[k]) {
      mapa[k].ingresos += Number(ingresos[i][3]) || 0;
    }
  }

  // 2. Gastos por categoría
  const acumuladoCatGlobal = {};
  const gastos = getSheet_('Gastos').getDataRange().getValues();
  for (let i = 1; i < gastos.length; i++) {
    const f = gastos[i][0];
    if (!(f instanceof Date)) continue;
    const k = clave(f.getFullYear(), f.getMonth() + 1);
    const importe = Number(gastos[i][5]) || 0;
    const cat = String(gastos[i][2] || 'Otros').trim();

    if (mapa[k]) {
      mapa[k].gastos += importe;
      if (cat) {
        mapa[k].catGastos[cat] = (mapa[k].catGastos[cat] || 0) + importe;
        acumuladoCatGlobal[cat] = (acumuladoCatGlobal[cat] || 0) + importe;
      }
    }
  }

  // 3. Tarjetas a pagar por mes
  const tarjetas = getSheet_('Tarjetas').getDataRange().getValues();
  for (let i = 1; i < tarjetas.length; i++) {
    const f = tarjetas[i][0];
    if (!(f instanceof Date)) continue;

    const importe = Number(tarjetas[i][4]) || 0;
    const cuotas = Number(tarjetas[i][5]) || 1;
    const cuotaMensual = Number(tarjetas[i][6]) || (cuotas ? importe / cuotas : importe);
    const primeraCuotaRaw = tarjetas[i][7];
    const primeraCuota = primeraCuotaRaw instanceof Date ? primeraCuotaRaw : (primeraCuotaRaw ? parseFechaLocal_(primeraCuotaRaw + '-01') : null);

    meses.forEach(m => {
      let mesAnt = m.mes - 1;
      let anioAnt = m.anio;
      if (mesAnt === 0) { mesAnt = 12; anioAnt--; }

      let caeEnEsteMes = false;
      if (cuotas > 1) {
        let pAnio, pMes;
        if (primeraCuota) {
          pAnio = primeraCuota.getFullYear();
          pMes = primeraCuota.getMonth() + 1;
        } else {
          pMes = f.getMonth() + 2;
          pAnio = f.getFullYear();
          if (pMes > 12) { pMes -= 12; pAnio++; }
        }
        const diff = (m.anio - pAnio) * 12 + (m.mes - pMes);
        if (diff >= 0 && diff < cuotas) {
          caeEnEsteMes = true;
        }
      } else {
        if (primeraCuota) {
          if (primeraCuota.getFullYear() === m.anio && primeraCuota.getMonth() + 1 === m.mes) {
            caeEnEsteMes = true;
          }
        } else {
          if (f.getFullYear() === anioAnt && f.getMonth() + 1 === mesAnt) {
            caeEnEsteMes = true;
          }
        }
      }

      if (caeEnEsteMes) {
        const montoCuota = (cuotas > 1 ? cuotaMensual : importe);
        const k = clave(m.anio, m.mes);
        const cat = 'Tarjeta';
        if (mapa[k]) {
          mapa[k].catGastos[cat] = (mapa[k].catGastos[cat] || 0) + montoCuota;
          acumuladoCatGlobal[cat] = (acumuladoCatGlobal[cat] || 0) + montoCuota;
        }
      }
    });
  }

  // Hallamos las 3 categorías con mayor gasto acumulado en los últimos 6 meses
  const top3Nombres = Object.keys(acumuladoCatGlobal)
    .sort((a, b) => acumuladoCatGlobal[b] - acumuladoCatGlobal[a])
    .slice(0, 3);

  const resultadoMeses = meses.map(m => ({
    label: m.label,
    ingresos: m.ingresos,
    gastos: m.gastos,
    top1: (top3Nombres[0] && m.catGastos[top3Nombres[0]]) || 0,
    top2: (top3Nombres[1] && m.catGastos[top3Nombres[1]]) || 0,
    top3: (top3Nombres[2] && m.catGastos[top3Nombres[2]]) || 0
  }));

  return {
    meses: resultadoMeses,
    top3Nombres: top3Nombres
  };
}

function eliminarMovimiento(tipo, fila, datosVerificacion) {
  tipo = String(tipo || '').toLowerCase();
  let sheetName = 'Gastos';
  if (tipo.indexOf('ingreso') !== -1) {
    sheetName = 'Ingresos';
  } else if (tipo.indexOf('tarjeta') !== -1) {
    sheetName = 'Tarjetas';
  }

  const sh = getSheet_(sheetName);
  const data = sh.getDataRange().getValues();
  let filaABorrar = -1;

  // 1. Intentar por fila sugerida si coincide
  if (fila && Number(fila) >= 2 && Number(fila) <= data.length) {
    const idx = Number(fila) - 1;
    const row = data[idx];
    let coincide = true;
    if (datosVerificacion) {
      if (datosVerificacion.monto !== undefined && datosVerificacion.monto !== null && datosVerificacion.monto !== '') {
        const importeHoja = sheetName === 'Ingresos' ? Number(row[3]) : (sheetName === 'Gastos' ? Number(row[5]) : Number(row[4]));
        if (Math.abs(importeHoja - Number(datosVerificacion.monto)) > 0.02) {
          coincide = false;
        }
      }
    }
    if (coincide) {
      filaABorrar = Number(fila);
    }
  }

  // 2. Si no coincidió o la fila cambió, buscar fila por datos de verificación
  if (filaABorrar === -1 && datosVerificacion) {
    const tz = Session.getScriptTimeZone();
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const f = row[0];
      const fStr = f instanceof Date ? Utilities.formatDate(f, tz, 'dd/MM/yyyy') : String(f || '');
      const descHoja = String(row[1] || '').trim();
      const importeHoja = sheetName === 'Ingresos' ? Number(row[3]) : (sheetName === 'Gastos' ? Number(row[5]) : Number(row[4]));

      const coincideImporte = (datosVerificacion.monto !== undefined && datosVerificacion.monto !== null && datosVerificacion.monto !== '')
        ? Math.abs(importeHoja - Number(datosVerificacion.monto)) < 0.02
        : true;

      const coincideDesc = datosVerificacion.descripcion
        ? descHoja.toLowerCase() === String(datosVerificacion.descripcion).trim().toLowerCase()
        : true;

      const coincideFecha = datosVerificacion.fecha
        ? (fStr === datosVerificacion.fecha || String(datosVerificacion.fecha).indexOf(fStr) !== -1 || fStr.indexOf(String(datosVerificacion.fecha)) !== -1)
        : true;

      if (coincideImporte && (coincideDesc || coincideFecha)) {
        filaABorrar = i + 1;
        break;
      }
    }
  }

  if (filaABorrar < 2) {
    throw new Error('No se pudo encontrar el movimiento a eliminar en la hoja ' + sheetName + '.');
  }

  sh.deleteRow(filaABorrar);

  const hoy = new Date();
  const mes = (datosVerificacion && datosVerificacion.mes) ? Number(datosVerificacion.mes) : (hoy.getMonth() + 1);
  const anio = (datosVerificacion && datosVerificacion.anio) ? Number(datosVerificacion.anio) : hoy.getFullYear();
  return getResumen(mes, anio);
}

function obtenerDetalleCategoria(categoria, mes, anio) {
  const hoy = new Date();
  mes = mes ? Number(mes) : (hoy.getMonth() + 1);
  anio = anio ? Number(anio) : hoy.getFullYear();
  categoria = String(categoria || '').trim();

  const tz = Session.getScriptTimeZone();
  const data = getSheet_('Gastos').getDataRange().getValues();
  const movimientos = [];

  for (let i = 1; i < data.length; i++) {
    const f = data[i][0];
    if (!(f instanceof Date)) continue;
    if (f.getFullYear() === anio && (f.getMonth() + 1) === mes) {
      const cat = String(data[i][2] || '').trim();
      if (cat.toLowerCase() === categoria.toLowerCase()) {
        const fStr = Utilities.formatDate(f, tz, 'dd/MM/yyyy');
        const fISO = Utilities.formatDate(f, tz, 'yyyy-MM-dd');
        movimientos.push({
          fila: i + 1,
          fecha: fStr,
          fechaISO: fISO,
          descripcion: String(data[i][1] || '').trim(),
          categoria: cat,
          entidad: String(data[i][3] || '').trim(),
          medio: String(data[i][4] || '').trim(),
          importe: Number(data[i][5]) || 0
        });
      }
    }
  }

  // Ordenar movimientos por fecha descendente (más recientes primero)
  movimientos.sort((a, b) => (b.fechaISO || '').localeCompare(a.fechaISO || ''));

  return {
    categoria: categoria,
    mes: mes,
    anio: anio,
    total: movimientos.reduce((acc, m) => acc + (m.importe || 0), 0),
    movimientos: movimientos
  };
}

function editarMovimiento(tipo, fila, datosNuevos, datosVerificacion) {
  tipo = String(tipo || 'gasto').toLowerCase();
  let sheetName = 'Gastos';
  if (tipo.indexOf('ingreso') !== -1) {
    sheetName = 'Ingresos';
  } else if (tipo.indexOf('tarjeta') !== -1) {
    sheetName = 'Tarjetas';
  }

  const sh = getSheet_(sheetName);
  const data = sh.getDataRange().getValues();
  let filaAEditar = -1;

  // 1. Intentar por fila sugerida
  if (fila && Number(fila) >= 2 && Number(fila) <= data.length) {
    const idx = Number(fila) - 1;
    const row = data[idx];
    let coincide = true;
    if (datosVerificacion && datosVerificacion.monto !== undefined && datosVerificacion.monto !== null && datosVerificacion.monto !== '') {
      const importeHoja = sheetName === 'Ingresos' ? Number(row[3]) : (sheetName === 'Gastos' ? Number(row[5]) : Number(row[4]));
      if (Math.abs(importeHoja - Number(datosVerificacion.monto)) > 0.02) {
        coincide = false;
      }
    }
    if (coincide) {
      filaAEditar = Number(fila);
    }
  }

  // 2. Si no coincide, buscar por datos de verificación
  if (filaAEditar === -1 && datosVerificacion) {
    const tz = Session.getScriptTimeZone();
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const f = row[0];
      const fStr = f instanceof Date ? Utilities.formatDate(f, tz, 'dd/MM/yyyy') : String(f || '');
      const descHoja = String(row[1] || '').trim();
      const importeHoja = sheetName === 'Ingresos' ? Number(row[3]) : (sheetName === 'Gastos' ? Number(row[5]) : Number(row[4]));

      const coincideImporte = (datosVerificacion.monto !== undefined && datosVerificacion.monto !== null && datosVerificacion.monto !== '')
        ? Math.abs(importeHoja - Number(datosVerificacion.monto)) < 0.02
        : true;

      const coincideDesc = datosVerificacion.descripcion
        ? descHoja.toLowerCase() === String(datosVerificacion.descripcion).trim().toLowerCase()
        : true;

      const coincideFecha = datosVerificacion.fecha
        ? (fStr === datosVerificacion.fecha || String(datosVerificacion.fecha).indexOf(fStr) !== -1 || fStr.indexOf(String(datosVerificacion.fecha)) !== -1)
        : true;

      if (coincideImporte && (coincideDesc || coincideFecha)) {
        filaAEditar = i + 1;
        break;
      }
    }
  }

  if (filaAEditar < 2) {
    throw new Error('No se pudo encontrar el movimiento a editar en la hoja ' + sheetName + '.');
  }

  // Aplicar cambios en Gastos (A: Fecha, B: Desc, C: Cat, D: Entidad, E: Medio, F: Importe)
  if (sheetName === 'Gastos') {
    if (datosNuevos.fecha) {
      sh.getRange(filaAEditar, 1).setValue(parseFechaLocal_(datosNuevos.fecha));
    }
    sh.getRange(filaAEditar, 2).setValue(String(datosNuevos.descripcion || '').trim());
    if (datosNuevos.categoria) {
      sh.getRange(filaAEditar, 3).setValue(String(datosNuevos.categoria).trim());
    }
    if (datosNuevos.entidad !== undefined) {
      sh.getRange(filaAEditar, 4).setValue(String(datosNuevos.entidad || '').trim());
    }
    if (datosNuevos.medio) {
      sh.getRange(filaAEditar, 5).setValue(String(datosNuevos.medio).trim());
    }
    if (datosNuevos.importe !== undefined && datosNuevos.importe !== null && datosNuevos.importe !== '') {
      sh.getRange(filaAEditar, 6).setValue(Number(datosNuevos.importe) || 0);
    }
  }

  const hoy = new Date();
  const mes = (datosVerificacion && datosVerificacion.mes) ? Number(datosVerificacion.mes) : (hoy.getMonth() + 1);
  const anio = (datosVerificacion && datosVerificacion.anio) ? Number(datosVerificacion.anio) : hoy.getFullYear();
  return getResumen(mes, anio);
}

// ====================================================================
// FASE B: SISTEMA DE PRESUPUESTOS POR CATEGORÍA Y METAS DE AHORRO
// ====================================================================

function getPresupuestosConfig() {
  try {
    const sh = getSheet_('Presupuestos');
    const last = sh.getLastRow();
    if (last < 2) return {};
    const values = sh.getRange(2, 1, last - 1, 2).getValues();
    const map = {};
    for (let i = 0; i < values.length; i++) {
      const cat = String(values[i][0] || '').trim();
      const monto = Number(values[i][1]) || 0;
      if (cat && monto > 0) {
        map[cat] = monto;
      }
    }
    return map;
  } catch (e) {
    console.error('Error al obtener presupuestos:', e);
    return {};
  }
}

function guardarPresupuestos(presupuestosMap) {
  try {
    const sh = getSheet_('Presupuestos');
    const last = sh.getLastRow();
    if (last >= 2) {
      sh.getRange(2, 1, last - 1, 2).clearContent();
    }
    const filas = [];
    if (presupuestosMap && typeof presupuestosMap === 'object') {
      for (const cat in presupuestosMap) {
        const monto = Number(presupuestosMap[cat]) || 0;
        if (cat.trim() && monto > 0) {
          filas.push([cat.trim(), monto]);
        }
      }
    }
    if (filas.length > 0) {
      sh.getRange(2, 1, filas.length, 2).setValues(filas);
    }
    return getPresupuestosConfig();
  } catch (e) {
    console.error('Error al guardar presupuestos:', e);
    throw new Error('No se pudieron guardar los presupuestos: ' + e.message);
  }
}

function getMetasAhorro() {
  try {
    const sh = getSheet_('MetasAhorro');
    const last = sh.getLastRow();
    if (last < 2) return [];
    const values = sh.getRange(2, 1, last - 1, 6).getValues();
    const tz = Session.getScriptTimeZone();
    const metas = [];
    for (let i = 0; i < values.length; i++) {
      const id = String(values[i][0] || ('meta_' + (i + 1))).trim();
      const nombre = String(values[i][1] || '').trim();
      if (!nombre) continue;
      const objetivo = Number(values[i][2]) || 0;
      const actual = Number(values[i][3]) || 0;
      let fechaLimite = values[i][4];
      if (fechaLimite instanceof Date) {
        fechaLimite = Utilities.formatDate(fechaLimite, tz, 'yyyy-MM-dd');
      } else {
        fechaLimite = String(fechaLimite || '');
      }
      const color = String(values[i][5] || '#38bdf8').trim();
      metas.push({
        id,
        nombre,
        objetivo,
        actual,
        fechaLimite,
        color
      });
    }
    return metas;
  } catch (e) {
    console.error('Error al obtener metas de ahorro:', e);
    return [];
  }
}

function guardarMetaAhorro(meta) {
  try {
    const sh = getSheet_('MetasAhorro');
    const nombre = String(meta.nombre || '').trim();
    if (!nombre) throw new Error('Ingresá un nombre para la meta.');
    const objetivo = Number(meta.objetivo) || 0;
    if (objetivo <= 0) throw new Error('El monto objetivo debe ser mayor a 0.');
    const actual = Math.max(0, Number(meta.actual) || 0);
    const fechaLimite = meta.fechaLimite ? parseFechaLocal_(meta.fechaLimite) : '';
    const color = meta.color || '#38bdf8';
    
    const last = sh.getLastRow();
    let filaEncontrada = -1;
    if (meta.id && last >= 2) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]));
      const idx = ids.indexOf(String(meta.id));
      if (idx !== -1) {
        filaEncontrada = idx + 2;
      }
    }

    if (filaEncontrada !== -1) {
      sh.getRange(filaEncontrada, 2, 1, 5).setValues([[nombre, objetivo, actual, fechaLimite, color]]);
    } else {
      const nuevoId = meta.id || ('meta_' + Date.now());
      sh.appendRow([nuevoId, nombre, objetivo, actual, fechaLimite, color]);
    }
    return getMetasAhorro();
  } catch (e) {
    console.error('Error al guardar meta de ahorro:', e);
    throw new Error('Error al guardar meta: ' + e.message);
  }
}

function aportarMetaAhorro(id, montoAportado) {
  try {
    const sh = getSheet_('MetasAhorro');
    const last = sh.getLastRow();
    if (last < 2) throw new Error('Meta no encontrada.');
    const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]));
    const idx = ids.indexOf(String(id));
    if (idx === -1) throw new Error('Meta no encontrada.');
    const fila = idx + 2;
    const actual = Number(sh.getRange(fila, 4).getValue()) || 0;
    const nuevoTotal = Math.max(0, actual + (Number(montoAportado) || 0));
    sh.getRange(fila, 4).setValue(nuevoTotal);
    return getMetasAhorro();
  } catch (e) {
    console.error('Error al aportar a la meta:', e);
    throw new Error('Error al aportar: ' + e.message);
  }
}

function eliminarMetaAhorro(id) {
  try {
    const sh = getSheet_('MetasAhorro');
    const last = sh.getLastRow();
    if (last < 2) return [];
    const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]));
    const idx = ids.indexOf(String(id));
    if (idx !== -1) {
      sh.deleteRow(idx + 2);
    }
    return getMetasAhorro();
  } catch (e) {
    console.error('Error al eliminar meta:', e);
    throw new Error('Error al eliminar meta: ' + e.message);
  }
}


