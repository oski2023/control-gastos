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

  return {
    categorias: ordenarAZ(limpiar(values.map(r => r[0]))),
    medios: ordenarAZ(limpiar(values.map(r => r[1]))),
    entidades: ordenarAZ(limpiar(values.map(r => r[2])))
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

function getGastosFijosConfig() {
  const sh = getSheet_('GastosFijos');
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, 3).getValues();
  const lista = [];
  for (let i = 0; i < vals.length; i++) {
    const cat = String(vals[i][0] || '').trim();
    if (!cat) continue;
    const dia = Number(vals[i][1]) || 1;
    const activoVal = String(vals[i][2] || 'SI').trim().toUpperCase();
    const activo = (activoVal !== 'NO' && activoVal !== 'PAUSADO' && activoVal !== 'INACTIVO');
    lista.push({
      fila: i + 2,
      categoria: cat,
      diaVencimiento: dia,
      activo: activo
    });
  }
  return lista.sort((a, b) => a.categoria.localeCompare(b.categoria, 'es', { sensitivity: 'base' }));
}

function addGastoFijo(categoria, diaVencimiento) {
  categoria = String(categoria || '').trim();
  diaVencimiento = Math.max(1, Math.min(31, Number(diaVencimiento) || 1));
  if (!categoria) throw new Error('Ingresá el nombre o categoría del gasto fijo.');

  const sh = getSheet_('GastosFijos');
  sh.appendRow([ categoria, diaVencimiento, 'SI' ]);
  return getGastosFijosConfig();
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
  const fijos = getSheet_('GastosFijos').getDataRange().getValues();
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
        descripcion: String(desc || '')
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

    const catKey = normCat(categoria);
    let pagosAnt = pagosMesAnt[catKey] || 0;
    let pagosAct = pagosMesActual[catKey] || 0;
    const listaPagosEsteMes = pagosInfoMesActual[catKey] || [];

    // 1. ¿El mes actual tiene pagos registrados?
    if (pagosAct > 0) {
      // Ya fue pagado este mes (está al día)
      const totalMonto = listaPagosEsteMes.reduce((acc, p) => acc + (p.importe || 0), 0);
      const ultPago = listaPagosEsteMes[listaPagosEsteMes.length - 1] || {};
      alDia.push({
        categoria,
        diaVencimiento,
        monto: totalMonto || ultPago.importe || 0,
        fechaPago: ultPago.fecha ? Utilities.formatDate(ultPago.fecha, tz, 'dd/MM/yyyy') : '',
        medio: ultPago.medio || '',
        cantidadPagos: pagosAct
      });
    } else {
      // No se ha pagado aún este mes
      // 1.a. Si tampoco se pagó el mes anterior, arrastra la deuda del mes anterior
      if (pagosAnt === 0) {
        const fechaVtoAnt = new Date(anioAnt, mesAnt - 1, diaVencimiento || 1);
        const diffMs = hoy.getTime() - fechaVtoAnt.getTime();
        const diasVencidoAnt = Math.max(1, Math.floor(diffMs / (1000 * 60 * 60 * 24)));

        pendientes.push({
          categoria,
          diaVencimiento,
          mesNombre: getNombreMes(mesAnt),
          mesNum: mesAnt,
          anio: anioAnt,
          periodo: 'anterior',
          esMesAnterior: true,
          diasVencido: diasVencidoAnt
        });
      }

      // 1.b. Pendiente o próximo en el mes actual
      if (diaActual >= diaVencimiento) {
        // Vencido este mes
        pendientes.push({
          categoria,
          diaVencimiento,
          mesNombre: getNombreMes(mesActual),
          mesNum: mesActual,
          anio: anioActual,
          periodo: 'actual',
          esMesAnterior: false,
          diasVencido: diaActual - diaVencimiento
        });
      } else {
        // Vence más adelante en el mes
        proximos.push({
          categoria,
          diaVencimiento,
          diasFaltan: diaVencimiento - diaActual
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
    gastosFijosConfig: getGastosFijosConfig()
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

