const SS = SpreadsheetApp.getActiveSpreadsheet();

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Control de Gastos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet_(name) {
  const sh = SS.getSheetByName(name);
  if (!sh) throw new Error('No existe la hoja: ' + name);
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

return {
  categorias: limpiar(values.map(r => r[0])),
  medios: limpiar(values.map(r => r[1])),
  entidades: limpiar(values.map(r => r[2]))
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
  const hoy = new Date();
  const mesActual = mes ? Number(mes) : (hoy.getMonth() + 1);
  const anioActual = anio ? Number(anio) : hoy.getFullYear();
  const esMesActual = (anioActual === hoy.getFullYear() && mesActual === hoy.getMonth() + 1);
  const diaActual = esMesActual ? hoy.getDate() : 31;

  const cargadosEsteMes = new Set();
  for (let i = 1; i < gastos.length; i++) {
    const f = gastos[i][0];
    if (!(f instanceof Date)) continue;
    if (f.getFullYear() === anioActual && f.getMonth() + 1 === mesActual) {
      cargadosEsteMes.add(String(gastos[i][2] || ''));
    }
  }

  // también revisamos las compras o pagos con tarjeta
  for (let i = 1; i < tarjetas.length; i++) {
    const f = tarjetas[i][0];
    if (!(f instanceof Date)) continue;
    if (f.getFullYear() === anioActual && f.getMonth() + 1 === mesActual) {
      cargadosEsteMes.add(String(tarjetas[i][3] || ''));
    }
  }

  const pendientes = [];
  for (let i = 1; i < fijos.length; i++) {
    const categoria = String(fijos[i][0] || '').trim();
    const diaVencimiento = Number(fijos[i][1]) || 0;
    if (!categoria) continue;

    if (!cargadosEsteMes.has(categoria) && diaActual >= diaVencimiento) {
      pendientes.push({ categoria, diaVencimiento, diasVencido: diaActual - diaVencimiento });
    }
  }

  pendientes.sort((a, b) => b.diasVencido - a.diasVencido);
  return pendientes;
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
  return { config: getConfig(), resumen: getResumen(mes, anio) };
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
        fechaObj: fecha,
        descripcion: String(data[i][1] || ''),
        categoria: String(data[i][2] || ''),
        entidad: String(data[i][3] || ''),
        medio: String(data[i][4] || ''),
        importe: Number(data[i][5]) || 0
      });
    } else {
      filas.push({
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
    return ingresos.slice(1)
      .filter(r => {
        const f = r[0];
        if (!(f instanceof Date)) return false;
        return (f.getFullYear() === anio && f.getMonth() + 1 === mes) && (r[1] || r[3]);
      })
      .map(r => ({
        fecha: formatF(r[0]),
        descripcion: r[1],
        tipoIngreso: r[2],
        monto: Number(r[3]) || 0,
        observaciones: r[4]
      }));
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
    return tarjetas.slice(1)
      .filter(r => {
        const f = r[0];
        if (!(f instanceof Date)) return false;
        return (f.getFullYear() === anio && f.getMonth() + 1 === mes) && (r[1] || r[4]);
      })
      .map(r => ({
        fecha: formatF(r[0]),
        descripcion: r[1],
        entidad: r[2],
        categoria: r[3],
        monto: Number(r[4]) || 0,
        cuotas: (r[5] || 1) + ' cuota(s)'
      }));
  }

  if (tipo === 'transferencia' || tipo === 'debito') {
    const medioBuscado = tipo === 'transferencia' ? 'Transferencia' : 'Débito';
    const gastos = getSheet_('Gastos').getDataRange().getValues();
    return gastos.slice(1)
      .filter(r => {
        const f = r[0];
        if (!(f instanceof Date)) return false;
        return (f.getFullYear() === anio && f.getMonth() + 1 === mes) && String(r[4] || '') === medioBuscado;
      })
      .map(r => ({
        fecha: formatF(r[0]),
        categoria: r[2],
        entidad: r[3],
        descripcion: r[1],
        monto: Number(r[5]) || 0
      }));
  }

  if (tipo === 'gastos') {
    const gastos = getSheet_('Gastos').getDataRange().getValues();
    return gastos.slice(1)
      .filter(r => {
        const f = r[0];
        if (!(f instanceof Date)) return false;
        return (f.getFullYear() === anio && f.getMonth() + 1 === mes) && (r[1] || r[5]);
      })
      .map(r => ({
        fecha: formatF(r[0]),
        categoria: r[2],
        entidad: r[3],
        descripcion: r[1],
        monto: Number(r[5]) || 0
      }));
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

