function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('ACTUALIZAR')
    .addItem('COMPRAS', 'obtenerDatosYActualizar')
    .addItem('SERVICIO', 'obtenerDatosYActualizar_s')
    .addToUi();
}
  let nuevosDatosHoja = [];
  let nuevosDatos = [];

function obtenerDatosYActualizar() {

  // IDs de las hojas de cálculo principales
  const ids = [
    { id: '13h3phJ0MXop5cvC4Rd79YeTZUflqNXRmG3YtYi5uRdk', base: 'P1' },
    { id: '1WP4JA0-2DSPKn2eKxL5DRDgk3dwGuOTlLfD2f03EZ1k', base: 'P2' },
    { id: '1eJ5h2huabI1y-6caDzKAbAPDHx4Za8AMyWnLQUz8hiE', base: 'P3' },
    { id: '1xvLwNUPviBNu_CEGo-N7C2DuiaH2fVqbOlOLH9RiTWE', base: 'P4' },
    { id: '1vEUH7Etw4wBENv5WDEKHXD3gal85iFUNdTVVQ3VKD0w', base: 'P5' },
    { id: '1D90cfJZGlLlcMhSdwVaeVIxDxGAAsA2Qt2qlaiakWRw', base: 'P6' },
    { id: '1bSQVZshCRJ_fJn9C6hEFcaDuFRwF9Q_XFelVlS2BFEE', base: 'P7' },
    { id: '1dVDJPbJdaI-q9ITj0Kc23U1TN_t1z8iRrIY8rqtqZ9s', base: 'P8' },
    { id: '1gUlTkn47r0DyPPQo8NoZSM1NFOHdGBYli3mdZ6nn1HY', base: 'P9' }
  ];

  // Pesos por proceso
  const pesosProcesos = {
    descubrimiento: 25,
    compra: 20,
    entrega: 20,
    lealtad: 25,
    habilitadores: 10
  };

  const consId = '1FWpTR65G5WLzIRraTtAejRq24aojsJYv5x7S_KZ44lg';

  // Columnas a extraer
  const columnas = {
    pais: 'C', ciclo: 'E', codigo: 'G', concesionario: 'F',
    ciudad: 'D', fechaVisita: 'A', sede: 'H'
  };

  const columnasCons = {
    pregunta: 'C', proceso: 'G', subproceso: 'H', elemento: 'I',
    numElemento: 'E', numPeso: 'K', porPeso: 'L'
  };

  const consSpreadsheet = SpreadsheetApp.openById(consId);
  const consSheet = consSpreadsheet.getSheetByName('SUBPROCESOS');
  const consData = consSheet.getDataRange().getValues();

  // Obtener datos de COORDENADAS
  const coordSheet = consSpreadsheet.getSheetByName('COORDENADAS');
  if (!coordSheet) {
    console.error("No se encontró la hoja COORDENADAS");
    return;
  }
  const coordData = coordSheet.getDataRange().getValues();

  function buscarCoordenada(pais, concesionario, sede, ciudad) {
    const normalize = (str) => String(str).trim().toUpperCase();
    
    for (let i = 1; i < coordData.length; i++) {
      const row = coordData[i];
      if (row.length < 5) continue;
      
      if (normalize(row[0]) === normalize(pais) &&
          normalize(row[1]) === normalize(concesionario) &&
          normalize(row[2]) === normalize(sede) &&
          normalize(row[3]) === normalize(ciudad)) {
        return row[4] || "";
      }
    }
    return "";
  }

  // Recorrer cada hoja de cálculo principal
  ids.forEach(({ id, base }) => {
    try {
      const spreadsheet = SpreadsheetApp.openById(id);
      const sheet = spreadsheet.getSheetByName('Respuestas de formulario 1');
      const formatoSheet = spreadsheet.getSheetByName('Formato');

      // Validar que existan datos
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) {
        console.log(`Hoja ${id} no tiene datos suficientes`);
        return;
      }

      const preguntas = data[0].slice(10);
      const ultimaPregunta = preguntas.findIndex(p => p === '');
      const preguntasFiltradas = ultimaPregunta === -1 ? preguntas : preguntas.slice(0, ultimaPregunta);

      const dataconteoP = {};
      const dataResp = {};

      data.slice(1).forEach((row) => {
        const pais = row[columnas.pais.charCodeAt(0) - 65];
        const ciclo = row[columnas.ciclo.charCodeAt(0) - 65];
        const codigo = row[columnas.codigo.charCodeAt(0) - 65];
        const concesionario = row[columnas.concesionario.charCodeAt(0) - 65];
        const ciudad = row[columnas.ciudad.charCodeAt(0) - 65];
        const fechaVisita = row[columnas.fechaVisita.charCodeAt(0) - 65];
        const sede = row[columnas.sede.charCodeAt(0) - 65];
        const fechaFormateada = Utilities.formatDate(new Date(fechaVisita), Session.getScriptTimeZone(), 'dd/MM/yyyy');

        preguntasFiltradas.forEach((pregunta, colIndex) => {
          const respuesta = row[10 + colIndex];
          if (respuesta !== '') {
            const filaCons = consData.find(fila => fila[columnasCons.pregunta.charCodeAt(0) - 65] === pregunta);
            const proceso = filaCons ? filaCons[columnasCons.proceso.charCodeAt(0) - 65] : '';
            const subproceso = filaCons ? filaCons[columnasCons.subproceso.charCodeAt(0) - 65] : '';
            const elemento = filaCons ? filaCons[columnasCons.elemento.charCodeAt(0) - 65] : '';
            const numElemento = filaCons ? filaCons[columnasCons.numElemento.charCodeAt(0) - 65] : '';
            const numPeso = filaCons ? filaCons[columnasCons.numPeso.charCodeAt(0) - 65] : '';
            const porPeso = filaCons ? filaCons[columnasCons.porPeso.charCodeAt(0) - 65] : '';
            const concat = `${pais}-${ciclo}-${codigo}-${concesionario}-${ciudad}-${fechaFormateada}-${id}-${base}-${numElemento}`;

            if (!dataconteoP[concat]) {
              dataconteoP[concat] = 1;
              dataResp[concat] = (respuesta === 'Si') ? 1 : 0;
            } else {
              dataconteoP[concat] += 1;
              if (respuesta === 'Si') dataResp[concat] += 1;
            }

            nuevosDatosHoja.push({
              pais, ciclo, codigo, concesionario, ciudad, fechaVisita, pregunta, 
              respuesta, proceso, subproceso, elemento, numElemento, sede, base, 
              concat, conteoP: null, puntajeP: null, aprobacion: null, numPeso, porPeso
            });
          }
        });
      });

      // Calcular aprobación para cada registro
      nuevosDatosHoja.forEach(item => {
        const key = item.concat;
        if (dataconteoP[key]) {
          item.conteoP = dataconteoP[key];
          item.puntajeP = dataResp[key];
          item.aprobacion = (dataconteoP[key] === dataResp[key]) ? 'SI' : 'NO';
        }
      });

      // Acumular datos de todas las hojas
      nuevosDatos = nuevosDatos.concat(nuevosDatosHoja);

      // Procesar para generar resumen
      let newDic = {};
      nuevosDatosHoja.forEach((item) => {
        if (!item.numElemento || !item.concat) return;
        
        const coordenada = buscarCoordenada(item.pais, item.concesionario, item.sede, item.ciudad);
        const newKey = item.concat;

        if (!newDic[newKey]) {
          newDic[newKey] = {
            paisNew: item.pais,
            cicloNew: item.ciclo,
            codigoNew: item.codigo,
            concesionairoNew: item.concesionario,
            ciudadNew: item.ciudad,
            fechavisitaNew: item.fechaVisita,
            procesoNew: item.proceso,
            subprocesoNew: item.subproceso,
            elementoNew: item.elemento,
            numElementoNew: item.numElemento,
            sedeNew: item.sede,
            baseNew: item.base,
            concatNew: item.concat,
            aprobacionNew: item.aprobacion,
            numPesoNew: item.numPeso,
            porPesoNew: item.porPeso,
            puntajeMaximoNew: item.aprobacion === 'SI' ? item.numPeso : 0,
            calificacionNew: item.aprobacion === 'SI' ? item.porPeso : 0,
            coordenadaNew: coordenada
          };
        }
      });

      const processDictionary = {
        descubrimiento: "1. DESCUBRIMIENTO",
        compra: "2. COMPRA",
        entrega: "3. ENTREGA",
        lealtad: "4. LEALTAD",
        habilitadores: "5. HABILITADORES"
      };

      let newData = {};
      Object.values(newDic).forEach(item => {
        const consede = `${item.concesionairoNew} - ${item.sedeNew}`;
        // Manejo seguro de calificacionNew
        let porcentajeStr = '0';
        if (typeof item.calificacionNew === 'string') {
          porcentajeStr = item.calificacionNew.replace('%', '') || '0';
        } else if (typeof item.calificacionNew === 'number') {
          porcentajeStr = item.calificacionNew.toString();
        }

        const porcentaje = parseFloat(porcentajeStr) || 0;

        if (!newData[consede]) {
          newData[consede] = {
            consede: consede,
            codigo: item.codigoNew,
            ciclo: item.cicloNew,
            pais: item.paisNew,
            ciudad: item.ciudadNew,
            concesionario: item.concesionairoNew,
            sede: item.sedeNew,
            nameSede: item.sedeNew,
            total: null,
            descubrimiento: null,
            compra: null,
            entrega: null,
            lealtad: null,
            habilitadores: null
          };
        }

        const proceso = item.procesoNew;
        if (proceso === processDictionary.descubrimiento) {
          newData[consede].descubrimiento += porcentaje;
        } else if (proceso === processDictionary.compra) {
          newData[consede].compra += porcentaje;
        } else if (proceso === processDictionary.entrega) {
          newData[consede].entrega += porcentaje;      
        } else if (proceso === processDictionary.lealtad) {
          newData[consede].lealtad += porcentaje;
        } else if (proceso === processDictionary.habilitadores) {
          newData[consede].habilitadores += porcentaje;
        }

      });
    
    for (const dataKey in newData) {
      let totalOptDes = newData[dataKey].descubrimiento || 0
      let totalOptCom = newData[dataKey].compra || 0
      let totalOptEnt = newData[dataKey].entrega || 0
      let totalOptLea = newData[dataKey].lealtad || 0
      let totalOptHab = newData[dataKey].habilitadores || 0 

      let porcentajeDes = (totalOptDes * pesosProcesos.descubrimiento)
      let porcentajeCom = (totalOptCom * pesosProcesos.compra) 
      let porcentajeEnt = (totalOptEnt * pesosProcesos.entrega)
      let porcentajeLea = (totalOptLea * pesosProcesos.lealtad)
      let porcentajeHab = (totalOptHab * pesosProcesos.habilitadores)
      
      newData[dataKey].descubrimiento = `${porcentajeDes.toFixed(2)}%`
      newData[dataKey].compra = `${porcentajeCom.toFixed(2)}%`
      newData[dataKey].entrega = `${porcentajeEnt.toFixed(2)}%`
      newData[dataKey].lealtad = `${porcentajeLea.toFixed(2)}%`
      newData[dataKey].habilitadores = `${porcentajeHab.toFixed(2)}%`

      let totalGen = porcentajeDes + porcentajeCom + porcentajeEnt + porcentajeLea + porcentajeHab

      newData[dataKey].total = `${totalGen.toFixed(2)}%`
    }

      const newArray = Object.values(newDic);
      const newArrayData = Object.values(newData);

      // Actualizar hojas de resultados
      const spreadsheetEf = SpreadsheetApp.openById(consId);
      
      // Hoja Compra - EF
      const formatoSheetEf = spreadsheetEf.getSheetByName('Compra - EF');
      formatoSheetEf.clear();
      formatoSheetEf.getRange(1, 1, 1, 19).setValues([[
        'pais', 'ciclo', 'Código', 'Concesionario', 'Ciudad', 'FechaVisita', 
        'Proceso', 'Subproceso', 'Elemento', 'numElemento', 'Sede', 'Base', 
        'Concatenado', 'aprobacion', 'numPeso', 'porPeso', 'Puntaje', 'Calificacion', 'Coordenada'
      ]]);
      
      if (newArray.length > 0) {
        const newArrayValues = newArray.map(obj => Object.values(obj));
        formatoSheetEf.getRange(2, 1, newArrayValues.length, newArrayValues[0].length).setValues(newArrayValues);
      }

      // Hoja Diagnostico - Manual Compra
      const formatoSheetDg = spreadsheetEf.getSheetByName('Diagnostico - Manual Compra');
      formatoSheetDg.clear();
      formatoSheetDg.getRange(1, 1, 1, 14).setValues([[
        'Concesionario - sede', 'Código', 'Ciclo', 'País', 'Ciudad', 
        'Concesionario', 'Sede', 'Nombre Sede', 'Total Ponderado', 
        '1. DESCUBRIMIENTO (25%)', '2. COMPRA (20%)', 
        '3. ENTREGA (20%)', '4. LEALTAD (25%)', '5. HABILITADORES (10%)'
      ]]);
      
      if (newArrayData.length > 0) {
        const newArrayValuesDg = newArrayData.map(obj => [
          obj.consede, obj.codigo, obj.ciclo, obj.pais, obj.ciudad,
          obj.concesionario, obj.sede, obj.nameSede, obj.total,
          obj.descubrimiento, obj.compra, obj.entrega, obj.lealtad, obj.habilitadores
        ]);
        formatoSheetDg.getRange(2, 1, newArrayValuesDg.length, newArrayValuesDg[0].length).setValues(newArrayValuesDg);
      }

      // Hoja Formato
      formatoSheet.clear();
      formatoSheet.getRange(1, 1, 1, 20).setValues([[
        'País', 'Ciclo', 'Código', 'Concesionario', 'Ciudad', 'Fecha Visita', 
        'Pregunta', 'Respuesta', 'Proceso', 'Subproceso1', 'Elemento', '#Elemento', 
        'Sede', 'Base', 'Concatenado', 'conteoP', 'puntajeP', 'aprobacion', 'numPeso', 'porPeso'
      ]]);
      
      if (nuevosDatosHoja.length > 0) {
        const nuevosDatosArray = nuevosDatosHoja.map(obj => Object.values(obj));
        formatoSheet.getRange(2, 1, nuevosDatosArray.length, nuevosDatosArray[0].length).setValues(nuevosDatosArray);
      }

    } catch (e) {
      console.error(`Error procesando ${id}:`, e);
    }
  });
}
