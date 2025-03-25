/* function onOpen() {
  // Obtener la interfaz de usuario de la hoja de cálculo
  const ui = SpreadsheetApp.getUi();

  // Crear un menú personalizado
  ui.createMenu('Acciones Personalizadas')
    .addItem('Ejecutar Actualización', 'obtenerDatosYActualizar') // Añadir un ítem al menú
    .addToUi(); // Añadir el menú a la interfaz de usuario
}*/


 function obtenerDatosYActualizar() {

  // IDs de las hojas de cálculo principales y su correspondiente base (P1, P2, etc.)
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

  // ID de la hoja de cálculo adicional para buscar datos
  const consId = '1FWpTR65G5WLzIRraTtAejRq24aojsJYv5x7S_KZ44lg';

  // Columnas a extraer de las hojas principales
  const columnas = {
    pais: 'C',
    ciclo: 'E',
    codigo: 'G',
    concesionario: 'F',
    ciudad: 'D',
    fechaVisita: 'A',
    sede: 'H'
  };

  // Columnas a buscar en la hoja adicional
  const columnasCons = {
    pregunta: 'C',
    proceso: 'G',
    subproceso: 'H',
    elemento: 'I',
    numElemento: 'E'
  };

  // Abrir la hoja adicional "SUBPROCESOS"
  const consSpreadsheet = SpreadsheetApp.openById(consId);
  const consSheet = consSpreadsheet.getSheetByName('SUBPROCESOS');
  const consData = consSheet.getDataRange().getValues();

  // Recorrer cada hoja de cálculo principal
  ids.forEach(({ id, base }) => {
    const spreadsheet = SpreadsheetApp.openById(id);
    const sheet = spreadsheet.getSheetByName('Respuestas de formulario 1');
    const formatoSheet = spreadsheet.getSheetByName('Formato');

    // Obtener los datos de la hoja "Respuestas de formulario 1"
    const data = sheet.getDataRange().getValues();

    // Obtener los encabezados de las preguntas (K1, L1, M1, etc.)
    const preguntas = data[0].slice(10); // Desde la columna K (índice 10) hasta el final
    const ultimaPregunta = preguntas.findIndex(p => p === ''); // Encontrar el primer campo vacío
    const preguntasFiltradas = ultimaPregunta === -1 ? preguntas : preguntas.slice(0, ultimaPregunta);

    // Preparar los datos para la hoja "Formato"
    const nuevosDatos = [];
    const dataConteo = {};
    const dataResp = {};
    data.slice(1).forEach((row, rowIndex) => { // Ignorar la fila de encabezados
      // Extraer las columnas especificadas
      const pais = row[columnas.pais.charCodeAt(0) - 65];
      const ciclo = row[columnas.ciclo.charCodeAt(0) - 65];
      const codigo = row[columnas.codigo.charCodeAt(0) - 65];
      const concesionario = row[columnas.concesionario.charCodeAt(0) - 65];
      const ciudad = row[columnas.ciudad.charCodeAt(0) - 65];
      const fechaVisita = row[columnas.fechaVisita.charCodeAt(0) - 65];
      const sede = row[columnas.sede.charCodeAt(0) - 65];
      const fechaFormateada = Utilities.formatDate(new Date(fechaVisita), Session.getScriptTimeZone(), 'dd/MM/yyyy');
     

      // Recorrer las preguntas y agregar filas en "Formato"
      preguntasFiltradas.forEach((pregunta, colIndex) => {
        const respuesta = row[10 + colIndex]; // Respuesta correspondiente a la pregunta
        if (respuesta !== '') { // Solo agregar si hay una respuesta
          // Buscar en la hoja adicional "SUBPROCESOS"
          const filaCons = consData.find(fila => fila[columnasCons.pregunta.charCodeAt(0) - 65] === pregunta);
          const proceso = filaCons ? filaCons[columnasCons.proceso.charCodeAt(0) - 65] : '';
          const subproceso = filaCons ? filaCons[columnasCons.subproceso.charCodeAt(0) - 65] : '';
          const elemento = filaCons ? filaCons[columnasCons.elemento.charCodeAt(0) - 65] : '';
          const numElemento = filaCons ? filaCons[columnasCons.numElemento.charCodeAt(0) - 65] : '';
          let concat = `${pais}-${ciclo}-${codigo}-${concesionario}-${ciudad}-${fechaFormateada}-${id}-${base}-${numElemento}`
          let conteo = null
          let puntaje = null
          let calificacion = null

          if (!dataConteo[concat] || dataConteo[concat].conteo == '') {
            dataConteo[concat] = 1 
            if (respuesta == 'Si') {
              dataResp[concat] = 1
            } else {
              dataResp[concat] = 0
            }
          } else {
            dataConteo[concat] += 1 
            if (respuesta == 'Si') {
              dataResp[concat] +=  1
            }
          }

          nuevosDatos.push({
            pais, ciclo, codigo, concesionario, ciudad, fechaVisita, pregunta, respuesta, proceso, subproceso, elemento, numElemento, sede, base, concat, conteo, puntaje, calificacion
        });
        }
      });
    });

    for (let index = 0; index < nuevosDatos.length; index++) {
      let key = nuevosDatos[index].concat
      if (dataConteo[key]) {
        let numConteo = dataConteo[key]
        let puntaje = dataResp[key]
        nuevosDatos[index].conteo = numConteo
        nuevosDatos[index].puntaje = puntaje

        if (numConteo !== puntaje) {
          nuevosDatos[index].calificacion = '0%'
        } else {
          nuevosDatos[index].calificacion = '100%'
        }
      }
    }
  

    let newDic = {}
    for (let j = 0; j < nuevosDatos.length; j++) {
      if (!nuevosDatos[j].numElemento) continue
      if (!nuevosDatos[j].concat) continue

      let newKey = nuevosDatos[j].concat
      if (!newDic[newKey] || !newDic[newKey].concede) {
        const paisNew = nuevosDatos[j].pais
        const cicloNew = nuevosDatos[j].ciclo
        const codigoNew = nuevosDatos[j].codigo
        const concesionairoNew = nuevosDatos[j].concesionario
        const ciudadNew = nuevosDatos[j].ciudad
        const fechavisitaNew = nuevosDatos[j].fechaVisita
        const procesoNew = nuevosDatos[j].proceso
        const subprocesoNew = nuevosDatos[j].subproceso
        const elementoNew = nuevosDatos[j].elemento
        const numElementoNew = nuevosDatos[j].numElemento
        const sedeNew = nuevosDatos[j].sede
        const baseNew = nuevosDatos[j].base
        const concatNew = nuevosDatos[j].concat
        const conteoNew = nuevosDatos[j].conteo
        const puntajeNew = nuevosDatos[j].puntaje
        const califiacionNew = nuevosDatos[j].calificacion

        newDic[newKey] = {
          paisNew,
          cicloNew,
          codigoNew,
          concesionairoNew,
          ciudadNew,
          fechavisitaNew,
          procesoNew,
          subprocesoNew,
          elementoNew,
          numElementoNew,
          sedeNew,
          baseNew,
          concatNew,
          conteoNew,
          puntajeNew,
          califiacionNew
        } 
      } else {
        continue
      }
    }

    const newArray = Object.values(newDic);
//   console.log("newArray",newArray);
// console.log("newArray",newArray.length);

    const processDictionary = {
      descubrimiento: "1. DESCUBRIMIENTO",
      compra: "2. COMPRA",
      entrega: "3. ENTREGA",
      lealtad: "4. LEALTAD",
      habilitadores: "5. HABILITADORES"
    }


    let newData = {}
    for (const key in newDic) {
      let calificacionDic = newDic[key].califiacionNew 
      let procesoDic = newDic[key].procesoNew 
      let concesionarioDic = newDic[key].concesionairoNew 
      let sedeDic = newDic[key].sedeNew 
      let consede = `${concesionarioDic} - ${sedeDic}`
      let codigoDic = newDic[key].codigoDic
      let cicloDic = newDic[key].cicloNew
      let paisDic = newDic[key].paisNew
      let ciudadDic = newDic[key].ciudadNew



      let cal = calificacionDic == '100%' ? true : false
      let paso = cal == true ? 1 : 0
      let reprobo = cal == false ? 1 : 0

      if (!newData[consede] || newData[consede].proceso == '') {

        newData[consede] = {
          concesionario: concesionarioDic,
          sede: sedeDic,
          consede: consede,
          codigo: codigoDic,
          ciclo: cicloDic,
          pais: paisDic,
          ciudad: ciudadDic,
          total: null, 
          descubrimiento: {paso: 0, reprobo: 0},
          compra: {paso: 0, reprobo: 0}, 
          entrega: {paso: 0, reprobo: 0}, 
          lealtad: {paso: 0, reprobo: 0}, 
          habilitadores: {paso: 0, reprobo: 0}
        } 

        if (procesoDic == processDictionary.descubrimiento) {
          newData[consede].descubrimiento.paso += paso
          newData[consede].descubrimiento.reprobo += reprobo
        } else if (procesoDic == processDictionary.compra) {
          newData[consede].compra.paso += paso
          newData[consede].compra.reprobo += reprobo
        } else if (procesoDic == processDictionary.entrega) {
          newData[consede].entrega.paso += paso
          newData[consede].entrega.reprobo += reprobo
        } else if (procesoDic == processDictionary.lealtad) {
          newData[consede].lealtad.paso += paso
          newData[consede].lealtad.reprobo += reprobo
        } else if (procesoDic == processDictionary.habilitadores) {
          newData[consede].habilitadores.paso += paso
          newData[consede].habilitadores.reprobo += reprobo
        }
      } else {
        if (procesoDic == processDictionary.descubrimiento) {
          newData[consede].descubrimiento.paso += paso
          newData[consede].descubrimiento.reprobo += reprobo
        } else if (procesoDic == processDictionary.compra) {
          newData[consede].compra.paso += paso
          newData[consede].compra.reprobo += reprobo
        } else if (procesoDic == processDictionary.entrega) {
          newData[consede].entrega.paso += paso
          newData[consede].entrega.reprobo += reprobo
        } else if (procesoDic == processDictionary.lealtad) {
          newData[consede].lealtad.paso += paso
          newData[consede].lealtad.reprobo += reprobo
        } else if (procesoDic == processDictionary.habilitadores) {
          newData[consede].habilitadores.paso += paso
          newData[consede].habilitadores.reprobo += reprobo
        }
      }
    }
    
    // console.log('newData', newData)
    // console.log('newData length', Object.keys(newData).length)
    
    let interacciones = 0
    for (const dataKey in newData) {
      let motrarLogs = true
      interacciones += 1

      let totalOptDes = newData[dataKey].descubrimiento.paso + newData[dataKey].descubrimiento.reprobo 
      let totalOptCom = newData[dataKey].compra.paso + newData[dataKey].compra.reprobo 
      let totalOptEnt = newData[dataKey].entrega.paso + newData[dataKey].entrega.reprobo 
      let totalOptLea = newData[dataKey].lealtad.paso + newData[dataKey].lealtad.reprobo 
      let totalOptHab = newData[dataKey].habilitadores.paso + newData[dataKey].habilitadores.reprobo 

      let porcentajeDes = (newData[dataKey].descubrimiento.paso / totalOptDes) * 100
      let porcentajeCom = (newData[dataKey].compra.paso / totalOptCom) * 100
      let porcentajeEnt = (newData[dataKey].entrega.paso / totalOptEnt) * 100
      let porcentajeLea = (newData[dataKey].lealtad.paso / totalOptLea) * 100
      let porcentajeHab = (newData[dataKey].habilitadores.paso / totalOptHab) * 100 
      
      newData[dataKey].descubrimiento = porcentajeDes.toFixed(2)
      newData[dataKey].compra = porcentajeCom.toFixed(2)
      newData[dataKey].entrega = porcentajeEnt.toFixed(2)
      newData[dataKey].lealtad = porcentajeLea.toFixed(2)
      newData[dataKey].habilitadores = porcentajeHab.toFixed(2)

      let totalGen = (porcentajeDes + porcentajeCom + porcentajeEnt + porcentajeLea + porcentajeHab) / 5 

      newData[dataKey].total = totalGen

      if (motrarLogs) {
        console.log('Key', dataKey)
        console.log('Descubrimiento',  newData[dataKey].descubrimiento.paso)
        console.log('Compra', newData[dataKey].compra.paso)
        console.log('Entrega',  newData[dataKey].entrega.paso)
        console.log('Lealtad',  newData[dataKey].lealtad.paso)
        console.log('Habilitadores',  newData[dataKey].habilitadores.paso)

        motrarLogs = false
      }
    }


    const spreadsheetEf = SpreadsheetApp.openById(consId);
    const formatoSheetEf = spreadsheetEf.getSheetByName('Compra - EF');
    formatoSheetEf.clear();
    formatoSheetEf.getRange(1, 1, 1, 16).setValues([[
      'pais', 'ciclo', 'Código', 'Concesionario', 'Ciudad', 'FechaVisita', 'Proceso', 'Subproceso', 'Elemento', 'numElemento', 'Sede', 'Base', 'Concatenado', 'Conteo', 'Puntaje', 'Calificacion'
    ]]);
    
    if (newArray.length > 0) {
  // Convertir el array de objetos en un array de arrays
  const newArrayValues = newArray.map(obj => Object.values(obj));
  
  // Obtener el número de filas y columnas
  const numFilas = newArrayValues.length;
  const numColumnas = newArrayValues[0].length;
  
  // Escribir los datos en la hoja "Compra - EF"
  formatoSheetEf.getRange(2, 1, numFilas, numColumnas).setValues(newArrayValues);
} else {
  console.log("No se encontraron datos para agregar a la hoja 'Compra - EF'.");
}

// Limpiar la hoja "Formato" y agregar los nuevos datos
formatoSheet.clear();
formatoSheet.getRange(1, 1, 1, 18).setValues([[
  'País', 'Ciclo', 'Código', 'Concesionario', 'Ciudad', 'Fecha Visita', 'Pregunta', 'Respuesta', 'Proceso', 'Subproceso1', 'Elemento', '#Elemento', 'Sede', 'Base', 'Concatenado', 'Conteo', 'Puntaje', 'Calificacion'
]]);

if (nuevosDatos.length > 0) {
  // Convertir el array de objetos en un array de arrays
  const nuevosDatosArray = nuevosDatos.map(obj => Object.values(obj));
  
  // Escribir los datos en la hoja "Formato"
  formatoSheet.getRange(2, 1, nuevosDatosArray.length, nuevosDatosArray[0].length).setValues(nuevosDatosArray);
} else {
  console.log("No se encontraron datos para agregar a la hoja 'Formato'.");
}
});
  }
