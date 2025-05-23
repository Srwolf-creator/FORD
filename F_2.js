let nuevosDatos_s = [];

 function obtenerDatosYActualizar_s() {

  // IDs de las hojas de cálculo principales y su correspondiente base (P1, P2, etc.)
  const ids = [
    { id: '1dhG5ZzEW1pj1LOA57Js-FUsQpJia8-Z1S_1zu5zTk8U', base: 'S1' },
    { id: '1cd0Uc4X9GckQrC5sRocyLyeHfT7eIuEE7K_zNxjm7fM', base: 'S2' },
    { id: '16s0842xmccrUv-SzpAFfvc3CXKCYtGrnk9GJqcR0HR8', base: 'S3' },
    { id: '1Q2H03yvb0oBP-jWpTgKyESa12vCVdCerkk9hyqSKmA0', base: 'S4' },
    { id: '18MmqOg5HAU372gYq0VesfLpskb33DSnoSQgHYpv3Igg', base: 'S5' },
    { id: '1xwDuRb-KJMM5H0jXnl8XDLcBEgz9miekkOQESriyfoc', base: 'S6' },
    { id: '1cP5K1I3LOIUuT0SiAABgXpCxreM1p-oJR60O0fpi9aw', base: 'S7' },
    { id: '1eLmm59Gq9bpSaGLc3pqJZao2F0RznyhV2EYeExv-s2U', base: 'S8' }
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

          nuevosDatos_s.push({
            pais, ciclo, codigo, concesionario, ciudad, fechaVisita, pregunta, respuesta, proceso, subproceso, elemento, numElemento, sede, base, concat, conteo, puntaje, calificacion
        });
        }
      });
    });

    for (let index = 0; index < nuevosDatos_s.length; index++) {
      let key = nuevosDatos_s[index].concat
      if (dataConteo[key]) {
        let numConteo = dataConteo[key]
        let puntaje = dataResp[key]
        nuevosDatos_s[index].conteo = numConteo
        nuevosDatos_s[index].puntaje = puntaje

        if (numConteo !== puntaje) {
          nuevosDatos_s[index].calificacion = '0%'
        } else {
          nuevosDatos_s[index].calificacion = '100%'
        }
      }
    }
  

    console.log('nuevosDatos_s length', nuevosDatos_s.length)

    let newDic = {}
    for (let j = 0; j < nuevosDatos_s.length; j++) {
      console.log('nuevosDatos_s numElemento ' + j, nuevosDatos_s[j].numElemento)
      console.log('nuevosDatos_s concat ' + j, nuevosDatos_s[j].concat)


      if (!nuevosDatos_s[j].numElemento) continue
      if (!nuevosDatos_s[j].concat) continue

      let newKey = nuevosDatos_s[j].concat
      if (!newDic[newKey] || !newDic[newKey].concede) {
        const paisNew = nuevosDatos_s[j].pais
        const cicloNew = nuevosDatos_s[j].ciclo
        const codigoNew = nuevosDatos_s[j].codigo
        const concesionairoNew = nuevosDatos_s[j].concesionario
        const ciudadNew = nuevosDatos_s[j].ciudad
        const fechavisitaNew = nuevosDatos_s[j].fechaVisita
        const procesoNew = nuevosDatos_s[j].proceso
        const subprocesoNew = nuevosDatos_s[j].subproceso
        const elementoNew = nuevosDatos_s[j].elemento
        const numElementoNew = nuevosDatos_s[j].numElemento
        const sedeNew = nuevosDatos_s[j].sede
        const baseNew = nuevosDatos_s[j].base
        const concatNew = nuevosDatos_s[j].concat
        const conteoNew = nuevosDatos_s[j].conteo
        const puntajeNew = nuevosDatos_s[j].puntaje
        const califiacionNew = nuevosDatos_s[j].calificacion

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

    console.log("newDic", newDic);

    const newArray = Object.values(newDic);

    // console.log("newArray",newArray);
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
      let codigoDic = newDic[key].codigoNew
      let cicloDic = newDic[key].cicloNew
      let paisDic = newDic[key].paisNew
      let ciudadDic = newDic[key].ciudadNew



      let cal = calificacionDic == '100%' ? true : false
      let paso = cal == true ? 1 : 0
      let reprobo = cal == false ? 1 : 0

      if (!newData[consede] || newData[consede].proceso == '') {

        newData[consede] = {
          consede: consede,
          codigo: codigoDic,
          ciclo: cicloDic,
          pais: paisDic,
          ciudad: ciudadDic,
          concesionario: concesionarioDic,
          sede: sedeDic,
          nameSede: sedeDic,
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
    
    // console.log('newData length', Object.keys(newData).length)
    
    for (const dataKey in newData) {
      let totalOptDes = (parseInt(newData[dataKey].descubrimiento.paso) + parseInt(newData[dataKey].descubrimiento.reprobo)) || 0
      let totalOptCom = (parseInt(newData[dataKey].compra.paso) + parseInt(newData[dataKey].compra.reprobo)) || 0
      let totalOptEnt = (parseInt(newData[dataKey].entrega.paso) + parseInt(newData[dataKey].entrega.reprobo)) || 0
      let totalOptLea = (parseInt(newData[dataKey].lealtad.paso) + parseInt(newData[dataKey].lealtad.reprobo)) || 0
      let totalOptHab = (parseInt(newData[dataKey].habilitadores.paso) + parseInt(newData[dataKey].habilitadores.reprobo)) || 0 

      let porcentajeDes = totalOptDes > 0 ? (newData[dataKey].descubrimiento.paso / totalOptDes) * 100 : 0
      let porcentajeCom = totalOptCom > 0 ? (newData[dataKey].compra.paso / totalOptCom) * 100 : 0
      let porcentajeEnt = totalOptEnt > 0 ? (newData[dataKey].entrega.paso / totalOptEnt) * 100 : 0
      let porcentajeLea = totalOptLea > 0 ? (newData[dataKey].lealtad.paso / totalOptLea) * 100 : 0
      let porcentajeHab = totalOptHab > 0 ? (newData[dataKey].habilitadores.paso / totalOptHab) * 100 : 0
      
      newData[dataKey].descubrimiento = `${porcentajeDes.toFixed(2)}%`
      newData[dataKey].compra = `${porcentajeCom.toFixed(2)}%`
      newData[dataKey].entrega = `${porcentajeEnt.toFixed(2)}%`
      newData[dataKey].lealtad = `${porcentajeLea.toFixed(2)}%`
      newData[dataKey].habilitadores = `${porcentajeHab.toFixed(2)}%`

      let totalGen = (porcentajeDes + porcentajeCom + porcentajeEnt + porcentajeLea + porcentajeHab) / 5 

      newData[dataKey].total = `${totalGen.toFixed(2)}%`
    }

    // console.log('newData', newData)
    const newArrayData = Object.values(newData)


    const spreadsheetEf = SpreadsheetApp.openById(consId);
    const formatoSheetEf = spreadsheetEf.getSheetByName('Servicio - EF');
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
      console.log("No se encontraron datos para agregar a la hoja 'Servicio - EF'.");
    }

    const formatoSheetDg = spreadsheetEf.getSheetByName('Diagnostico - Manual Servicio');
    formatoSheetDg.clear();
    formatoSheetDg.getRange(1, 1, 1, 14).setValues([[
      'Concesionario - sede', 'Código','Ciclo', 'País', 'Ciudad', 'Concesionario', 'Sede', 'Nombre Sede', 'Total', '1. Descubrimiento', '2. Compra', '3. Entrega', '4. Lealtad', '5. Habilitadores'
    ]]);
    
    if (newArrayData.length > 0) {
      // Convertir el array de objetos en un array de arrays
      const newArrayValuesDg = newArrayData.map(obj => Object.values(obj));
      
      // Obtener el número de filas y columnas
      const numFilas = newArrayValuesDg.length;
      const numColumnas = newArrayValuesDg[0].length;
      
      // Escribir los datos en la hoja "Compra - EF"
      formatoSheetDg.getRange(2, 1, numFilas, numColumnas).setValues(newArrayValuesDg);
    } else {
      console.log("No se encontraron datos para agregar a la hoja Diagnostico - Manual Servicio'.");
    }

// Limpiar la hoja "Formato" y agregar los nuevos datos
formatoSheet.clear();
formatoSheet.getRange(1, 1, 1, 18).setValues([[
  'País', 'Ciclo', 'Código', 'Concesionario', 'Ciudad', 'Fecha Visita', 'Pregunta', 'Respuesta', 'Proceso', 'Subproceso1', 'Elemento', '#Elemento', 'Sede', 'Base', 'Concatenado', 'Conteo', 'Puntaje', 'Calificacion'
]]);

if (nuevosDatos_s.length > 0) {
  // Convertir el array de objetos en un array de arrays
  const nuevosDatos_sArray = nuevosDatos_s.map(obj => Object.values(obj));
  
  // Escribir los datos en la hoja "Formato"
  formatoSheet.getRange(2, 1, nuevosDatos_sArray.length, nuevosDatos_sArray[0].length).setValues(nuevosDatos_sArray);
} else {
  console.log("No se encontraron datos para agregar a la hoja 'Formato'.");
}
});
  } 
