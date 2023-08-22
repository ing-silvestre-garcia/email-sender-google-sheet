function myFunction() {
  // Obtener la hoja de cálculo activa
  var hojaMaritimo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FRAVA-Maritimo");
  // Obtener la hoja de cálculo activa
  var hojaAereo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FRAVA-Aereo");
  // Obtener la hoja de cálculo activa
  var hojaPaqueteria = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FRAVA-Paquetería");

  operacionHoja(hojaMaritimo,"Maritimo");
  operacionHoja(hojaAereo,"Aereo");
  operacionHoja(hojaPaqueteria,"Paqueteria");
}

function operacionHoja(hoja,nombreHoja){
  Logger.log("Operacion de hoja " + nombreHoja );
  var data = hoja.getDataRange().getValues();

  // Fecha actual (sin la hora, solo fecha)
  var fechaActual = new Date();
  fechaActual.setHours(0, 0, 0, 0); // Establecer la hora a 00:00:00

  // Obtener la hoja "Datos de contacto"
  var hojaContacto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos de contacto");

  // Obtener la hoja "Panel Control"
  var hojaPanelControl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Panel Control");
  
  var campos = {};

  var numColumnas = data[5].length;
  for (var x = 0; x < numColumnas; x++) { //empezamos en 0 por ser culimna A y 4 porque son 4 columnas en la hoja de panel de control
    var campoEncabezado = data[5][x];//queda fijo 5 porque es el numero de renglon de los encabezados

    var filaCampo = buscarEnHojaPanelControl(hojaPanelControl, campoEncabezado); 

    if (filaCampo !== -1) {
      // Obtener el valor de la columna D en "Panel Control"
      var valorColumnaD = hojaPanelControl.getRange("D" + filaCampo).getValue();

      if (valorColumnaD === "Si") {
        // Obtener los valores de las columnas B y C en "Panel Control"
        var condicion = hojaPanelControl.getRange("B" + filaCampo).getValue();

        var valor = hojaPanelControl.getRange("C" + filaCampo).getValue();

        // Verificar si el campoEncabezado existe en "campos"
        if (!campos.hasOwnProperty(campoEncabezado)) {
          campos[x] = {}; // Inicializar el campo como un objeto vacío
        }

        campos[x]['condicion'] = condicion;
        campos[x]['valor'] = valor;
        campos[x]['nombre'] = campoEncabezado;
      }
    }
  }

  Logger.log("campos " + JSON.stringify(campos));

  // Recorrer los datos
  for (var i = 6; i < data.length; i++) { // Empezar en 6 para omitir la fila de encabezado

    var idValor = data[i][0]; // Columna que contiene el id del cliente

    var filaDestinatario = buscarDestinatarioEnHoja(hojaContacto, idValor); //Saber si ese cliente tiene destinatario
    if (filaDestinatario !== -1) {
      // Obtener los valores de la columna B (nombre) y C (correo)
      var nombre = hojaContacto.getRange("B" + filaDestinatario).getValue();
      var correo = hojaContacto.getRange("C" + filaDestinatario).getValue();

      buscarCampoCondicion(campos,data,i,nombre,correo);
    } else {
      // Registro de que no se encontró el destinatario en la hoja "Datos de contacto"
      Logger.log("No se encontró información de contacto para el destinatario: " + idValor);
    }

    switch (nombreHoja) {
      case "Maritimo":
        var fechaEtd = data[i][17]; // Columna que contiene la fecha
        var fechaEta = data[i][18]; // Columna que contiene la fecha
        break;
      case "Aereo":
        var fechaEtd = data[i][15]; // Columna que contiene la fecha
        var fechaEta = data[i][16]; // Columna que contiene la fecha
        break;
      case "Paqueteria":
        var fechaEtd = data[i][19]; // Columna que contiene la fecha
        var fechaEta = data[i][20]; // Columna que contiene la fecha
        break;
    }

    fechaEtd.setHours(0, 0, 0, 0);
    fechaEta.setHours(0, 0, 0, 0);

    // Verificar si la fecha en la casilla es futura con respecto al día actual
    Logger.log("Fecha ETD " + (fechaEtd - fechaActual) + " === " + (24 * 60 * 60 * 1000));
    if (fechaEtd instanceof Date && (fechaEtd - fechaActual) === (24 * 60 * 60 * 1000)) {
      // Buscar el destinatario en la hoja "Datos de contacto"
      

      if (filaDestinatario !== -1) {
        
        // Obtener los valores de la columna B (nombre) y C (correo)
        var nombre = hojaContacto.getRange("B" + filaDestinatario).getValue();
        var correo = hojaContacto.getRange("C" + filaDestinatario).getValue();

        // Construir el mensaje del correo
        var mensaje = "Hola " + nombre + ", este es un correo de ejemplo decir que la fecha ETD esta proxima: " + fechaEtd;

        // Enviar el correo
        enviarCorreo(correo, 'Fecha ETD', mensaje);
        
        // Registro del envío del correo (opcional)
        Logger.log("Se envió un correo a " + nombre + " ("+ correo +") para la fecha " + fechaEtd);
      } else {
        // Registro de que no se encontró el destinatario en la hoja "Datos de contacto"
        Logger.log("No se encontró información de contacto para el destinatario: " + idValor);
      }
    } else {
      // Registro de que no se envió el correo (opcional)
      Logger.log("No se envió un correo a " + nombre + " porque la fecha " + fechaEtd + " no es futura.");
    }

    // Verificar si la fecha en la casilla es futura con respecto al día actual
    if (fechaEta instanceof Date && (fechaEta - fechaActual) === (24 * 60 * 60 * 1000)) {
      // Construir el mensaje del correo
      var mensaje = "Hola " + nombre + ", este es un correo de ejemplo decir que la fecha ETA esta proxima: " + fechaEta;

      // Enviar el correo
      enviarCorreo(correo, 'Fecha ETA', mensaje);
      
      // Registro del envío del correo (opcional)
      Logger.log("Se envió un correo a " + nombre + " para la fecha " + fechaEta);
    } else {
      // Registro de que no se envió el correo (opcional)
      Logger.log("No se envió un correo a " + nombre + " porque la fecha " + fechaEta + " no es futura.");
    }
  }
}

function buscarCampoCondicion(campos,data,renglon,nombre,correo) {
  // Obtener un array de las claves del objeto campos
  var claves = Object.keys(campos);

  // Recorrer el array de claves usando un bucle forEach
  claves.forEach(function(indice) {
    console.log("Índice:", indice);
    console.log("Valor:", campos[indice]);

    if(data[renglon][indice]){
      var enviaCorreo = 0;
      var template = '';
      var valorCampo = data[renglon][indice];
      Logger.log("valorCampo " + valorCampo);
      switch (campos[indice]['condicion']) {
        case 'Igual':
          if(campos[indice]['valor'] == valorCampo) {
            enviaCorreo = 1;
            template = campos[indice]['nombre']
          } 
          break;
        case 'Diferente':
          if(campos[indice]['valor'] != valorCampo) {
            enviaCorreo = 1;
            template = campos[indice]['nombre']
          }
          break;
        case 'Mayor':
          if(valorCampo > campos[indice]['valor']) {
            enviaCorreo = 1;
            template = campos[indice]['nombre']
          }
          break;
        case 'Menor':
          if(valorCampo < campos[indice]['valor']) {
            enviaCorreo = 1;
            template = campos[indice]['nombre']
          }
          break;
        default:
          enviaCorreo = 0;
      }
      if(enviaCorreo==1) {
        // Enviar el correo
        if(template!=''){
          var valoresEncontrados = buscarEnTemplates(template);
          if (valoresEncontrados) {
            var asunto = reemplazarTextoTemplate(valoresEncontrados[0],"__CLIENTE__", nombre);
            var cuerpo = reemplazarTextoTemplate(valoresEncontrados[1],"__CLIENTE__", nombre);
            
          } else {
            var cuerpo = mensaje
            var asunto = titulo
          }
        }
        enviarCorreo(correo, asunto, cuerpo);
      }
    }
  });
}

// Función para buscar el destinatario en la hoja "Datos de contacto"
function buscarDestinatarioEnHoja(hoja, destinatario) {
  var dataContacto = hoja.getRange("A2:C" + hoja.getLastRow()).getValues();

  for (var i = 0; i < dataContacto.length; i++) {
    if (dataContacto[i][0] === destinatario) {
      return i + 2; // Devolver el número de fila + 2 para tener en cuenta el desplazamiento de la fila 2 (A2:C2)
    }
  }

  return -1; // Devolver -1 si no se encontró el destinatario
}

// Función para buscar el destinatario en la hoja "Panel Control"
function buscarEnHojaPanelControl(hoja, campo) {
  var dataContacto = hoja.getRange("A2:D" + hoja.getLastRow()).getValues();

  for (var i = 0; i < dataContacto.length; i++) {
    if (dataContacto[i][0] === campo) {
      return i + 2; // Devolver el número de fila + 2 para tener en cuenta el desplazamiento de la fila 2 (A2:C2)
    }
  }

  return -1; // Devolver -1 si no se encontró el destinatario
}

function enviarCorreo(destinatario,asunto,mensaje) {
  
  // Enviar el correo
    GmailApp.sendEmail(destinatario, asunto, mensaje);
}

function reemplazarTextoTemplate(textoTemplate,etiqueta, valor) {
  // Utilizar el método replace() para buscar y reemplazar la cadena "__CLIENTE__" por el valor de la variable cliente
  var resultado = textoTemplate.replace(etiqueta, valor);
  return resultado;
}

function buscarEnTemplates(cadenaBusqueda) {
  // Obtener la hoja "Templates"
  var hojaTemplates = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Templates");

  // Obtener los datos de la hoja "Templates"
  var dataTemplates = hojaTemplates.getDataRange().getValues();

  // Buscar la cadena en la columna A y obtener los valores de las columnas B y C
  for (var i = 0; i < dataTemplates.length; i++) {
    var valorColumnaA = dataTemplates[i][0];

    if (valorColumnaA === cadenaBusqueda) {
      var valorColumnaB = dataTemplates[i][1];
      var valorColumnaC = dataTemplates[i][2];
      // Hacer lo que necesites con los valores encontrados (valorColumnaB y valorColumnaC)
      return [valorColumnaB, valorColumnaC]; // Retorna un array con los valores de la columna B y C
    }
  }

  // Si no se encontró la cadena, puedes retornar un valor por defecto o null
  return null;
}

