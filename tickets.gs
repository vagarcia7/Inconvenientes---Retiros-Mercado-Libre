 // Esto que se hace acá es para poder extraer fecha y hora exacta
  var tiempoTranscurrido = Date.now()+(1000*60*60)*2
  var fecha = new Date(tiempoTranscurrido);
  var dia = fecha.getDate()
  var mes = fecha.getMonth()+1
  var agno = fecha.getFullYear()
  var fechaActual = dia + "/" + mes + "/" + agno 
  var fechaActualOk = fechaActual.split("/")
  var horaActual = fecha.getHours()
  var casillasEmails = ['problemsolverretiros@mercadolibre.com','miriam.leguizamon@mercadolibre.com','daniel.suarez@mercadolibre.com','valentin.garcia@mercadolibre.com']

  if(fechaActualOk[0].length == 1){
    let tmp = fechaActualOk[0]
    fechaActualOk[0] = 0 + tmp
  }

  if(fechaActualOk[1].length == 1){
    let tmp = fechaActualOk[1]
    fechaActualOk[1] = 0 + tmp
  }  

  var diaDeHoy = new Date(fechaActualOk[2],fechaActualOk[1],fechaActualOk[0])
  
  var horasTurnoManiana = [6,7,8,9,10,11,12,13]
  var horasTurnoTarde = [14,15,16,17,18,19,20,21]
  var horasTurnoNoche = [22,23,0,1,2,3,4,5,6]
  var diaDeSemana = fecha.getDay()
  
  // if para cuando es Sábado
  if(diaDeSemana == 6){
    horasTurnoManiana = [8,9,10,11,12,13,14,15]
    horasTurnoNoche = [22,23,0,1,2,3,4,5,6]
  }
  
  // if para cuando es Domingo
  if(diaDeSemana == 0){
    horasTurnoTarde = [8,9,10,11,12,13,14,15,16]
    horasTurnoNoche = [21,22,23,0,1,2,3,4,5]
  }


function emailTickets(){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('SEGUIMIENTO DE TICKETS'), true);
  var datos = spreadsheet.getRange('A2:D').getDisplayValues()

  for(i in datos){
    let rango = datos[i]
    let date = rango[0]
    let fechaX = date.split("/")
    let fechaUno = new Date(fechaX[2],fechaX[1],fechaX[0])
    let nroTicket = rango[1]
    let estado = rango[3]
    let diferenciaDeFechas = Math.floor((diaDeHoy - fechaUno) / (1000*60*60*24)) 
    
    let urlTicket = "<a href='https://mercadolibre.atlassian.net/servicedesk/customer/portal/93/" + nroTicket +"' class='button'>" + nroTicket + "</a>"
    if(estado == "PENDIENTE" && diferenciaDeFechas >= 1){
      for (email in casillasEmails){
        MailApp.sendEmail({
          to: casillasEmails[email],
          subject: "Inconvenientes y tickets | Mensaje Automático",
          htmlBody: "El ticket " + urlTicket + ", que fue cargado el " + date + " se encuentra según la planilla en estado " + "<span style='color: #ff0000;'>PENDIENTE</span>" + " hace " + diferenciaDeFechas + " días. En caso de que este no sea el estado correcto, hacer click " + "<a href='https://docs.google.com/spreadsheets/d/16_VLLpu03e7VkgGQX-4rhcJVS1Y4OwiyGUynk2O7KlY/edit#gid=1659804873&range=A2:D' class='button'><strong>aquí</strong></a>"
      })
     }
    }
  }
}



function emailCierreDeTurno(){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Cierre de turno PS'), true);
  var turnoOk = ""

  if (horasTurnoManiana.includes(horaActual)){
    turnoOk = "Turno Mañana"
  }

  if (horasTurnoTarde.includes(horaActual)){
    turnoOk = "Turno Tarde"
  }

  if (horasTurnoNoche.includes(horaActual)){
    turnoOk = "Turno Noche"
  }

  var asuntoDelEmail = "Cierre de turno PS | " + turnoOk + " | " + fechaActual
  var rangoDeDatos = spreadsheet.getRange('A3:B').getDisplayValues()
  var datosA = []
  for (i in rangoDeDatos){
    var datosB = rangoDeDatos[i]
    if(datosB[0] != ""){
      datosA.push(datosB)
    }
  }


  var mensajeDelEmail = "<h3>Listado de inconvenientes: </h3>" + 
    "<table class='table' style='width: 65%'><thead> <tr><th scope='col'>#</th><th scope='col'>Tipo de problema</th><th scope='col'>Detalle</th></tr></thead><tbody>"
    
  if(datosA.length == 0){
    mensajeDelEmail = "No se registraron inconvenientes en el turno"
  }

  else if (datosA.length >= 1){
    let counter = 1
    for (i in datosA){
      if (datosA[i] == datosA[datosA.length-1]){
        if (datosA[i][1] == ""){
          mensajeDelEmail += "<tr>" + "<th scope='row'>" + counter + "</th><td><center>" + datosA[i][0] + "</td></tr></tbody></table>"
        } else{
          mensajeDelEmail += "<tr>" + "<th scope='row'>" + counter + "</th><td><center>" + datosA[i][0] + "</center></td><td><center>" + datosA[i][1] + "</  td></tr></tbody></table>"
        }
      }
      else{
        if (datosA[i][1] == ""){
      mensajeDelEmail += "<tr>" + "<th scope='row'>" + counter + "</th><td><center>" + datosA[i][0] + "</center></td></tr>"
        }else{
      mensajeDelEmail += "<tr>" + "<th scope='row'>" + counter + "</th><td><center>" + datosA[i][0] + "</center></td><td><center>" + datosA[i][1] + "</center></td></tr>"
        }
      }
      counter++
    }
  }
  
 // Envío de email 
  for (email in casillasEmails){
    MailApp.sendEmail({
      to: casillasEmails[email],
      subject:asuntoDelEmail,
      htmlBody: mensajeDelEmail,
    })
  }

  SpreadsheetApp.getUi().alert("Email enviado satisfactoriamente ;)")

}
