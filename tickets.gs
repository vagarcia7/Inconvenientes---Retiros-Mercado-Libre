  var fecha = new Date();
  var dia = fecha.getUTCDate()
  var mes = fecha.getMonth()+1
  var agno = fecha.getFullYear()
  var fechaActual = dia + "/" + mes + "/" + agno 
  var fechaActualOk = fechaActual.split("/")

  if(fechaActualOk[0].length == 1){
    let tmp = fechaActualOk[0]
    fechaActualOk[0] = 0 + tmp
  }

  if(fechaActualOk[1].length == 1){
    let tmp = fechaActualOk[1]
    fechaActualOk[1] = 0 + tmp
  }  

  var diaDeHoy = new Date(fechaActualOk[2],fechaActualOk[1],fechaActualOk[0])

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
    let diferenciaDeFechas = Math.floor((diaDeHoy - fechaUno) / (1000*60*60*24) + 3) 
    let urlTicket = "<a href='https://mercadolibre.atlassian.net/servicedesk/customer/portal/93/" + nroTicket +"' class='button'>" + nroTicket + "</a>"
    if(estado == "PENDIENTE" && diferenciaDeFechas >= 1){
      MailApp.sendEmail({
        to: "valentin.garcia@mercadolibre.com",
        subject: "Inconvenientes y tickets | Mensaje Automático",
        htmlBody: "El ticket " + urlTicket + ", que fue cargado el " + date + " se encuentra según la planilla en estado " + "<span style='color: #ff0000;'>PENDIENTE</span>" + " hace " + diferenciaDeFechas + " días. En caso de que este no sea el estado correcto, hacer click " + "<a href='https://docs.google.com/spreadsheets/d/16_VLLpu03e7VkgGQX-4rhcJVS1Y4OwiyGUynk2O7KlY/edit#gid=1659804873&range=A2:D' class='button'><strong>aquí</strong></a>"
      })
    }
  }
}



function emailCierreDeTurno(){
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Cierre de turno PS (En pruebas)'), true);
  var turno = spreadsheet.getRange('A3').getDisplayValue()
  var turnoOk = ""

  if(turno == "TM"){
    turnoOk = "Turno Mañana"
  }
  else if(turno == "TT"){
    turnoOk = "Turno Tarde"
  }
  else if(turno == "TN"){
    turnoOk = "Turno Noche" 
  }

  var asuntoDelEmail = "Cierre de turno PS | " + turnoOk + " | " + fechaActualOk
  var rangoDeDatos = spreadsheet.getRange('B3:C').getDisplayValues()
  var datosA = []
  for (i in rangoDeDatos){
    var datosB = rangoDeDatos[i]
    if(datosB[0] != "" && datosB[1] != ""){
      datosA.push(datosB)
    }
  }

  var mensajeDelEmail = ""
  if(datosA.length == 0){
    mensajeDelEmail = "No se registraron inconvenientes en el turno"
  }
  else if(datosA.length == 1){
    mensajeDelEmail = "Listado de inconvenientes:" + "<br><ul><li>" + datosA[0][0] + " (" + datosA[0][1] + ")" + "</li></ul>"
  }
  else if(datosA.length == 2){
    mensajeDelEmail = "Listado de inconvenientes:" + "<br><ul><li>" + datosA[0][0] + " (" + datosA[0][1] + ")" + "</li><li>" + datosA[1][0] + " (" + datosA[1][1] + ")" + "</li></ul>"
  }
  else if(datosA.length == 3){
    mensajeDelEmail = "Listado de inconvenientes:" + "<br><ul><li>" + datosA[0][0] + " (" + datosA[0][1] + ")" + "</li><li>" + datosA[1][0] + " (" + datosA[1][1] + ")" + "</li><li>" + datosA[2][0] + " (" + datosA[2][1] + ")" + "</li></ul>"
  }
  else if(datosA.length == 4){
    mensajeDelEmail = "Listado de inconvenientes:" + "<br><ul><li>" + datosA[0][0] + " (" + datosA[0][1] + ")"  + "</li><li>" + datosA[1][0] + " (" + datosA[1][1] + ")" + "</li><li>" + datosA[2][0] + " (" + datosA[2][1] + ")" + "</li><li>" + datosA[3][0] + " (" + datosA[3][1] + ")" + "</li></ul>"
  }
  else if(datosA.length == 5){
    mensajeDelEmail = "Listado de inconvenientes:" + "<br><ul><li>" + "" + "" + datosA[0][0] + " (" + datosA[0][1] + ")" + "</li><li>" + datosA[1][0] + " (" + datosA[1][1] + ")" + "</li><li>" + datosA[2][0] + " (" + datosA[2][1] + ")" + "</li><li>" + datosA[3][0] + " (" + datosA[3][1] + ")" + "</li><li>" + datosA[4][0] + " (" + datosA[4][1] + ")" + "</li></ul>"
  }

  MailApp.sendEmail({
    to: 'valentin.garcia@mercadolibre.com',
    subject:asuntoDelEmail,
    htmlBody: mensajeDelEmail,
  })

  SpreadsheetApp.getUi().alert("Email enviado satisfactoriamente ;)")
}
