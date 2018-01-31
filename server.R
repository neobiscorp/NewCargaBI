##################SETUP##############################
#Load Libraries
library(shiny)      #Dashboard
library(RMySQL)     #Secundary MySQL connection, Kill connections
library(DBI)        #Primary MySQL connection
library(openxlsx)   #Read xlsx files
library(data.table) #For merge, rbind and dataframe works


#comentario vacío
#Function to kill MySQL connections
killDbConnections <- function () {
  all_cons <- dbListConnections(MySQL())
  print(all_cons)
  for (con in all_cons)
    +  dbDisconnect(con)
  print(paste(length(all_cons), " connections killed."))
}
#Function to return the check.names = TRUE back to normal
fixnames<-function(x){
  a<-names(x)
  d<-c()
  for (i in 1:length(a)){
    b<-unlist(strsplit(a[i], "[.]"))
    c<-paste(b,collapse=" ")
    d[i]<-c
  }
  names(x)<-d
  return(x)
}


#Increase the maxium size of an uploaded file to 250mb
options(shiny.maxRequestSize = 250 * 1024 ^ 2)

#Function to convert ITEM months to number
y <-
  c(
    "ene. " = "01",
    "feb. " = "02",
    "mar. " = "03",
    "abr. " = "04",
    "may. " = "05",
    "jun. " = "06",
    "jul. " = "07",
    "ago. " = "08",
    "sept. " = "09",
    "oct. " = "10",
    "nov. " = "11",
    "dic. " = "12"
  )

#############Start Shiny Server With session############
shinyServer(function(input, output, session) {
  #Run function after pressing execute button
  observeEvent(input$execute, {
    updateButton(session, "execute", style = "warning")
    #name input variables
    client <- input$Client
    export <- input$Export
    plantilla <- input$plantilla
    tipos <- input$tipos
    presupuesto <- input$presupuesto
    cuentas <- input$cuentas
    contrato <- input$contrato
    proveedor <-input$proveedor
    usuarioid <-input$usuarioid
    customfield <- input$customfield
    divisa<- input$divisa
    proveedor<<-proveedor
    usuarioid<<-usuarioid
    customfield<<-customfield
    divisa<<-divisa
    #Change name of client input to the same name of the database
    {

    if (client == "Licitacion Movil (Entel y Movistar)") {
      client <- "lmovil"
    }
    if (client == "Informe Gestion Telefonia") {
      client <- "igm"
    }
    if (client == "Anomalias de Facturacion Movil") {
      client <- "afm"
    }
    if (client=="Informe Gestion Impresion"){
      client <- "igi"
    }
    if (client == "Informe Gestion Enlace"){
      client <- "ige"
    }
    if (client == "tester"){
      client <- "test"
    }
      client<<-client
    }
    
    
    #Create DB connection variable
    DB <- dbConnect(
      MySQL(),
      user = "root",
      password = "",
      dbname = paste0(client)
    )

    #Run this function in case theres a file in the Factura upload
    if (!is.null(input$factura)) {
      dbGetQuery(DB, "SET NAMES 'utf8';")
      #Read the csv of the factura
      facturas<<-read.csv2(input$factura[['datapath']],encoding = "UTF-8",check.names = FALSE,header = TRUE)
      #Delete the Estado Column
      a<-names(facturas)
      a[1]<-"Factura"
      names(facturas)<-a
      facturas[, 'Estado'] <- NULL
      
      #Rename the columns of the file
      names(facturas)[names(facturas) == 'Mes de facturación'] <- 'Mes'
      
      #Convert the ITEM months to YYYY/MM/01 "Número de factura"
      for (i in 1:12) {
        facturas['Mes'] <-
          lapply(facturas['Mes'], function(x)
            gsub(names(y[i]), y[[i]], x))
      }
      facturas['Fecha'] <-
        lapply(facturas['Mes'], function(x)
          paste(substr(x , 3 , 6),
                substr(x , 1 , 2),
                "01",
                sep = "/"))
      #Delete the Column MES that its not needed now
      facturas[, 'Mes'] <- NULL
      facturas <<- facturas
      
      #If the client is Parque Arauco, run the following
      if (client == "afm"){
        facturas[,'Fecha de facturación']<-NULL
      facturas[,'Importado el']<-NULL
      facturas[,'Etiqueta centro de facturación']<-NULL
      dbWriteTable(
        DB,
        "facturas",
        facturas,
        field.types = NULL,
        row.names = FALSE,
        overwrite = TRUE,
        append = FALSE,
        allow.keywords = FALSE
      )
        facturas<<-facturas
        }
      else{
        
        dbWriteTable(
          DB,
          "facturas",
          facturas,
          field.types = NULL,
          row.names = FALSE,
          overwrite = TRUE,
          append = FALSE,
          allow.keywords = FALSE
        )
      }
    }
    #Run the following if the user uploads a file to the usos file input
    if (!is.null(input$usos)) {
      #Clear and create variable and read CSV file
      dbGetQuery(DB, "SET NAMES 'utf8';")
      dataFilesUF <<- NULL
      #dataFilesUF <<- lapply(input$usos[['datapath']], read.csv2)
     
      #CAMBIO PENDIENTE, EN CONJUNTO CAMBIAR LA LECTURA DE TITULOS EN TODOS LOS IF DE ADELANTE EN CONSECUENCIA
      dataFilesUF <<- lapply(input$usos[['datapath']], function(x) read.csv2(x,encoding = "UTF-8",check.names = FALSE,header = TRUE))
      #Append all usos files to the same dataframe
      
      uso <<- rbindlist(dataFilesUF,fill = TRUE)
      #uso <<- rbindlist(dataFilesUF)
    
      a<-as.character(names(uso))
      d<-matrix(nrow = length(a),ncol = 1)
      for(i in 1:length(a)){
        c<-substr((a[i]),3,10)
        if(c=="cceso"){
          print("Encontrado")
          d[i,1]<-"Acceso"
        }
        else if (c=="ombre"){
          print("Encontrado 2")
          d[i,1]<-"Nombre"
      }
      else{
        d[i,1]<-a[i]
      }
      }
      b<-as.character(d)
      colnames(uso)<-b
     
      rm(a,b,c,d)
      uso<<-uso
      
      uso[["Período de2"]]<-uso[["Período de"]]
      
      #Convert the ITEM months to YYYY/MM/01"
      for (k in 1:12) {
        uso[, 'Período de2']<-
          lapply(uso[, 'Período de2'], function(x)
            gsub(names(y[k]), y[[k]], x))
      }
      
        uso[, 'Fecha'] <-
          lapply(uso[, 'Período de2'], function(x)
            paste(substr(x , 3 , 6),
                  substr(x , 1 , 2),
                  "01",
                  sep = "/"))
        uso[["Período de2"]]<-NULL
        
        uso[,'Acceso fix'] <-
          lapply(uso[,'Acceso'], function(x)
            substring(x, 3))
      uso[["Fecha"]]<- as.Date(uso[["Fecha"]])
        uso<<-uso
       
 ########CAMBIOS A USO POR SELECCION DE DIVISA Y POR TIPO DE INFORME #######
        
        #Se cambian los nombres de todas las columnas para que queden parametricas al ingresar a bi
        if(divisa == "CLP"){
          names(uso)[names(uso) == 'Total (CLP)'] <- 'Total'
          names(uso)[names(uso) == 'Plano tarifario (CLP)'] <- 'Plano tarifario'
          names(uso)[names(uso) == 'Plano tarifario (rebajado) (CLP)'] <- 'Plano tarifario (rebajado)'
          names(uso)[names(uso) == 'Servicios (CLP)'] <- 'Servicios'
          names(uso)[names(uso) == 'Servicios > Opciones (CLP)'] <- 'Servicios > Opciones'
          names(uso)[names(uso) == 'Servicios > Puesta en servicio (CLP)'] <- 'Servicios > Puesta en servicio'
          names(uso)[names(uso) == 'Servicios > Materiales (CLP)'] <- 'Servicios > Materiales'
          names(uso)[names(uso) == 'Servicios > Otros (CLP)'] <- 'Servicios > Otros'
          names(uso)[names(uso) == 'Descuentos (CLP)'] <- 'Descuentos'
          names(uso)[names(uso) == 'Descuento > Plano tarifario (CLP)'] <- 'Descuento > Plano tarifario'
          names(uso)[names(uso) == 'Descuento > Opciones (CLP)'] <- 'Descuento > Opciones'
          names(uso)[names(uso) == 'Descuentos > Uso (CLP)'] <- 'Descuentos > Uso'
          names(uso)[names(uso) == 'Descuento > Refacturación (CLP)'] <- 'Descuento > Refacturación'
          names(uso)[names(uso) == 'Descuento > Otros (CLP)'] <- 'Descuento > Otros'
          names(uso)[names(uso) == 'Uso (CLP)'] <- 'Uso'
          names(uso)[names(uso) == 'Uso (rebajado) (CLP)'] <- 'Uso (rebajado)'
          names(uso)[names(uso) == 'Voz (CLP)'] <- 'Voz'
          names(uso)[names(uso) == 'Voz nacional (CLP)'] <- 'Voz nacional'
          names(uso)[names(uso) == 'Voz inter. (CLP)'] <- 'Voz inter.'
          names(uso)[names(uso) == 'Voz roaming (CLP)'] <- 'Voz roaming'
          names(uso)[names(uso) == 'Voz hacia Int (CLP)'] <- 'Voz hacia Int'
          names(uso)[names(uso) == 'Voz nac. estánd. (CLP)'] <- 'Voz nac. estánd.'
          names(uso)[names(uso) == 'Voz nac. espec. (CLP)'] <- 'Voz nac. espec.'
          names(uso)[names(uso) == 'Voz nac. VPN (CLP)'] <- 'Voz nac. VPN'
          names(uso)[names(uso) == 'Voz Roaming In (CLP)'] <- 'Voz Roaming In'
          names(uso)[names(uso) == 'Voz Roaming Out (CLP)'] <- 'Voz Roaming Out'
          names(uso)[names(uso) == 'Voz Roaming desc. (CLP)'] <- 'Voz Roaming desc.'
          names(uso)[names(uso) == 'Datos (CLP)'] <- 'Datos'
          names(uso)[names(uso) == 'Datos inter. (CLP)'] <- 'Datos inter.'
          names(uso)[names(uso) == 'Datos nac. (CLP)'] <- 'Datos nac.'
          names(uso)[names(uso) == 'Datos espec. (CLP)'] <- 'Datos espec.'
          names(uso)[names(uso) == 'SMS (CLP)'] <- 'SMS'
          names(uso)[names(uso) == 'SMS nac. (CLP)'] <- 'SMS nac.'
          names(uso)[names(uso) == 'SMS inter. (CLP)'] <- 'SMS inter.'
          names(uso)[names(uso) == 'SMS espec. (CLP)'] <- 'SMS espec.'
          names(uso)[names(uso) == 'SMS hacia inter. (CLP)'] <- 'SMS hacia inter.'
          names(uso)[names(uso) == 'SMS roaming (CLP)'] <- 'SMS roaming'
          names(uso)[names(uso) == 'MMS (CLP)'] <- 'MMS'
          names(uso)[names(uso) == 'MMS nac. (CLP)'] <- 'MMS nac.'
          names(uso)[names(uso) == 'MMS inter. (CLP)'] <- 'MMS inter.'
          names(uso)[names(uso) == 'MMS espec. (CLP)'] <- 'MMS espec.'
          names(uso)[names(uso) == 'MMS hacia inter. (CLP)'] <- 'MMS hacia inter.'
          names(uso)[names(uso) == 'MMS roaming (CLP)'] <- 'MMS roaming'
          names(uso)[names(uso) == 'SMS/MMS (CLP)'] <- 'SMS/MMS'
          names(uso)[names(uso) == 'SMS/MMS nac. (CLP)'] <- 'SMS/MMS nac.'
          names(uso)[names(uso) == 'SMS/MMS inter. (CLP)'] <- 'SMS/MMS inter.'
          names(uso)[names(uso) == 'SMS/MMS espec. (CLP)'] <- 'SMS/MMS espec.'
          names(uso)[names(uso) == 'SMS/MMS hacia inter. (CLP)'] <- 'SMS/MMS hacia inter.'
          names(uso)[names(uso) == 'SMS/MMS roaming (CLP)'] <- 'SMS/MMS roaming'
          names(uso)[names(uso) == 'Business Hours (CLP)'] <- 'Business Hours'
          names(uso)[names(uso) == 'Private (CLP)'] <- 'Private'
          names(uso)[names(uso) == 'Horario laboral (CLP)'] <- 'Horario laboral'
          names(uso)[names(uso) == 'Horario privado (CLP)'] <- 'Horario privado'
          names(uso)[names(uso) == 'Exterior (CLP)'] <- 'Exterior'
          names(uso)[names(uso) == 'Interior (CLP)'] <- 'Interior'
          names(uso)[names(uso) == 'Copias B/N (CLP)'] <- 'Copias B/N'
          names(uso)[names(uso) == 'Copias Color (CLP)'] <- 'Copias Color'
          names(uso)[names(uso) == 'Copias (CLP)'] <- 'Copias'
          names(uso)[names(uso) == 'Minimo Copias (CLP)'] <- 'Minimo Copias'
          names(uso)[names(uso) == 'Minimo Copias B/N (CLP)'] <- 'Minimo Copias B/N'
          names(uso)[names(uso) == 'Minimo Copias Color (CLP)'] <- 'Minimo Copias Color'
          names(uso)[names(uso) == 'IVA (CLP)'] <- 'IVA'
          }
        else if (divisa == "UF"){
          names(uso)[names(uso) == 'Total (UF)'] <- 'Total'
          names(uso)[names(uso) == 'Plano tarifario (UF)'] <- 'Plano tarifario'
          names(uso)[names(uso) == 'Plano tarifario (rebajado) (UF)'] <- 'Plano tarifario (rebajado)'
          names(uso)[names(uso) == 'Servicios (UF)'] <- 'Servicios'
          names(uso)[names(uso) == 'Servicios > Opciones (UF)'] <- 'Servicios > Opciones'
          names(uso)[names(uso) == 'Servicios > Puesta en servicio (UF)'] <- 'Servicios > Puesta en servicio'
          names(uso)[names(uso) == 'Servicios > Materiales (UF)'] <- 'Servicios > Materiales'
          names(uso)[names(uso) == 'Servicios > Otros (UF)'] <- 'Servicios > Otros'
          names(uso)[names(uso) == 'Descuentos (UF)'] <- 'Descuentos'
          names(uso)[names(uso) == 'Descuento > Plano tarifario (UF)'] <- 'Descuento > Plano tarifario'
          names(uso)[names(uso) == 'Descuento > Opciones (UF)'] <- 'Descuento > Opciones'
          names(uso)[names(uso) == 'Descuentos > Uso (UF)'] <- 'Descuentos > Uso'
          names(uso)[names(uso) == 'Descuento > Refacturación (UF)'] <- 'Descuento > Refacturación'
          names(uso)[names(uso) == 'Descuento > Otros (UF)'] <- 'Descuento > Otros'
          names(uso)[names(uso) == 'Uso (UF)'] <- 'Uso'
          names(uso)[names(uso) == 'Uso (rebajado) (UF)'] <- 'Uso (rebajado)'
          names(uso)[names(uso) == 'Voz (UF)'] <- 'Voz'
          names(uso)[names(uso) == 'Voz nacional (UF)'] <- 'Voz nacional'
          names(uso)[names(uso) == 'Voz inter. (UF)'] <- 'Voz inter.'
          names(uso)[names(uso) == 'Voz roaming (UF)'] <- 'Voz roaming'
          names(uso)[names(uso) == 'Voz hacia Int (UF)'] <- 'Voz hacia Int'
          names(uso)[names(uso) == 'Voz nac. estánd. (UF)'] <- 'Voz nac. estánd.'
          names(uso)[names(uso) == 'Voz nac. espec. (UF)'] <- 'Voz nac. espec.'
          names(uso)[names(uso) == 'Voz nac. VPN (UF)'] <- 'Voz nac. VPN'
          names(uso)[names(uso) == 'Voz Roaming In (UF)'] <- 'Voz Roaming In'
          names(uso)[names(uso) == 'Voz Roaming Out (UF)'] <- 'Voz Roaming Out'
          names(uso)[names(uso) == 'Voz Roaming desc. (UF)'] <- 'Voz Roaming desc.'
          names(uso)[names(uso) == 'Datos (UF)'] <- 'Datos'
          names(uso)[names(uso) == 'Datos inter. (UF)'] <- 'Datos inter.'
          names(uso)[names(uso) == 'Datos nac. (UF)'] <- 'Datos nac.'
          names(uso)[names(uso) == 'Datos espec. (UF)'] <- 'Datos espec.'
          names(uso)[names(uso) == 'SMS (UF)'] <- 'SMS'
          names(uso)[names(uso) == 'SMS nac. (UF)'] <- 'SMS nac.'
          names(uso)[names(uso) == 'SMS inter. (UF)'] <- 'SMS inter.'
          names(uso)[names(uso) == 'SMS espec. (UF)'] <- 'SMS espec.'
          names(uso)[names(uso) == 'SMS hacia inter. (UF)'] <- 'SMS hacia inter.'
          names(uso)[names(uso) == 'SMS roaming (UF)'] <- 'SMS roaming'
          names(uso)[names(uso) == 'MMS (UF)'] <- 'MMS'
          names(uso)[names(uso) == 'MMS nac. (UF)'] <- 'MMS nac.'
          names(uso)[names(uso) == 'MMS inter. (UF)'] <- 'MMS inter.'
          names(uso)[names(uso) == 'MMS espec. (UF)'] <- 'MMS espec.'
          names(uso)[names(uso) == 'MMS hacia inter. (UF)'] <- 'MMS hacia inter.'
          names(uso)[names(uso) == 'MMS roaming (UF)'] <- 'MMS roaming'
          names(uso)[names(uso) == 'SMS/MMS (UF)'] <- 'SMS/MMS'
          names(uso)[names(uso) == 'SMS/MMS nac. (UF)'] <- 'SMS/MMS nac.'
          names(uso)[names(uso) == 'SMS/MMS inter. (UF)'] <- 'SMS/MMS inter.'
          names(uso)[names(uso) == 'SMS/MMS espec. (UF)'] <- 'SMS/MMS espec.'
          names(uso)[names(uso) == 'SMS/MMS hacia inter. (UF)'] <- 'SMS/MMS hacia inter.'
          names(uso)[names(uso) == 'SMS/MMS roaming (UF)'] <- 'SMS/MMS roaming'
          names(uso)[names(uso) == 'Business Hours (UF)'] <- 'Business Hours'
          names(uso)[names(uso) == 'Private (UF)'] <- 'Private'
          names(uso)[names(uso) == 'Horario laboral (UF)'] <- 'Horario laboral'
          names(uso)[names(uso) == 'Horario privado (UF)'] <- 'Horario privado'
          names(uso)[names(uso) == 'Exterior (UF)'] <- 'Exterior'
          names(uso)[names(uso) == 'Interior (UF)'] <- 'Interior'
          names(uso)[names(uso) == 'Copias B/N (UF)'] <- 'Copias B/N'
          names(uso)[names(uso) == 'Copias Color (UF)'] <- 'Copias Color'
          names(uso)[names(uso) == 'Copias (UF)'] <- 'Copias'
          names(uso)[names(uso) == 'Minimo Copias (UF)'] <- 'Minimo Copias'
          names(uso)[names(uso) == 'Minimo Copias B/N (UF)'] <- 'Minimo Copias B/N'
          names(uso)[names(uso) == 'Minimo Copias Color (UF)'] <- 'Minimo Copias Color'
          names(uso)[names(uso) == 'IVA (UF)'] <- 'IVA'
          
          
        }
        
        uso<<-uso
        uso[["Acceso"]]<-as.character(uso[["Acceso"]])
        #Se seleccionan las columnas que requiere el informe
        if (client=="igm"){
          CF<-subset(uso,is.na(uso[["Acceso"]]))
          uso1<-subset(uso,!is.na(uso[["Acceso"]]))
          if(length(CF[["Acceso"]])>0){
            CF[["Acceso"]]<-CF[["Nombre"]]
            CF[["Acceso fix"]]<-CF[["Nombre"]]
            CF[["Centro de facturación"]]<-CF[["Nombre"]]
            CF[["Tipo"]]<-"Centro de facturación"
            uso<-rbind(CF,uso1)
          }
          uso[["Nombre"]]<-NULL
          rm(CF,uso1)
          
          uso<<-uso
        uso<-subset(uso,select =c("Acceso",
                    "Proveedor",
                    "Período de",
                    "Usuario",
                    "Equipo",
                    "Tipo",
                    "Centro de facturación",
                    "Total",
                    "Plano tarifario",
                    "Uso",
                    "Servicios",
                    "Descuentos",
                    "Descuento > Plano tarifario",
                    "Voz",
                    "Voz nacional",
                    "Voz hacia Int",
                    "Voz roaming",
                    "Datos",
                    "Datos nac.",
                    "Datos inter.",
                    "SMS/MMS",
                    "SMS/MMS nac.",
                    "SMS/MMS inter.",
                    "SMS/MMS espec.",
                    "Voz (sec)",
                    "Voz nac. (sec)",
                    "Voz hacia inter. (sec)",
                    "Voz roaming (sec)",
                    "Datos (KB)",
                    "Datos nac. (KB)",
                    "Datos inter. (KB)",
                    "N.° SMS/MMS",
                    "Fecha",
                    "Acceso fix"))
        }
        else if (client == "afm"){
          uso<-subset(uso,select = c("Acceso",
                                     "Proveedor",
                                     "Período de",
                                     "Tipo",
                                     "Centro de facturación",
                                     "Total",
                                     "Plano tarifario",
                                     "Descuentos",
                                     "Servicios",
                                     "Servicios > Opciones",
                                     "Servicios > Otros",
                                     "Descuento > Otros",
                                     "Descuento > Opciones",
                                     "Voz",
                                     "Voz nacional",
                                     "Voz roaming",
                                     "Datos",
                                     "Datos nac.",
                                     "Datos inter.",
                                     "Uso (rebajado)",
                                     "N.° Voz nacional",
                                     "Datos inter. (KB)",
                                     "Voz nac. (sec)",
                                     "Voz roaming (sec)",
                                     "Fecha",
                                     "Acceso fix"
                                     ))
        }
        else if (client == "ige"){
          uso<-subset(uso,select = c("Acceso",
                                     "Proveedor",
                                     "Período de",
                                     "Tipo",
                                     "Centro de facturación",
                                     "Equipo",
                                     "Total",
                                     "Plano tarifario",
                                     "Descuentos",
                                     "Fecha",
                                     "Acceso fix"
                                     ))
        }
        else if (client == "igi"){
          uso<-subset(uso,select = c("Acceso",
                                     "Período de",
                                     "Tipo",
                                     "Usuario ID",
                                     "Centro de facturación",
                                     "Proveedor",
                                     "Copias",
                                     "Copias B/N",
                                     "Copias Color",
                                     "N.° Copias B/N",
                                     "N.° Copias Color",
                                     "N.° Copias",
                                     "Total",
                                     "Contador Color Actual",
                                     "Contador B/N Actual",
                                     "Descuentos",
                                     "Plano tarifario",
                                     "Fecha",                
                                     "Acceso fix"
                                      ))
        }
        else if (client == "lmovil"){
          uso<-subset(uso,!is.na(uso[["Acceso"]]))
        }#FALTA RE HACER LMOVIL COMPLETO DE UNA MANERA OPTIMIZADA
  
##############Carga a localhost############
      uso<-subset(uso,TRUE)
uso<<-uso
        dbGetQuery(DB, "SET NAMES 'latin1';")
      dbWriteTable(
        DB,
        "usos",
        uso,
        field.types = NULL,
        row.names = FALSE,
        overwrite = TRUE,
        append = FALSE,
        allow.keywords = FALSE
      )
      
      uso <<- uso
     
    }
    #Run the following code if theres a file in the export file input
    if (!is.null(export)) {
      #if xlsx file then Set names latin1 else if csv2 then Set names utf8
      dbGetQuery(DB, "SET NAMES 'latin1';")
      #Copy the file uploaded and add the .xlsx file type to the temp file
      file.copy(export$datapath,
                paste(export$datapath, ".xlsx", sep = ""))
      
      #######################################USERS############
      
      if (client =="afm" | client == "test") {
        #Read the xlsx file at the sheet USERS
        USERS <- read.xlsx(export$datapath,
                           sheet = "USERS",
                           startRow = 1)
        #Only select the following columns of the file
        USERS <- USERS[c(1, 5, 9)]
        #Rename the columns to the following
        names(USERS) <-
          c("UUI",
            "Nombre",
            "Estado")
        #In case the following file exist, delete it
        file.remove("USERS.txt")
        #Create an txt file from the previus columns, this is needed for getting the right encoding UTF8
        write.table(USERS, file = "USERS.txt", fileEncoding = "UTF-8")
        #Read the txt File
        USERS <- read.table(file = "USERS.txt", encoding = "UTF-8")
        #Upload to the DB the data
        dbWriteTable(
          DB,
          "users",
          USERS,
          field.types = NULL ,
          row.names = FALSE,
          overwrite = TRUE,
          append = FALSE,
          allow.keywords = FALSE
        )
        USERS <<- USERS
      }
      #######################################ACCESSES############
      
      ACCESSES <<- read.xlsx(export$datapath,
                            sheet = "ACCESSES",
                            startRow = 1)
 
      file.remove("ACCESSES.txt")
   
      if (client == "igm"|client =="igi"|client=="ige"| client == "test"){
        #Only select the columns with the following titles
        if(client == "ige"){
          if (customfield == "Si"){
            
            write.table(ACCESSES, file = "ACCESSES.txt", fileEncoding = "UTF-8")
            ACCESSES2 <-
              read.table(file = "ACCESSES.txt", encoding = "UTF-8")
            
            CUSTOM <- subset(ACCESSES2,select = c("ACCESS.NUMBER",
                                                  "CUSTOM_FIELD.BW",
                                                  "CUSTOM_FIELD.Descripción",
                                                  "CUSTOM_FIELD.Destino",
                                                  "CUSTOM_FIELD.Prioridad",
                                                  "CUSTOM_FIELD.Red.de.acceso",
                                                  "CUSTOM_FIELD.Geo.Referencia",
                                                  "CUSTOM_FIELD.Tipología",
                                                  "CUSTOM_FIELD.Edificación",
                                                  "CUSTOM_FIELD.CAPEX.OPEX"))
            
            names(CUSTOM)<-c("Acceso",
                             "BW",
                             "Descripción",
                             "Destino",
                             "Prioridad",
                             "Red de acceso",
                             "Geo Referencia",
                             "Tipología",
                             "Edificación",
                             "CAPEX/OPEX")
            
            CUSTOM<<-CUSTOM
          }
          
        }
        if(client == "igm"){
          if (customfield == "Si"){
            
            write.table(ACCESSES, file = "ACCESSES.txt", fileEncoding = "UTF-8")
            ACCESSES2 <-
              read.table(file = "ACCESSES.txt", encoding = "UTF-8")
            
            CUSTOM <- subset(ACCESSES2,select = c("ACCESS.NUMBER",
                                                  "CUSTOM_FIELD.Tipología"))
            
            names(CUSTOM)<-c("Acceso",
                             "Tipología")
            
            CUSTOM<<-CUSTOM
          }
          }
        a<-as.Date(ACCESSES$ELIGIBILITY.DATE, origin="1899-12-30")
        ACCESSES[,'Fecha Renovación']<-as.Date(ACCESSES$ELIGIBILITY.DATE, origin="1899-12-30")
        ACCESSES <-
          subset(
            ACCESSES,
            select = c(
              "ACCESS.NUMBER",
              "TYPE",
              "STATUS",
              "CARRIER_ORG:1",
              "CARRIER_ORG:2",
              "CARRIER_ORG:3",
              if(is.null(ACCESSES[["MANAGEMENT_ORG:1"]])==FALSE)
              {"MANAGEMENT_ORG:1"},
              if(is.null(ACCESSES[["MANAGEMENT_ORG:2"]])==FALSE)
              {"MANAGEMENT_ORG:2"},
              if(is.null(ACCESSES[["MANAGEMENT_ORG:3"]])==FALSE)
              {"MANAGEMENT_ORG:3"},
              if(is.null(ACCESSES[["MANAGEMENT_ORG:4"]])==FALSE)
              {"MANAGEMENT_ORG:4"},
              if(is.null(ACCESSES[["MANAGEMENT_ORG:5"]])==FALSE)
              {"MANAGEMENT_ORG:5"},
              if(is.null(ACCESSES[["MANAGEMENT_ORG:6"]])==FALSE)
              {"MANAGEMENT_ORG:6"},
              if(is.null(ACCESSES[["MANAGEMENT_ORG:7"]])==FALSE)
              {"MANAGEMENT_ORG:7"},
              if(is.null(ACCESSES[["MANAGEMENT_ORG:8"]])==FALSE)
              {"MANAGEMENT_ORG:8"},
              "Fecha Renovación"
            )
          )
        write.table(ACCESSES, file = "ACCESSES.txt", fileEncoding = "UTF-8")
        ACCESSES <-
          read.table(file = "ACCESSES.txt", encoding = "UTF-8")
        names(ACCESSES) <-
          c("Acceso",
            "Tipo",
            "Estado",
            "Proveedor",
            "Proveedor Nivel 2",
            "Proveedor Nivel 3",
            if(is.null(ACCESSES$MANAGEMENT_ORG.1)==FALSE)
            {"MANAGEMENTORG1"},
            if(is.null(ACCESSES$MANAGEMENT_ORG.2)==FALSE)
            {"MANAGEMENTORG2"},
            if(is.null(ACCESSES$MANAGEMENT_ORG.3)==FALSE)
            {"MANAGEMENTORG3"},
            if(is.null(ACCESSES$MANAGEMENT_ORG.4)==FALSE)
            {"MANAGEMENTORG4"},
            if(is.null(ACCESSES$MANAGEMENT_ORG.5)==FALSE)
            {"MANAGEMENTORG5"},
            if(is.null(ACCESSES$MANAGEMENT_ORG.6)==FALSE)
            {"MANAGEMENTORG6"},
            if(is.null(ACCESSES$MANAGEMENT_ORG.7)==FALSE)
            {"MANAGEMENTORG7"},
            if(is.null(ACCESSES$MANAGEMENT_ORG.8)==FALSE)
            {"MANAGEMENTORG8"},
            "Fecha Renovación"
            
            )
        #Create a column with the access number without the country code
        ACCESSES[, 'Acceso fix'] <-
          lapply(ACCESSES['Acceso'], function(x)
            substring(x, 3))
        ACCESSES<<-ACCESSES
        if(is.null(ACCESSES[["MANAGEMENTORG8"]])==FALSE&
           is.null(ACCESSES[["MANAGEMENTORG7"]])==FALSE&
           is.null(ACCESSES[["MANAGEMENTORG6"]])==FALSE&
           is.null(ACCESSES[["MANAGEMENTORG5"]])==FALSE&
           is.null(ACCESSES[["MANAGEMENTORG4"]])==FALSE&
           is.null(ACCESSES[["MANAGEMENTORG3"]])==FALSE&
           is.null(ACCESSES[["MANAGEMENTORG2"]])==FALSE&
           is.null(ACCESSES[["MANAGEMENTORG1"]])==FALSE)
        {
          MNG<-subset(ACCESSES,select = c("MANAGEMENTORG1",
                                          "MANAGEMENTORG2",
                                          "MANAGEMENTORG3",
                                          "MANAGEMENTORG4",
                                          "MANAGEMENTORG5",
                                          "MANAGEMENTORG6",
                                          "MANAGEMENTORG7",
                                          "MANAGEMENTORG8"))
          a<-duplicated(MNG[["MANAGEMENTORG8"]])
          MNG[["duplicados"]]<-a
          MNG<-subset(MNG,MNG[["duplicados"]]=="FALSE")
          MNG[["duplicados"]]<-NULL
          MNG<<-MNG
          a<-data.frame("Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco")
          MNG<-data.frame(rbind(as.matrix(MNG),as.matrix(a)))
          MNG<<-MNG
          dbWriteTable(
            DB,
            "accesses",
            ACCESSES,
            field.types = list(
              Acceso = "varchar(255)",
              Proveedor = "varchar(255)",
              Estado = "varchar(255)",
              `Fecha Renovación` = "date",
              Tipo = "varchar(255)",
              `Proveedor Nivel 2` = "varchar(255)",
              `Proveedor Nivel 3` = "varchar(255)",
              `MANAGEMENTORG1` = "varchar(255)",
              `MANAGEMENTORG2` = "varchar(255)",
              `MANAGEMENTORG3` = "varchar(255)",
              `MANAGEMENTORG4` = "varchar(255)",
              `MANAGEMENTORG5` = "varchar(255)",
              `MANAGEMENTORG6` = "varchar(255)",
              `MANAGEMENTORG7` = "varchar(255)",
              `MANAGEMENTORG8` = "varchar(255)",
              `Acceso fix` = "varchar(255)"
            ) ,
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
        }
        else if(is.null(ACCESSES[["MANAGEMENTORG7"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG6"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG5"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG4"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG3"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG2"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG1"]])==FALSE)
        {
          MNG<-subset(ACCESSES,select = c("MANAGEMENTORG1",
                                          "MANAGEMENTORG2",
                                          "MANAGEMENTORG3",
                                          "MANAGEMENTORG4",
                                          "MANAGEMENTORG5",
                                          "MANAGEMENTORG6",
                                          "MANAGEMENTORG7"))
          a<-duplicated(MNG[["MANAGEMENTORG7"]])
          MNG[["duplicados"]]<-a
          MNG<-subset(MNG,MNG[["duplicados"]]=="FALSE")
          MNG[["duplicados"]]<-NULL
          MNG<<-MNG
          a<-data.frame("Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco")
          MNG<-data.frame(rbind(as.matrix(MNG),as.matrix(a)))
          MNG<<-MNG
          dbWriteTable(
            DB,
            "accesses",
            ACCESSES,
            field.types = list(
              Acceso = "varchar(255)",
              Proveedor = "varchar(255)",
              Estado = "varchar(255)",
              `Fecha Renovación` = "date",
              Tipo = "varchar(255)",
              `Proveedor Nivel 2` = "varchar(255)",
              `Proveedor Nivel 3` = "varchar(255)",
              `MANAGEMENTORG1` = "varchar(255)",
              `MANAGEMENTORG2` = "varchar(255)",
              `MANAGEMENTORG3` = "varchar(255)",
              `MANAGEMENTORG4` = "varchar(255)",
              `MANAGEMENTORG5` = "varchar(255)",
              `MANAGEMENTORG6` = "varchar(255)",
              `MANAGEMENTORG7` = "varchar(255)",
              `Acceso fix` = "varchar(255)"
            ) ,
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
        }
        else if(is.null(ACCESSES[["MANAGEMENTORG6"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG5"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG4"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG3"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG2"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG1"]])==FALSE)
        {
          MNG<<-subset(ACCESSES,select = c("MANAGEMENTORG1",
                                          "MANAGEMENTORG2",
                                          "MANAGEMENTORG3",
                                          "MANAGEMENTORG4",
                                          "MANAGEMENTORG5",
                                          "MANAGEMENTORG6"))
          a<-duplicated(MNG[["MANAGEMENTORG6"]])
          MNG[["duplicados"]]<-a
          MNG<-subset(MNG,MNG[["duplicados"]]=="FALSE")
          MNG[["duplicados"]]<-NULL
          MNG<<-MNG
          a<-data.frame("Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco")
          MNG<-data.frame(rbind(as.matrix(MNG),as.matrix(a)))
          MNG<<-MNG
          dbWriteTable(
            DB,
            "accesses",
            ACCESSES,
            field.types = list(
              Acceso = "varchar(255)",
              Proveedor = "varchar(255)",
              Estado = "varchar(255)",
              `Fecha Renovación` = "date",
              Tipo = "varchar(255)",
              `Proveedor Nivel 2` = "varchar(255)",
              `Proveedor Nivel 3` = "varchar(255)",
              `MANAGEMENTORG1` = "varchar(255)",
              `MANAGEMENTORG2` = "varchar(255)",
              `MANAGEMENTORG3` = "varchar(255)",
              `MANAGEMENTORG4` = "varchar(255)",
              `MANAGEMENTORG5` = "varchar(255)",
              `MANAGEMENTORG6` = "varchar(255)",
              `Acceso fix` = "varchar(255)"
            ) ,
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
        }
        else if(is.null(ACCESSES[["MANAGEMENTORG5"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG4"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG3"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG2"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG1"]])==FALSE)
        {
          MNG<-subset(ACCESSES,select = c("MANAGEMENTORG1",
                                          "MANAGEMENTORG2",
                                          "MANAGEMENTORG3",
                                          "MANAGEMENTORG4",
                                          "MANAGEMENTORG5"))
          a<-duplicated(MNG[["MANAGEMENTORG5"]])
          MNG[["duplicados"]]<-a
          MNG<-subset(MNG,MNG[["duplicados"]]=="FALSE")
          MNG[["duplicados"]]<-NULL
          MNG<<-MNG
          a<-data.frame("Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco")
          MNG<-data.frame(rbind(as.matrix(MNG),as.matrix(a)))
          MNG<<-MNG
          dbWriteTable(
            DB,
            "accesses",
            ACCESSES,
            field.types = list(
              Acceso = "varchar(255)",
              Proveedor = "varchar(255)",
              Estado = "varchar(255)",
              `Fecha Renovación` = "date",
              Tipo = "varchar(255)",
              `Proveedor Nivel 2` = "varchar(255)",
              `Proveedor Nivel 3` = "varchar(255)",
              `MANAGEMENTORG1` = "varchar(255)",
              `MANAGEMENTORG2` = "varchar(255)",
              `MANAGEMENTORG3` = "varchar(255)",
              `MANAGEMENTORG4` = "varchar(255)",
              `MANAGEMENTORG5` = "varchar(255)",
              `Acceso fix` = "varchar(255)"
            ) ,
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
        }
        else if(is.null(ACCESSES[["MANAGEMENTORG4"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG3"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG2"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG1"]])==FALSE)
        {
          MNG<-subset(ACCESSES,select = c("MANAGEMENTORG1",
                                          "MANAGEMENTORG2",
                                          "MANAGEMENTORG3",
                                          "MANAGEMENTORG4"))
          a<-duplicated(MNG[["MANAGEMENTORG4"]])
          MNG[["duplicados"]]<-a
          MNG<-subset(MNG,MNG[["duplicados"]]=="FALSE")
          MNG[["duplicados"]]<-NULL
          MNG<<-MNG
          a<-data.frame("Sin Ceco","Sin Ceco","Sin Ceco","Sin Ceco")
          MNG<-data.frame(rbind(as.matrix(MNG),as.matrix(a)))
          MNG<<-MNG
          dbWriteTable(
            DB,
            "accesses",
            ACCESSES,
            field.types = list(
              Acceso = "varchar(255)",
              Proveedor = "varchar(255)",
              Estado = "varchar(255)",
              `Fecha Renovación` = "date",
              Tipo = "varchar(255)",
              `Proveedor Nivel 2` = "varchar(255)",
              `Proveedor Nivel 3` = "varchar(255)",
              `MANAGEMENTORG1` = "varchar(255)",
              `MANAGEMENTORG2` = "varchar(255)",
              `MANAGEMENTORG3` = "varchar(255)",
              `MANAGEMENTORG4` = "varchar(255)",
              `Acceso fix` = "varchar(255)"
            ) ,
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
        }
        else if(is.null(ACCESSES[["MANAGEMENTORG3"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG2"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG1"]])==FALSE)
        {
          MNG<-subset(ACCESSES,select = c("MANAGEMENTORG1",
                                          "MANAGEMENTORG2",
                                          "MANAGEMENTORG3"))
          a<-duplicated(MNG[["MANAGEMENTORG3"]])
          MNG[["duplicados"]]<-a
          MNG<-subset(MNG,MNG[["duplicados"]]=="FALSE")
          MNG[["duplicados"]]<-NULL
          MNG<<-MNG
          a<-data.frame("Sin Ceco","Sin Ceco","Sin Ceco")
          MNG<-data.frame(rbind(as.matrix(MNG),as.matrix(a)))
          MNG<<-MNG
          dbWriteTable(
            DB,
            "accesses",
            ACCESSES,
            field.types = list(
              Acceso = "varchar(255)",
              Proveedor = "varchar(255)",
              Estado = "varchar(255)",
              `Fecha Renovación` = "date",
              Tipo = "varchar(255)",
              `Proveedor Nivel 2` = "varchar(255)",
              `Proveedor Nivel 3` = "varchar(255)",
              `MANAGEMENTORG1` = "varchar(255)",
              `MANAGEMENTORG2` = "varchar(255)",
              `MANAGEMENTORG3` = "varchar(255)",
              `Acceso fix` = "varchar(255)"
            ) ,
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
        }
        else if(is.null(ACCESSES[["MANAGEMENTORG2"]])==FALSE&
                is.null(ACCESSES[["MANAGEMENTORG1"]])==FALSE)
        {
          MNG<-subset(ACCESSES,select = c("MANAGEMENTORG1",
                                          "MANAGEMENTORG2"))
          a<-duplicated(MNG[["MANAGEMENTORG2"]])
          MNG[["duplicados"]]<-a
          MNG<-subset(MNG,MNG[["duplicados"]]=="FALSE")
          MNG[["duplicados"]]<-NULL
          MNG<<-MNG
          a<-data.frame("Sin Ceco","Sin Ceco")
          MNG<-data.frame(rbind(as.matrix(MNG),as.matrix(a)))
          MNG<<-MNG
          dbWriteTable(
            DB,
            "accesses",
            ACCESSES,
            field.types = list(
              Acceso = "varchar(255)",
              Proveedor = "varchar(255)",
              Estado = "varchar(255)",
              `Fecha Renovación` = "date",
              Tipo = "varchar(255)",
              `Proveedor Nivel 2` = "varchar(255)",
              `Proveedor Nivel 3` = "varchar(255)",
              `MANAGEMENTORG1` = "varchar(255)",
              `MANAGEMENTORG2` = "varchar(255)",
              `Acceso fix` = "varchar(255)"
            ) ,
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
        }
        else if(is.null(ACCESSES[["MANAGEMENTORG1"]])==FALSE)
        {
          MNG<-subset(ACCESSES,select = c("MANAGEMENTORG1"))
          a<-duplicated(MNG[["MANAGEMENTORG1"]])
          MNG[["duplicados"]]<-a
          MNG<-subset(MNG,MNG[["duplicados"]]=="FALSE")
          MNG[["duplicados"]]<-NULL
          MNG<<-MNG
          a<-data.frame("Sin Ceco")
          MNG<-data.frame(rbind(as.matrix(MNG),as.matrix(a)))
          MNG<<-MNG
          dbWriteTable(
            DB,
            "accesses",
            ACCESSES,
            field.types = list(
              Acceso = "varchar(255)",
              Proveedor = "varchar(255)",
              Estado = "varchar(255)",
              `Fecha Renovación` = "date",
              Tipo = "varchar(255)",
              `Proveedor Nivel 2` = "varchar(255)",
              `Proveedor Nivel 3` = "varchar(255)",
              `MANAGEMENTORG1` = "varchar(255)",
              `Acceso fix` = "varchar(255)"
            ) ,
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
        }
        else 
        {
          dbWriteTable(
            DB,
            "accesses",
            ACCESSES,
            field.types = list(
              Acceso = "varchar(255)",
              Proveedor = "varchar(255)",
              Estado = "varchar(255)",
              `Fecha Renovación` = "date",
              Tipo = "varchar(255)",
              `Proveedor Nivel 2` = "varchar(255)",
              `Proveedor Nivel 3` = "varchar(255)",
              `Acceso fix` = "varchar(255)"
            ) ,
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
        }
        
        ACCESSES <<- ACCESSES
        
      }
      else if (client == "lmovil") {
        #Only select the columns with the following titles
        
        ACCESSES <-
          subset(
            ACCESSES,
            select = c(
              "ACCESS.NUMBER",
              "CARRIER_ORG:1",
              "CARRIER_ORG:2",
              "CARRIER_ORG:3"
            )
          )
        write.table(ACCESSES, file = "ACCESSES.txt", fileEncoding = "UTF-8")
        ACCESSES <-
          read.table(file = "ACCESSES.txt", encoding = "UTF-8")
        names(ACCESSES) <-
          c("Acceso",
            "Proveedor",
            "Proveedor Nivel 2",
            "Proveedor Nivel 3")
        #Create a column with the access number without the country code
        ACCESSES[, 'Acceso fix'] <-
          lapply(ACCESSES['Acceso'], function(x)
            substring(x, 3))
        
        dbWriteTable(
          DB,
          "accesses",
          ACCESSES,
          field.types = list(
            Acceso = "varchar(255)",
            Proveedor = "varchar(255)",
            `Proveedor Nivel 2` = "varchar(255)",
            `Proveedor Nivel 3` = "varchar(255)",
            `Acceso fix` = "varchar(255)"
          ) ,
          row.names = FALSE,
          overwrite = TRUE,
          append = FALSE,
          allow.keywords = FALSE
        )
        ACCESSES <<- ACCESSES
        
      }
      file.remove("ACCESSES.txt")
      #######################################DEVICES############
      
        DEVICES <- read.xlsx(export$datapath,
                             sheet = "DEVICES",
                             startRow = 1)
        file.remove("DEVICES.txt")
        
        if (client == "test") {
          DEVICES <- DEVICES[c(1, 2, 3, 4, 9)]
          write.table(DEVICES, file = "DEVICES.txt", fileEncoding = "UTF-8")
          DEVICES <-
            read.table(file = "DEVICES.txt", encoding = "UTF-8")
          names(DEVICES) <-
            c("Tipo",
              "Modelo",
              "REFNUM",
              "IMEI",
              "Estado")
        }
        else{
          DEVICES <-NULL
        }
   if (!is.null(DEVICES)){
        dbWriteTable(
          DB,
          "devices",
          DEVICES,
          field.types = NULL,
          row.names = FALSE,
          overwrite = TRUE,
          append = FALSE,
          allow.keywords = FALSE
        )
        DEVICES<<-DEVICES
   }
      #######################################ASSOCIATIONS############
      if (client != "lmovil"& client != "igm" & client != "igi" & client != "ige") {
        ASSOCIATIONS <- read.xlsx(export$datapath,
                                  sheet = "ASSOCIATIONS",
                                  startRow = 1)
        ASSOCIATIONS <- ASSOCIATIONS[c(1:4)]
        write.table(ASSOCIATIONS,
                    file = "ASSOCIATIONS.txt",
                    fileEncoding = "UTF-8")
        ASSOCIATIONS <-
          read.table(file = "ASSOCIATIONS.txt", encoding = "UTF-8")
        names(ASSOCIATIONS) <- c("Acceso", "UUI", "IMEI", "REFNUM")
        dbWriteTable(
          DB,
          "associations",
          ASSOCIATIONS,
          field.types = NULL ,
          row.names = FALSE,
          overwrite = TRUE,
          append = FALSE,
          allow.keywords = FALSE
        )
        file.remove("ASSOCIATIONS.txt")
        ASSOCIATIONS <<- ASSOCIATIONS
      }
      #######################################PRODUCT_ASSOCIATIONS############
      if (client != "lmovil"& client != "igm" & client != "igi" & client != "ige") {
        PRODUCT_ASSOCIATIONS <- read.xlsx(export$datapath,
                                          sheet = "PRODUCT ASSOCIATIONS",
                                          startRow = 1)
        PRODUCT_ASSOCIATIONS <- PRODUCT_ASSOCIATIONS[c(1:2)]
        file.remove("PRODUCT_ASSOCIATIONS.txt")
        write.table(PRODUCT_ASSOCIATIONS,
                    file = "PRODUCT_ASSOCIATIONS.txt",
                    fileEncoding = "UTF-8")
        PRODUCT_ASSOCIATIONS <-
          read.table(file = "PRODUCT_ASSOCIATIONS.txt", encoding = "UTF-8")
        names(PRODUCT_ASSOCIATIONS) <- c("Acceso", "Plan")
        dbWriteTable(
          DB,
          "product_associations",
          PRODUCT_ASSOCIATIONS,
          field.types = NULL ,
          row.names = FALSE,
          overwrite = TRUE,
          append = FALSE,
          allow.keywords = FALSE
        )
      }
  
    }
    #Run the following code if theres a file in the presupuesto file input
    if (!is.null(presupuesto)){
      dbGetQuery(DB, "SET NAMES 'latin1';")
      dataFilesUF2 <- NULL
      dataFilesUF2 <- lapply(input$presupuesto[['datapath']], function(x) read.csv2(x,encoding = "UTF-8",check.names = FALSE,header = TRUE))
      
      #Append all usos files to the same dataframe
      
      presupuestos <- rbindlist(dataFilesUF2,fill = TRUE)
      
      
      a<-as.character(names(presupuestos))
      d<-matrix(nrow = length(a),ncol = 1)
      for(i in 1:length(a)){
        c<-substr((a[i]),3,10)
        if(c=="cceso"){
          print("Encontrado")
          d[i,1]<-"Acceso"
        }
        else if (c=="ombre"){
          print("Encontrado 2")
          d[i,1]<-"Nombre"
        }
        else{
          d[i,1]<-a[i]
        }
      }
      b<-as.character(d)
      colnames(presupuestos)<-b
      
      rm(a,b,c,d)
      presupuestos<<-presupuestos
      if (client=="igm"){
        CF<-subset(presupuestos,is.na(presupuestos[["Acceso"]]))
        presupuestos1<-subset(presupuestos,!is.na(presupuestos[["Acceso"]]))
        if(length(CF[["Acceso"]])>0){
          CF[["Acceso"]]<-CF[["Nombre"]]
          CF[["Centro de facturación"]]<-CF[["Nombre"]]
          CF[["Tipo"]]<-"Centro de facturación"
          presupuestos<-rbind(CF,presupuestos1)
        }
        presupuestos[["Nombre"]]<-NULL
        rm(CF,presupuestos1)
        
        presupuestos<<-presupuestos
      }
      if(divisa == "CLP"){
        names(presupuestos)[names(presupuestos) == 'Total (CLP)'] <- 'Total'
        names(presupuestos)[names(presupuestos) == 'Plano tarifario (CLP)'] <- 'Plano tarifario'
        names(presupuestos)[names(presupuestos) == 'Descuentos (CLP)'] <- 'Descuentos'

      }
      else if (divisa == "UF"){
        names(presupuestos)[names(presupuestos) == 'Total (UF)'] <- 'Total'
        names(presupuestos)[names(presupuestos) == 'Plano tarifario (UF)'] <- 'Plano tarifario'
        names(presupuestos)[names(presupuestos) == 'Descuentos (UF)'] <- 'Descuentos'
        
      }
      presupuestos<<-presupuestos
      
      
      for (k in 1:12) {
        presupuestos[, 'Período de'] <-
          lapply(presupuestos[, 'Período de'], function(x)
            gsub(names(y[k]), y[[k]], x))
      }
      {
        presupuestos[, 'Fecha'] <-
          lapply(presupuestos[, 'Período de'], function(x)
            paste(substr(x , 3 , 6),
                  substr(x , 1 , 2),
                  "01",
                  sep = "/"))
        #Delete the Column MES that its not needed now
        presupuestos[, 'Período de'] <- NULL
        presupuestos<-subset(presupuestos,select = c("Acceso",
                                                     "Tipo",
                                                     "Usuario ID",
                                                     "Centro de facturación",
                                                     "Proveedor",
                                                     "Total",
                                                     "Descuentos",
                                                     "Plano tarifario",
                                                     "Fecha"))
        presupuestos<<-presupuestos
      
        
      }
      
      dbWriteTable(
        DB,
        "presupuesto",
        presupuestos,
        field.types = NULL,
        row.names = FALSE,
        overwrite = TRUE,
        append = FALSE,
        allow.keywords = FALSE
      )
    }
    #Run the following code if theres a file in the planes file input
    if (!is.null(plantilla)) {
      #if xlsx file then Set names latin1 else if csv2 then Set names utf8
      dbGetQuery(DB, "SET NAMES 'utf8';")
      
        file.copy(plantilla$datapath,
                  paste(plantilla$datapath, ".xlsx", sep = ""))
        
        PLANTILLA <<- read.xlsx(plantilla$datapath,
                          sheet = "Uso por accesoproducto",
                          startRow = 1,
                          check.names = FALSE)
       
        PLANTILLA<-fixnames(PLANTILLA)
        gnames<-names(PLANTILLA)
        write.table(PLANTILLA,
                    file = "Plantilla.txt",
                    fileEncoding = "UTF-8")
        PLANTILLA <- read.table(file = "Plantilla.txt", encoding = "UTF-8",check.names = FALSE)
        
    
        names(PLANTILLA)<-gnames
        if(divisa=="CLP"){
        names(PLANTILLA)[names(PLANTILLA) == 'Importe de las opciones facturadas (CLP)'] <- 'Importe de las opciones facturadas'
        names(PLANTILLA)[names(PLANTILLA) == 'Importe descuentos sobre plano tarifario (CLP)'] <- 'Importe descuentos sobre plano tarifario'
        names(PLANTILLA)[names(PLANTILLA) == 'Importe de las opciones descontadas (CLP)'] <- 'Importe de las opciones descontadas'
        }
        else if (divisa == "UF"){
        names(PLANTILLA)[names(PLANTILLA) == 'Importe de las opciones facturadas (UF)'] <- 'Importe de las opciones facturadas'
        names(PLANTILLA)[names(PLANTILLA) == 'Importe descuentos sobre plano tarifario (UF)'] <- 'Importe descuentos sobre plano tarifario'
        names(PLANTILLA)[names(PLANTILLA) == 'Importe de las opciones descontadas (UF)'] <- 'Importe de las opciones descontadas'
        }
        PLANTILLA<<-PLANTILLA
        
       
        dbWriteTable(
          DB,
          "plantilla",
          PLANTILLA,
          field.types = NULL,
          row.names = FALSE,
          overwrite = TRUE,
          append = FALSE,
          allow.keywords = FALSE
        )
        file.remove("Plantilla.txt")
    }
    #Run the following code if theres a file in the nombre and link text input
    if (!is.null(input$nombre)) {
      #Set variable names
      
      nombre <<- input$nombre
      if (nombre=="Aguas Andinas"){
      link<<-"https://cdn.pbrd.co/images/GWEts67.jpg"}
      else if (nombre =="Carabineros de Chile"){
        link<<-"https://cdn.pbrd.co/images/GWEu27P.png"
      }
      else if (nombre =="Claro"){
        link<<-"https://cdn.pbrd.co/images/GWEuhGn.jpg"
      }
      else if (nombre =="Copec"){
        link<<-"https://cdn.pbrd.co/images/GWEuxvg.png"
      }
      else if (nombre =="Enap"){
        link<<-"https://cdn.pbrd.co/images/GWEuMor.gif"
      }
      else if (nombre =="Hogar de Cristo"){
        link<<-"https://cdn.pbrd.co/images/H1pXkPQO.png"
      }
      else if (nombre =="Nuevo Pudahuel"){
        link<<-"https://cdn.pbrd.co/images/GWEvad0.png"
      }
      else if (nombre =="Parque Arauco"){
        link<<-"https://cdn.pbrd.co/images/GWEvklk.png"
      }
      else if (nombre =="Rhona"){
        link<<-"https://cdn.pbrd.co/images/GWEvuiJ.png"
      }
      else if (nombre =="SAAM"){
        link<<-"https://cdn.pbrd.co/images/GWEvFgU.png"
      }
      else if (nombre =="Subsole"){
        link<<-"https://cdn.pbrd.co/images/GWEvOqH.png"
      }
      else if (nombre =="Walmart"){
        link<<-"https://cdn.pbrd.co/images/GWEvY8X.jpg"
      }
      else{print("No se encontro un link Para el nombre seleccionado")}
      #Create table logo_cliente if doesnt exist
      dbSendQuery(
        DB,
        "CREATE TABLE IF NOT EXISTS `logo_cliente` (
        `id` int(11) NOT NULL,
        `Nombre Cliente` varchar(255) NOT NULL,
        `Link` text NOT NULL);"
  )
      
      #Delete all data in logo_cliente
      dbSendQuery(DB, "TRUNCATE `logo_cliente`;")
      
      #Insert data in logo_cliente (if exist update)
      dbSendQuery(
        DB,
        paste(
          "INSERT INTO `logo_cliente`(`id`, `Nombre Cliente`, `Link`) VALUES (1,",
          input$nombre,
          ",",
          link,
          ")
          on duplicate key update
          `Nombre Cliente` = values(`Nombre Cliente`), `Link` = values(`Link`);",
          sep = '\''
        )
      )
      
    }
    #Run the following code if theres a file in the contrato file input
    if (!is.null(contrato)){
      dbGetQuery(DB, "SET NAMES 'latin1';")
      file.copy(contrato$datapath,
                paste(contrato$datapath, ".xlsx", sep = ""))
      if (proveedor == "Movistar CL") {
      ########################################MOVISTAR_PLANES############
        
      MOVISTAR_PLANES <- read.xlsx(contrato$datapath,
                                        sheet = "Movistar Planes",
                                        startRow = 1,
                                    na.strings = TRUE)
      file.remove("MOVISTAR_PLANES.txt")
      write.table(MOVISTAR_PLANES,
                  file = "MOVISTAR_PLANES.txt",
                  fileEncoding = "UTF-8")
      MOVISTAR_PLANES <-
        read.table(file = "MOVISTAR_PLANES.txt", encoding = "UTF-8")
      names(MOVISTAR_PLANES) <- c("Producto","Descripción","Tipo","Precio (CLP)","Voz (min)","Datos (KB)","Precio/min (CLP)","PrecioSC/min (CLP)","Precio/SMS (CLP)","Precio/KB (CLP)")
 
      dbWriteTable(
        DB,
        "movistar_planes",
        MOVISTAR_PLANES,
        field.types = NULL ,
        row.names = FALSE,
        overwrite = TRUE,
        append = FALSE,
        allow.keywords = FALSE
      )
      
      file.remove("MOVISTAR_PLANES.txt")
      MOVISTAR_PLANES<<-MOVISTAR_PLANES
      
      ########################################MOVISTAR_OPCIONES############
      
      
      MOVISTAR_OPCIONES <<- read.xlsx(contrato$datapath,
                                   sheet = "Movistar Opciones",
                                   startRow = 1,
                                   na.strings = TRUE)
      file.remove("MOVISTAR_OPCIONES.txt")
      write.table(MOVISTAR_OPCIONES,
                  file = "MOVISTAR_OPCIONES.txt",
                  fileEncoding = "UTF8")
      MOVISTAR_OPCIONES <<-
        read.table(file = "MOVISTAR_OPCIONES.txt", encoding = "UTF-8")
      names(MOVISTAR_OPCIONES) <<- c("Opciones","Tipo","Valor Opción","Minutos","KB","Duración")
      
      dbWriteTable(
        DB,
        "movistar_opciones",
        MOVISTAR_OPCIONES,
        field.types = NULL ,
        row.names = FALSE,
        overwrite = TRUE,
        append = FALSE,
        allow.keywords = FALSE
      )
      file.remove("MOVISTAR_OPCIONES.txt")
      ########################################MOVISTAR_PAISES############
      MOVISTAR_PAISES <<- read.xlsx(contrato$datapath,
                                     sheet = "Movistar Paises",
                                     startRow = 1)
      file.remove("MOVISTAR_PAISES.txt")
      write.table(MOVISTAR_PAISES,
                  file = "MOVISTAR_PAISES.txt",
                  fileEncoding = "UTF-8")
      MOVISTAR_PAISES <-
        read.table(file = "MOVISTAR_PAISES.txt", encoding = "UTF-8")
      MOVISTAR_PAISES<-fixnames(MOVISTAR_PAISES)
      dbWriteTable(
        DB,
        "movistar_paises",
        MOVISTAR_PAISES,
        field.types = NULL ,
        row.names = FALSE,
        overwrite = TRUE,
        append = FALSE,
        allow.keywords = FALSE
      )
      file.remove("MOVISTAR_PAISES.txt")
      MOVISTAR_PAISES<<-MOVISTAR_PAISES
      ########################################MOVISTAR_ZONAS############
      MOVISTAR_ZONAS <<- read.xlsx(contrato$datapath,
                                    sheet = "Movistar Zonas",
                                    startRow = 1)
      file.remove("MOVISTAR_ZONAS.txt")
      write.table(MOVISTAR_ZONAS,
                  file = "MOVISTAR_ZONAS.txt",
                  fileEncoding = "UTF-8")
      MOVISTAR_ZONAS <-
        read.table(file = "MOVISTAR_ZONAS.txt", encoding = "UTF-8")
      MOVISTAR_ZONAS<-fixnames(MOVISTAR_ZONAS)
      dbWriteTable(
        DB,
        "movistar_zonas",
        MOVISTAR_ZONAS,
        field.types = NULL ,
        row.names = FALSE,
        overwrite = TRUE,
        append = FALSE,
        allow.keywords = FALSE
      )
      file.remove("MOVISTAR_ZONAS.txt")
      MOVISTAR_ZONAS<<-MOVISTAR_ZONAS
      }
    }
    #Run the following code if theres a file in the tipos file input
    {
      # if (!is.null(tipos)) {
      #   file.copy(tipos$datapath, paste(tipos$datapath, ".xlsx", sep = ""))
      #   
      #   TIPO <- read.xlsx(tipos$datapath,
      #                     sheet = 1,
      #                     startRow = 1)
      #   
      #   TIPO <- TIPO[c(1:4)]
      #   
      #   file.remove("TIPO.txt")
      #   write.table(TIPO,
      #               file = "TIPO.txt",
      #               fileEncoding = "UTF8")
      #   TIPO <- read.table(file = "TIPO.txt", encoding = "UTF8")
      #   
      #   names(TIPO) <- c("Producto",
      #                    "Descripcion Plan",
      #                    "Tipo",
      #                    "Proveedor")
      #   
      #   dbWriteTable(
      #     DB,
      #     "tipo_plan",
      #     TIPO,
      #     field.types = list(
      #       `Producto` = "varchar(255)",
      #       `Descripcion Plan` = "varchar(255)",
      #       `Tipo` = "varchar(255)",
      #       `Proveedor` = "varchar(255)"
      #     ),
      #     row.names = FALSE,
      #     overwrite = TRUE,
      #     append = FALSE,
      #     allow.keywords = FALSE
      #   )
      #   
      #   TIPO <<- TIPO
      #   
      #   file.remove("TIPO.txt")
      #   
      # }
    }# ACA SE ENCUENTRA TIPOS DE LMOVIL COMENTADO
    #Run the following code if theres a file in the cuentas file input
    {
    # if (!is.null(cuentas)) {
    #   #Copy the file uploaded and add the .xlsx file type to the temp file
    #   file.copy(cuentas$datapath,
    #             paste(cuentas$datapath, ".xlsx", sep = ""))
    #   
    #   CUENTAS <- read.xlsx(cuentas$datapath,
    #                        sheet = "CUENTAS",
    #                        startRow = 1)
    #   
    #   CUENTAS <- CUENTAS[c(1:4)]
    #   
    #   file.remove("CUENTAS.txt")
    #   write.table(CUENTAS,
    #               file = "CUENTAS.txt",
    #               fileEncoding = "UTF8")
    #   CUENTAS <- read.table(file = "CUENTAS.txt", encoding = "UTF8")
    #   
    #   names(CUENTAS) <- c("Empresa",
    #                       "RUT",
    #                       "Cuenta Cliente",
    #                       "Proveedor")
    #   
    #   dbWriteTable(
    #     DB,
    #     "cuentas",
    #     CUENTAS,
    #     field.types = list(
    #       `Empresa` = "varchar(255)",
    #       `RUT` = "varchar(255)",
    #       `Cuenta Cliente` = "varchar(255)",
    #       `Proveedor` = "varchar(255)"
    #     ),
    #     row.names = FALSE,
    #     overwrite = TRUE,
    #     append = FALSE,
    #     allow.keywords = FALSE
    #   )
    #   
    #   CUENTAS <<- CUENTAS
    #   
    #   file.remove("CUENTAS.txt")
    #   
    # }
    }#ACA SE ENCUENTRA CUENTAS DE LMOVIL COMENTADO
    #Run the following code if theres a file in the cdr file input
    if (!is.null(input$cdr)) {
      dbGetQuery(DB, "SET NAMES 'latin1';")
      CDRFile <<- NULL
      #Read CDR file with correct fileencoding
      CDRFile <<-
        lapply(input$cdr[['datapath']], function(x) read.csv2(x,encoding = "UTF-8"))
      cdr <<- rbindlist(CDRFile)
      cdr[, 'X.U.FEFF.Usuario'] <- NULL
      cdr<-fixnames(cdr)
      cdr[, 'Red recurrente'] <- NULL
      cdr[, 'Red destinada'] <- NULL
      cdr[, 'Organización de gestiòn'] <- NULL
      cdr[, 'VPN'] <- NULL
      cdr[, 'Llamadas internas'] <- NULL
      cdr[, 'Número de llamada fix'] <-
        lapply(cdr[, 'Número de llamada'], function(x)
          substring(x, 3))
      cdr<<-cdr
      
      dbWriteTable(
        DB,
        "cdr",
        cdr,
        field.types = list(
          `Número de llamada` = "char(11)",
          `Número de llamada fix` = "int(9)",
          `Número llamado` = "varchar(20)",
          `Tipo de llamada` = "ENUM('Datos','MMS','SMS','Voz','E-mail','Desconocidos') NOT NULL",
          `Fecha de llamada` = "DATE",
          `Geografía` = "ENUM('A internacional','Local','Regional','Nacional desconocido','Roaming entrante','Roaming saliente','Roaming desconocido','Internacional desconocido','Desconocidos') NOT NULL",
          `País emisor` = "varchar(40)",
          `País destinatario` = "varchar(40)",
          `Duración` = "SMALLINT(8) UNSIGNED NOT NULL",
          `Volumen` = "MEDIUMINT(10) UNSIGNED NOT NULL",
          `Precio` = "FLOAT(10,2) NOT NULL",
          `Tarificación` = "ENUM('En el plan','Más alla del plan','Desconocidos','Fuera de plan') NOT NULL",
          `Tecnología` = "varchar(40)",
          `Servicio llamado` = "varchar(255)",
          `Organización Proveedor` = "varchar(255)"
        ),
        row.names = FALSE,
        overwrite = TRUE,
        append = FALSE,
        allow.keywords = FALSE
      )
      
      
    }
    
    
    
    
    ##################################CAMBIOS SOBRE TABLAS###################
Consolidado<-NULL
    
    
    ##################################CAMBIOS A USOS#########################
    if(!is.null(input$usos)){
    month1 <- sapply(uso[,'Fecha'], substr, 6, 7)
    month <- as.numeric(month1)
    rm(month1)
    year1<-sapply(uso[,'Fecha'],substr,1,4)
    year<-as.numeric(year1)
    rm(year1)
    AAA<-data.frame(uso[,'Fecha'])
    AAA[,'month']<-month
    AAA[,'year']<-year
    BBB<-subset(AAA,
                AAA["year"]==max(year))
    rm(month,year)
    fin1<-max(BBB[["month"]])
    fin2<-max(BBB[["year"]])
    rm(BBB)
    for (i in 1:length(uso[["Acceso"]])){
      uso[["Meses"]][i]<-(fin2-AAA[["year"]][i])*12+fin1-AAA[["month"]][i]+1
    }
    uso2<-data.frame()
    for (i in 1:max(uso[["Meses"]])){
      uso1<-subset(uso,uso[["Meses"]]== i)
      
      uso1[,'Porcentaje']<-uso1[,'Total']/sum(uso1[["Total"]])
      
      
      uso2<-rbind(uso2,uso1)
    }
    uso<-uso2
    uso<<-uso
    }
    ########################CAMBIOS A PLANTILLA SERVICIOS FACTURADOS##########
    if ((client=="igm"&!is.null(plantilla))|(client=="afm"&!is.null(contrato))&!is.null(plantilla)){
    SFACTURADOS<-subset(PLANTILLA,select = c("Acceso",
                                             "Estado acceso",
                                             "Producto",
                                             "Tipo de producto",
                                             "Centro de facturación",
                                             "Importe de las opciones facturadas",
                                             "Importe descuentos sobre plano tarifario",
                                             "Importe de las opciones descontadas"))
    SFACTURADOS<<-SFACTURADOS
    #source("pj_igmSF.r", local = TRUE)
   SFACTURADOS[["Acceso"]]<-as.character(SFACTURADOS[["Acceso"]])
    
    
      SFOpciones<-subset(SFACTURADOS,
                         SFACTURADOS[["Tipo de producto"]]=="Option")
      SFOpciones<<-SFOpciones
      SFPlanes<-subset(SFACTURADOS,
                       SFACTURADOS[["Tipo de producto"]]=="Plano tarifario")
      
      SFPlanes<<-SFPlanes
      SFPlanes2<-SFPlanes
      ######################Ajuste de Servicios Facturados################
      
      ###############################################################################Parte 1 <-Ver duplicados de no duplicados
      a<-duplicated(SFPlanes2[["Acceso"]],fromLast = FALSE)
      b<-duplicated(SFPlanes2[["Acceso"]],fromLast = TRUE)
      SFPlanes2[["Duplicados"]]<-a
      SFPlanes2[["Duplicados2"]]<-b
      SFduplicados<-subset(SFPlanes2,SFPlanes2[["Duplicados"]]=="TRUE"|
                             SFPlanes2[["Duplicados2"]]=="TRUE")
      SF_no_duplicados<-subset(SFPlanes2,SFPlanes2[["Duplicados"]]=="FALSE"&
                                 SFPlanes2[["Duplicados2"]]=="FALSE")
      #############################################################################Parte 2 <- De los duplicados segmentar los que se sacan (sin cobro), de los que no
      if(length(SFduplicados[["Acceso"]])>0){
        if (client == "igm"){
        SF_a_evaluar<- SFduplicados
        SF_fueradecontratoSC<-subset(SF_a_evaluar,
                                     SF_a_evaluar[["Importe de las opciones descontadas"]]==0)
        SF_fueradecontratoCC<-subset(SF_a_evaluar,
                                     SF_a_evaluar[["Importe de las opciones descontadas"]]!=0)
        
        SF_CPduplicados<-SF_fueradecontratoCC
        }
        else if (client == "afm"){
          SF_a_evaluar<- merge(SFduplicados,
                               MOVISTAR_PLANES,
                               by = "Producto",
                               all.x = TRUE)
          SF_fueradecontratoSC<-subset(SF_a_evaluar,
                                       is.na(SF_a_evaluar[["Tipo"]])==TRUE &
                                         SF_a_evaluar[["Importe de las opciones descontadas"]]==0)
          SF_fueradecontratoCC<-subset(SF_a_evaluar,
                                       is.na(SF_a_evaluar[["Tipo"]])==TRUE &
                                         SF_a_evaluar[["Importe de las opciones descontadas"]]!=0)
          SF_en_contrato<-subset(SF_a_evaluar,
                                 is.na(SF_a_evaluar[["Tipo"]])==FALSE)
          SF_CPduplicados<-rbind(SF_en_contrato,SF_fueradecontratoCC)
        }
        #########################################################################Parte 3  <- Ver si los que tienen cobro aun estan duplicados tras haber sacado los sin cobro
        SF_CPduplicados[["Duplicado"]]<-NULL
        SF_CPduplicados[["Duplicado2"]]<-NULL
        a<-duplicated(SF_CPduplicados[["Acceso"]],fromLast = FALSE)
        b<-duplicated(SF_CPduplicados[["Acceso"]],fromLast = TRUE)
        SF_CPduplicados[["Duplicados"]]<-a
        SF_CPduplicados[["Duplicados2"]]<-b
        SFduplicados2<-subset(SF_CPduplicados,SF_CPduplicados[["Duplicados"]]=="TRUE"|SF_CPduplicados[["Duplicados2"]]=="TRUE")
        SFduplicadosbuenos<-subset(SF_CPduplicados,SF_CPduplicados[["Duplicados"]]=="FALSE"&SF_CPduplicados[["Duplicados2"]]=="FALSE")
        ###########################################################################Parte 4 <- Se suman los cobros de los que tienen multiples cobros creando tabla aparte
        if(length(SFduplicados2[["Acceso"]])>0){
          accesosunicos<-as.list(unique(SFduplicados2["Acceso"]))
          i<-1
          
          Acc<-c()
          EstAcc<-c()
          Prod<-c()
          TdProd<-c()
          CdF<-c()
          IOF<-c()
          IDPT<-c()
          IOD<-c()
          Accf<-c()
          for(i in 1:lengths(accesosunicos)){
            SFUnico<-subset(SFPlanes2,SFPlanes2[["Acceso"]] == as.character(accesosunicos[["Acceso"]][i]))
            
            Acc[i]<-SFUnico[["Acceso"]][1]
            EstAcc[i]<-"No determinado"
            Prod[i]<-"Multi producto"
            TdProd[i]<-"Plano tarifario"
            CdF[i]<-SFUnico[["Centro de facturación"]][1]
            IOF[i]<-sum(SFUnico[["Importe de las opciones facturadas"]][1:length(SFUnico[["Acceso"]])])
            IDPT[i]<-sum(SFUnico[["Importe descuentos sobre plano tarifario"]][1:length(SFUnico[["Acceso"]])])
            IOD[i]<-sum(SFUnico[["Importe de las opciones descontadas"]][1:length(SFUnico[["Acceso"]])])
            
          }
          Accesos<-(unique(SFduplicados2["Acceso"]))
          Accesos["Estado acceso"]<-EstAcc
          Accesos["Producto"]<-Prod
          Accesos["Tipo de producto"]<-TdProd
          Accesos["Centro de facturación"]<-CdF
          Accesos["Importe de las opciones facturadas"]<-IOF
          Accesos["Importe descuentos sobre plano tarifario"]<-IDPT
          Accesos["Importe de las opciones descontadas"]<-IOD
          
          SFUnicos<-Accesos
          
        }
        else {
          SFUnicos<-subset(SFACTURADOS,SFACTURADOS[["Acceso"]]==1&SFACTURADOS[["Acceso"]]!=1)
        }
        #############################################################################Parte 5 Uniones
        #SF_Final tiene los No duplicados de la Parte 1, Los duplicados cuyo par tenia cobro 0 de la parte 3 (SFduplicadosbuenos) y los SF Unicos que salen de la consolidacion de los que tienen multiples cobros
        #SF_Apartados tiene los duplicados con cobro
        SF_no_duplicados[["Duplicados"]]<-NULL
        SF_no_duplicados[["Duplicados2"]]<-NULL
         
        SFduplicadosbuenos2<-subset(SFduplicadosbuenos,select = c("Acceso",
                                                                  "Estado acceso",
                                                                  "Producto",
                                                                  "Tipo de producto",
                                                                  "Centro de facturación",
                                                                  "Importe de las opciones facturadas",
                                                                  "Importe descuentos sobre plano tarifario",
                                                                  "Importe de las opciones descontadas"))
        
        if (length(SFUnicos[["Acceso"]])>0){
        SF_Final<-rbind(SFUnicos,SF_no_duplicados,SFduplicadosbuenos2)
        }
        else{
          SF_Final<-rbind(SF_no_duplicados,SFduplicadosbuenos2)
        }
        SFduplicados2<-subset(SFduplicados2,select = c("Acceso",
                                                       "Estado acceso",
                                                       "Producto",
                                                       "Tipo de producto",
                                                       "Centro de facturación",
                                                       "Importe de las opciones facturadas",
                                                       "Importe descuentos sobre plano tarifario",
                                                       "Importe de las opciones descontadas"))
        SF_fueradecontratoSC<-subset(SF_fueradecontratoSC,select = c("Acceso",
                                                                     "Estado acceso",
                                                                     "Producto",
                                                                     "Tipo de producto",
                                                                     "Centro de facturación",
                                                                     "Importe de las opciones facturadas",
                                                                     "Importe descuentos sobre plano tarifario",
                                                                     "Importe de las opciones descontadas"))
        if(length(SFduplicados2[["Acceso"]])>0){
          SFduplicados2<-SFduplicados2[order(x = SFduplicados2[["Acceso"]]),]
          SFduplicados2["Revisar"]<-1
        }
        if(length(SF_fueradecontratoSC[["Acceso"]])>0){
          SF_fueradecontratoSC["Revisar"]<-2
        }
        SF_Apartados<-rbind(SFduplicados2,SF_fueradecontratoSC)
      }
      else{
        SF_Final<-SF_no_duplicados
        SFPlanes2["Revisar"]<-0
        SF_Apartados<-subset(SFPlanes2,SFPlanes2[["Duplicados"]]=="TRUE"|
                               SFPlanes2[["Duplicados2"]]=="TRUE")
        
      }
      SF_Final[["Duplicados"]]<-NULL
      SF_Final[["Duplicados2"]]<-NULL
      #rm(SF_a_evaluar, SF_CPduplicados, SF_en_contrato, SF_fueradecontratoCC,SFduplicadosbuenos,SFduplicadosbuenos2,SF_fueradecontratoSC)
      rm(SF_no_duplicados,SFduplicados,SFduplicados2,SFPlanes2,SFUnicos,SFPlanesA,SFPlanesDb)
      SFPlanes_final<-SF_Final
      if(length(SFOpciones[["Acceso"]])>0){
        
        SF_Final<-rbind(SF_Final,SFOpciones)
      }
      SF_Final<<-SF_Final
      SF_Apartados<<-SF_Apartados
      SFPlanes_final<<-SFPlanes_final
    }
    print("SFACTURADOS LISTO")
    ############################UNION USO - ACCESSES #################
    if(client != "afm"){
    if(!is.null(input$usos)&!is.null(export)){
      if(usuarioid=="Si"){
        backup<-uso
        CantMNG<-as.numeric(length(names(MNG)))
        uso1<-subset(uso,uso[["Usuario ID"]]=="")
        uso2<-subset(uso,uso[["Usuario ID"]]!="")
        if (length(uso1[["Acceso"]])>0){
          uso1[["Usuario ID"]]<-NULL
          uso1[["Usuario ID"]]<-"Sin Ceco"
          uso<<-rbind(uso1,uso2)}
        uso<<-uso
        ACCESSES2<-ACCESSES
        ACCESSES2[["Tipo"]]<-NULL
        ACCESSES2[["Proveedor"]]<-NULL
        ACCESSES2[["Acceso fix"]]<-NULL
        ACCESSES2[["MANAGEMENTORG1"]]<-NULL
        ACCESSES2[["MANAGEMENTORG2"]]<-NULL
        ACCESSES2[["MANAGEMENTORG3"]]<-NULL
        ACCESSES2[["MANAGEMENTORG4"]]<-NULL
        ACCESSES2[["MANAGEMENTORG5"]]<-NULL
        ACCESSES2[["MANAGEMENTORG6"]]<-NULL
        ACCESSES2[["MANAGEMENTORG7"]]<-NULL
        ACCESSES2[["MANAGEMENTORG8"]]<-NULL
        if(CantMNG>=8){
          print("Search by MNG lvl 8")
          if(length(uso[["Acceso"]])>0){
          uso<-merge(uso,MNG,by.x = "Usuario ID",by.y = "MANAGEMENTORG8",all.x = TRUE)
          uso8<-subset(uso,!is.na(uso[["MANAGEMENTORG1"]]))
          uso<-subset(uso,is.na(uso[["MANAGEMENTORG1"]]))
          uso[["MANAGEMENTORG1"]]<-NULL
          uso[["MANAGEMENTORG2"]]<-NULL
          uso[["MANAGEMENTORG3"]]<-NULL
          uso[["MANAGEMENTORG4"]]<-NULL
          uso[["MANAGEMENTORG5"]]<-NULL
          uso[["MANAGEMENTORG6"]]<-NULL
          uso[["MANAGEMENTORG7"]]<-NULL
          uso[["MANAGEMENTORG8"]]<-NULL
          uso8[["MANAGEMENTORG8"]]<-uso8[["Usuario ID"]]
          }
        }
        if(CantMNG>=7){
          print("Search by MNG lvl 7")
          if(length(uso[["Acceso"]])>0){
            uso<-merge(uso,MNG,by.x = "Usuario ID",by.y = "MANAGEMENTORG7",all.x = TRUE)
            uso7<-subset(uso,!is.na(uso[["MANAGEMENTORG1"]]))
            uso<-subset(uso,is.na(uso[["MANAGEMENTORG1"]]))
            uso[["MANAGEMENTORG1"]]<-NULL
            uso[["MANAGEMENTORG2"]]<-NULL
            uso[["MANAGEMENTORG3"]]<-NULL
            uso[["MANAGEMENTORG4"]]<-NULL
            uso[["MANAGEMENTORG5"]]<-NULL
            uso[["MANAGEMENTORG6"]]<-NULL
            uso[["MANAGEMENTORG7"]]<-NULL
            uso[["MANAGEMENTORG8"]]<-NULL
            uso7[["MANAGEMENTORG7"]]<-uso7[["Usuario ID"]]
          }
        }
        if(CantMNG>=6){
          print("Search by MNG lvl 6")
          if(length(uso[["Acceso"]])>0){
            uso<-merge(uso,MNG,by.x = "Usuario ID",by.y = "MANAGEMENTORG6",all.x = TRUE)
            uso6<-subset(uso,!is.na(uso[["MANAGEMENTORG1"]]))
            uso<-subset(uso,is.na(uso[["MANAGEMENTORG1"]]))
            uso[["MANAGEMENTORG1"]]<-NULL
            uso[["MANAGEMENTORG2"]]<-NULL
            uso[["MANAGEMENTORG3"]]<-NULL
            uso[["MANAGEMENTORG4"]]<-NULL
            uso[["MANAGEMENTORG5"]]<-NULL
            uso[["MANAGEMENTORG6"]]<-NULL
            uso[["MANAGEMENTORG7"]]<-NULL
            uso[["MANAGEMENTORG8"]]<-NULL
            uso6[["MANAGEMENTORG6"]]<-uso6[["Usuario ID"]]
          }
        }
        if(CantMNG>=5){
          print("Search by MNG lvl 5")
          if(length(uso[["Acceso"]])>0){
            uso<-merge(uso,MNG,by.x = "Usuario ID",by.y = "MANAGEMENTORG5",all.x = TRUE)
            uso5<-subset(uso,!is.na(uso[["MANAGEMENTORG1"]]))
            uso<-subset(uso,is.na(uso[["MANAGEMENTORG1"]]))
            uso[["MANAGEMENTORG1"]]<-NULL
            uso[["MANAGEMENTORG2"]]<-NULL
            uso[["MANAGEMENTORG3"]]<-NULL
            uso[["MANAGEMENTORG4"]]<-NULL
            uso[["MANAGEMENTORG5"]]<-NULL
            uso[["MANAGEMENTORG6"]]<-NULL
            uso[["MANAGEMENTORG7"]]<-NULL
            uso[["MANAGEMENTORG8"]]<-NULL
            uso5[["MANAGEMENTORG5"]]<-uso5[["Usuario ID"]]
          }
        }
        if(CantMNG>=4){
          print("Search by MNG lvl 4")
          if(length(uso[["Acceso"]])>0){
            uso<-merge(uso,MNG,by.x = "Usuario ID",by.y = "MANAGEMENTORG4",all.x = TRUE)
            uso4<-subset(uso,!is.na(uso[["MANAGEMENTORG1"]]))
            uso<-subset(uso,is.na(uso[["MANAGEMENTORG1"]]))
            uso[["MANAGEMENTORG1"]]<-NULL
            uso[["MANAGEMENTORG2"]]<-NULL
            uso[["MANAGEMENTORG3"]]<-NULL
            uso[["MANAGEMENTORG4"]]<-NULL
            uso[["MANAGEMENTORG5"]]<-NULL
            uso[["MANAGEMENTORG6"]]<-NULL
            uso[["MANAGEMENTORG7"]]<-NULL
            uso[["MANAGEMENTORG8"]]<-NULL
            uso4[["MANAGEMENTORG4"]]<-uso4[["Usuario ID"]]
          }
        }
        if(CantMNG>=3){
          print("Search by MNG lvl 3")
          if(length(uso[["Acceso"]])>0){
            uso<-merge(uso,MNG,by.x = "Usuario ID",by.y = "MANAGEMENTORG3",all.x = TRUE)
            uso3<-subset(uso,!is.na(uso[["MANAGEMENTORG1"]]))
            uso<-subset(uso,is.na(uso[["MANAGEMENTORG1"]]))
            uso[["MANAGEMENTORG1"]]<-NULL
            uso[["MANAGEMENTORG2"]]<-NULL
            uso[["MANAGEMENTORG3"]]<-NULL
            uso[["MANAGEMENTORG4"]]<-NULL
            uso[["MANAGEMENTORG5"]]<-NULL
            uso[["MANAGEMENTORG6"]]<-NULL
            uso[["MANAGEMENTORG7"]]<-NULL
            uso[["MANAGEMENTORG8"]]<-NULL
            uso3[["MANAGEMENTORG3"]]<-uso3[["Usuario ID"]]
          }
        }
        if(CantMNG>=2){
          print("Search by MNG lvl 2")
          if(length(uso[["Acceso"]])>0){
            uso<-merge(uso,MNG,by.x = "Usuario ID",by.y = "MANAGEMENTORG2",all.x = TRUE)
            uso2<-subset(uso,!is.na(uso[["MANAGEMENTORG1"]]))
            uso<-subset(uso,is.na(uso[["MANAGEMENTORG1"]]))
            uso[["MANAGEMENTORG1"]]<-NULL
            uso[["MANAGEMENTORG2"]]<-NULL
            uso[["MANAGEMENTORG3"]]<-NULL
            uso[["MANAGEMENTORG4"]]<-NULL
            uso[["MANAGEMENTORG5"]]<-NULL
            uso[["MANAGEMENTORG6"]]<-NULL
            uso[["MANAGEMENTORG7"]]<-NULL
            uso[["MANAGEMENTORG8"]]<-NULL
            uso2[["MANAGEMENTORG2"]]<-uso2[["Usuario ID"]]
          }
        }
        if(CantMNG>=1){
          print("Search by MNG lvl 1")
          if(length(uso[["Acceso"]])>0){
            uso<-merge(uso,MNG,by.x = "Usuario ID",by.y = "MANAGEMENTORG1",all.x = TRUE)
            uso[["MANAGEMENTORG1"]]<-uso[["Usuario ID"]]
            uso1<-subset(uso,!is.na(uso[["MANAGEMENTORG1"]]))
            uso<-subset(uso,is.na(uso[["MANAGEMENTORG1"]]))
            uso[["MANAGEMENTORG1"]]<-NULL
            uso[["MANAGEMENTORG2"]]<-NULL
            uso[["MANAGEMENTORG3"]]<-NULL
            uso[["MANAGEMENTORG4"]]<-NULL
            uso[["MANAGEMENTORG5"]]<-NULL
            uso[["MANAGEMENTORG6"]]<-NULL
            uso[["MANAGEMENTORG7"]]<-NULL
            uso[["MANAGEMENTORG8"]]<-NULL
          }
        }
        if(CantMNG==8){
          Consolidado<-rbind(uso1,uso2,uso3,uso4,uso5,uso6,uso7,uso8)
        }
        else if(CantMNG==7){
          Consolidado<-rbind(uso1,uso2,uso3,uso4,uso5,uso6,uso7)
        }
        else if(CantMNG==6){
          Consolidado<-rbind(uso1,uso2,uso3,uso4,uso5,uso6)
        }
        else if(CantMNG==5){
          Consolidado<-rbind(uso1,uso2,uso3,uso4,uso5)
        }
        else if(CantMNG==4){
          Consolidado<-rbind(uso1,uso2,uso3,uso4)
        }
        else if(CantMNG==3){
          Consolidado<-rbind(uso1,uso2,uso3)
        }
        else if(CantMNG==2){
          Consolidado<-rbind(uso1,uso2)
        }
        else if(CantMNG==1){
          Consolidado<-uso1
        }
        Consolidado<-merge(Consolidado,ACCESSES2, by = "Acceso",all.x = TRUE)
        uso<-backup
    }
      else if (usuarioid == "No"){
        ACCESSES2<-ACCESSES
        ACCESSES2[["Tipo"]]<-NULL
        ACCESSES2[["Proveedor"]]<-NULL
        ACCESSES2[["Acceso fix"]]<-NULL
        
        Consolidado<-merge(uso,ACCESSES2,by = "Acceso",all.x = TRUE)
      }
    }}
    Consolidado<<-Consolidado #Consolidado contiene uso y ACCESSES
    #################################INFORME DE GESTION DE TELEFONIA#####
    
    if(client == "igm"){
      if(!is.null(input$usos)&!is.null(export)&!is.null(plantilla)&!is.null(input$factura)){
        
        #se modifica plan eliminando las columnas innecesarias
        PLAN2 <- SF_Final
        #Se dejan solo los que son plano tarifario para el analisis
        PLAN2 <- subset(PLAN2,
                        PLAN2[["Tipo de producto"]] == "Plano tarifario")
        PLAN2<-subset(PLAN2,select = c("Acceso","Producto","Importe de las opciones descontadas"))
        PLAN2[["Acceso"]]<-as.character(PLAN2[["Acceso"]])
        Consolidado[["Acceso"]]<-as.character(Consolidado[["Acceso"]])
        Consolidado<- merge(Consolidado,PLAN2,by.x ="Acceso",by.y = "Acceso", all.x = TRUE)
        
        SinUsos<-Consolidado
        SinUsos<-subset(SinUsos,SinUsos[["Tipo"]]!="Centro de facturación")
        SinUsos<<-SinUsos
        
        
        SinUsos[,'usocant']<-SinUsos[,'Voz (sec)']+SinUsos[,'Datos (KB)']+SinUsos[,'N.° SMS/MMS']
        SinUsos <-
          subset(
            SinUsos,
            select = c(
              "Acceso",
              "Fecha",
              "Total",
              "usocant",
              "Meses"
            )
          )
        
        SinUsos<-subset(SinUsos,SinUsos[["usocant"]]==0)
        SinUsos<-subset(SinUsos,SinUsos[["Meses"]]<=3)
        SinUsos<<-SinUsos
        
        #Excepcion Aguas Andinas
        if(!is.null(nombre)){
          if(nombre == "Aguas Andinas"){
            probar<<-subset(Consolidado,Consolidado[["Centro de facturacion"]] == '-')
            if(length(probar[["Acceso"]])>0){
              UAADP_usos2<-subset(Consolidado,Consolidado[["Centro de facturacion"]]!='-')
              probar[["Centro de facturacion"]]<-NULL
              probar[,'Centro de facturacion']<-probar[,'Proveedor Nivel 3']
              Consolidado<-rbind(probar,UAADP_usos2)
              rm(probar,UAADP_usos2)
            }
          }
        }
        
        Consolidado[,'Descuento > Otros']<-Consolidado[,'Descuentos'] - Consolidado[,'Descuento > Plano tarifario']
        Consolidado<<-Consolidado
        dbWriteTable(
          DB,
          "sinusos",
          SinUsos,
          field.types = list(
            `Acceso` = "varchar(255)",
            `Total` = "double(15,2)",
            `usocant` = "double(15,2)",
            `Meses` = "double(15,2)",
            `Fecha` = "varchar(255)"
          ),
          row.names = FALSE,
          overwrite = TRUE,
          append = FALSE,
          allow.keywords = FALSE
        )
        
        
      #source("pj_igm_db.r", local = TRUE)
      }
    }
    #################################INFORME DE GESTION DE IMPRESION######
    if (client == "igi"){
        Consolidado<<-Consolidado
      }
    #################################INFORME DE GESTION DE ENLACES#########
    if (client == "ige"|client == "igm"){

      if(!is.null(input$usos)&!is.null(export)){
        Consolidado1<-subset(Consolidado,Consolidado[["Equipo"]]=="")
        Consolidado2<-subset(Consolidado,Consolidado[["Equipo"]]!="")
        if (length(Consolidado1[["Acceso"]])>0){
          Consolidado1[["Equipo"]]<-NULL
          Consolidado1[["Equipo"]]<-"Otro"
          Consolidado<<-rbind(Consolidado1,Consolidado2)
        }
        
        #Se busca un merge de trazabilidad por proyeccion del periodo actual
        if (customfield == "Si"){
          Consolidado<-merge(Consolidado,CUSTOM,by = "Acceso", all.x = TRUE)}
        else if (customfield=="No"){
          Consolidado<-uso
        }
        Consolidado<<-Consolidado
        
      }
    }
    #################################ANOMALIAS DE FACTURACION MOVIL########
    if (client == "afm"){
      if(!is.null(input$usos) & !is.null(input$factura) & !is.null(input$cdr) & !is.null(plantilla) & !is.null(contrato)){
        {
          cdr2<<-subset(cdr,
                        (cdr[["Geografia"]]!="Nacional desconocido"
                         & cdr[["Tipo de llamada"]]!="SMS"))
          cdr3<<-subset(cdr,
                        (cdr[["Servicio llamado"]]=="Números especiales" 
                         & cdr[["Tipo de llamada"]]!="SMS"))
          #source("pj_afm.r", local = TRUE)
          {
            ################Consolidado############
            break
            SFPlanes_final<-SFPlanes_final
            Fact<<-merge(uso,
                         SFPlanes_final,
                         by = c("Acceso",
                                "Centro de facturacion"),
                         all.x = TRUE)
            facturas2<-facturas
            facturas2[,'Proveedor']<-NULL
            facturas2[,'Total sin impuestos']<-NULL
            facturas2[,'Total imp. incluidos']<-NULL
            facturas2[,'Importe IVA']<-NULL
            facturas2[,'Divisa']<-NULL
            facturas2[,'N. accesos facturados']<-NULL
            facturas2[,'Fecha']<-NULL
            facturas2<<-facturas2
            Fact<<-merge(Fact,facturas2,by = "Centro de facturacion", all.x = TRUE)
            Fact[["Acceso fix"]]<-NULL
            Contratoplanes<-subset(MOVISTAR_PLANES,
                                   select = c("Producto",
                                              "Tipo",
                                              "Precio (CLP)",
                                              "Voz (min)",
                                              "Datos (KB)",
                                              "Precio/min (CLP)",
                                              "PrecioSC/min (CLP)",
                                              "Precio/SMS (CLP)",
                                              "Precio/KB (CLP)"))
            Contratoplanes[,'Tipo Contrato']<-Contratoplanes[,'Tipo']
            Contratoplanes[["Tipo"]]<-NULL
            Consolidado<<-merge(Fact,Contratoplanes,by = "Producto",all.x = TRUE)
            Consolidado[,'Voz nac. (min)']<-Consolidado[,'Voz nacional (seg)']/60
            ####################MIN ADICIONAL#################
            
            MIN_ADICIONAL1<<-subset(Consolidado,Consolidado[["Voz nac. (min)"]]>Consolidado[["Voz (min)"]]&Consolidado[["Voz (min)"]]>0)
            MIN_ADICIONAL1[,'Delta minutos']<-MIN_ADICIONAL1[,'Voz nac. (min)']-MIN_ADICIONAL1[,'Voz (min)']
            MIN_ADICIONAL1[,'Precio Real']<- (MIN_ADICIONAL1[,'Voz nac. (min)']-MIN_ADICIONAL1[,'Voz (min)'])*MIN_ADICIONAL1[,'PrecioSC/min (CLP)']
            MIN_ADICIONAL1[,'Delta']<-MIN_ADICIONAL1[,'Voz nacional (CLP)']-MIN_ADICIONAL1[,'Precio Real']
            MIN_ADICIONAL2<-subset(Consolidado,Consolidado[["Voz nac. (min)"]]<Consolidado[["Voz (min)"]]&Consolidado[["Voz nacional (CLP)"]]>0&Consolidado[["PrecioSC/min (CLP)"]]>0)
            MIN_ADICIONAL2[,'Delta minutos']<-MIN_ADICIONAL2[,'Voz nac. (min)']-MIN_ADICIONAL2[,'Voz (min)']
            MIN_ADICIONAL2[,'Precio Real']<- 0
            MIN_ADICIONAL2[,'Delta']<-MIN_ADICIONAL2[,'Voz nacional (CLP)']-MIN_ADICIONAL2[,'Precio Real']
            MIN_ADICIONAL2<<-MIN_ADICIONAL2
            cdr4<-subset(cdr3,select = c("Numero de llamada","Servicio llamado"))
            
            MIN_ADICIONAL<-rbind(MIN_ADICIONAL1,MIN_ADICIONAL2)
            
            MIN_ADICIONAL<-merge(MIN_ADICIONAL,cdr3,by.x = "Acceso",by.y = "Numero de llamada",all.x = TRUE)
            a<-duplicated(MIN_ADICIONAL[["Acceso"]],fromLast = FALSE)
            MIN_ADICIONAL[["Duplicados"]]<-a
            MIN_ADICIONAL<-subset(MIN_ADICIONAL,MIN_ADICIONAL[["Duplicados"]]=="FALSE")
            MIN_ADICIONAL[["Duplicados"]]<-NULL
            MIN_ADICIONAL<<-MIN_ADICIONAL
            print("Total min adicional")
            print(sum(MIN_ADICIONAL[["Delta"]]))
            ##########################CONTRATO PLAN Y A GRANEL##############
            PlanContrato<<-subset(Consolidado,Consolidado[["Importe de las opciones descontadas (CLP)"]]!=Consolidado[["Precio (CLP)"]]&Consolidado[["Precio (CLP)"]]!=0&Consolidado[["Estado acceso"]]=="Activo")
            PlanContrato[,'Delta']<<-PlanContrato[,'Importe de las opciones descontadas (CLP)']-PlanContrato[,'Precio (CLP)']
            PlanContrato<<-subset(PlanContrato,PlanContrato[["Delta"]]>0)
            print("Total Plan Contrato")
            print(sum(PlanContrato[["Delta"]]))
            PlanContratoGranel<<-subset(Consolidado,Consolidado[["Precio/min (CLP)"]]>0&Consolidado[["Estado acceso"]]=="Activo")
            PlanContratoGranel[,'Precio real (CLP)']<<-(PlanContratoGranel[,'Voz nacional (seg)']/60)*PlanContratoGranel[,'Precio/min (CLP)']
            PlanContratoGranel[,'Delta']<<-PlanContratoGranel[["Voz nacional (CLP)"]]+PlanContratoGranel[["Importe de las opciones descontadas (CLP)"]]-((PlanContratoGranel[,'Voz nacional (seg)']/60)*PlanContratoGranel[,'Precio/min (CLP)'])
            PlanContratoGranel<<-subset(PlanContratoGranel,PlanContratoGranel[["Delta"]]>0)
            print("Total Plan Granel")
            print(sum(PlanContratoGranel[["Delta"]]))
            ##############################PAISES ROAMING VOZ####################
            
            RoaVoz<-subset(Consolidado,
                           Consolidado[["Estado acceso"]]=="Activo"&
                             Consolidado[["Voz roaming (CLP)"]]>0)
            
            cdrA<-subset(cdr2,cdr2[["Tipo de llamada"]]=="Voz")
            cdrB<-cdrA
            a<-duplicated(cdrA[["Pais emisor"]])
            b<-duplicated(cdrB[["Pais destinatario"]])
            cdrA[["duplicados"]]<-a
            cdrA<-subset(cdrA,cdrA[["duplicados"]]=="FALSE")
            
            cdrB[["duplicados"]]<-b
            cdrB<-subset(cdrB,cdrB[["duplicados"]]=="FALSE")
            paisesA<-data.frame(cdrA[["Pais emisor"]])
            names(paisesA)<-c("Pais")
            paisesB<-data.frame(cdrB[["Pais destinatario"]])
            names(paisesB)<-c("Pais")
            paises<-rbind(paisesA,paisesB)
            a<-duplicated(paises[["Pais"]])
            paises[["duplicados"]]<-a
            paises<-subset(paises,paises[["duplicados"]]=="FALSE")
            rm(cdrA,cdrB,paisesA,paisesB,a,b)
            
            PMatriz<-matrix(nrow = length(RoaVoz[["Acceso"]]),ncol = length(paises[["Pais"]]))
            
            for(i in 1:length(RoaVoz[["Acceso"]])){
              cdrC<-subset(cdr2,cdr2[["Tipo de llamada"]]=="Voz" &
                             cdr2[["Numero de llamada"]]==RoaVoz[["Acceso"]][i])
              for(j in 1:length(paises[["Pais"]])){
                P<-as.character(paises[["Pais"]][j])
                for(k in 1:length(cdrC[["Numero de llamada"]])){
                  Pe<-as.character(cdrC[["Pais emisor"]][k])
                  Pd<-as.character(cdrC[["Pais destinatario"]][k])
                  if(Pe==P){
                    PMatriz[i,j]<-P
                  }
                  else if(Pd==P){
                    PMatriz[i,j]<-P
                  }
                }
              }
            }
            PMatriz<-data.frame(PMatriz)
            names(PMatriz)<-paises[["Pais"]]
            RoaVoz<-subset(RoaVoz,select = c("Acceso",
                                             "Estado acceso",
                                             "Tipo",
                                             "Producto",
                                             "Centro de facturacion",
                                             "Cuenta cliente",
                                             "Factura",
                                             "Total (CLP)",
                                             "Plano tarifario (CLP)",
                                             "Servicios (CLP)",
                                             "Servicios opciones (CLP)",
                                             "Servicios otros (CLP)",
                                             "Descuentos (CLP)",
                                             "Descuentos opciones (CLP)",
                                             "Descuentos otros (CLP)",
                                             "Uso rebajado (CLP)",
                                             "Voz (CLP)",
                                             "Voz roaming (CLP)",
                                             "Voz roaming (seg)"))
            RoaVoz<-cbind(RoaVoz,PMatriz)
            RoaVoz<<-RoaVoz
            
            ##############################PAISES ROAMING DATOS ###################
            
            RoaDat<-subset(Consolidado,
                           Consolidado[["Estado acceso"]]=="Activo"&
                             Consolidado[["Datos inter (CLP)"]]>0)
            cdrA<-subset(cdr2,cdr2[["Tipo de llamada"]]=="Datos")
            a<-duplicated(cdrA[["Pais emisor"]])
            cdrA[["duplicados"]]<-a
            cdrA<-subset(cdrA,cdrA[["duplicados"]]=="FALSE")
            paises<-data.frame(cdrA[["Pais emisor"]])
            names(paises)<-c("Pais")
            rm(cdrA,cdrB,a)
            PMatriz<-matrix(nrow = length(RoaDat[["Acceso"]]),ncol = length(paises[["Pais"]]))
            for(i in 1:length(RoaDat[["Acceso"]])){
              cdrC<-subset(cdr2,cdr2[["Tipo de llamada"]]=="Datos" &
                             cdr2[["Numero de llamada"]]==RoaDat[["Acceso"]][i])
              for(j in 1:length(paises[["Pais"]])){
                P<-as.character(paises[["Pais"]][j])
                for(k in 1:length(cdrC[["Numero de llamada"]])){
                  Pe<-as.character(cdrC[["Pais emisor"]][k])
                  Pd<-as.character(cdrC[["Pais destinatario"]][k])
                  if(Pe==P){
                    PMatriz[i,j]<-P
                  }
                  else if(Pd==P){
                    PMatriz[i,j]<-P
                  }
                }
              }
            }
            PMatriz<-data.frame(PMatriz)
            names(PMatriz)<-paises[["Pais"]]
            RoaDat<-subset(RoaDat,select = c("Acceso",
                                             "Estado acceso",
                                             "Tipo",
                                             "Producto",
                                             "Centro de facturacion",
                                             "Cuenta cliente",
                                             "Factura",
                                             "Total (CLP)",
                                             "Plano tarifario (CLP)",
                                             "Servicios (CLP)",
                                             "Servicios opciones (CLP)",
                                             "Servicios otros (CLP)",
                                             "Descuentos (CLP)",
                                             "Descuentos opciones (CLP)",
                                             "Descuentos otros (CLP)",
                                             "Uso rebajado (CLP)",
                                             "Datos (CLP)",
                                             "Datos inter (CLP)",
                                             "Datos inter (KB)"))
            RoaDat<-cbind(RoaDat,PMatriz)
            RoaDat<<-RoaDat
            
          }
          ############Consolidado###########
          dbWriteTable(
            DB,
            "consolidado",
            Fact,
            field.types = NULL ,
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
          ##################Servicios Facturados###############
          SFPlanes<-subset(SFPlanes_final,select = c("Acceso","Estado acceso","Producto","Tipo de producto","Centro de facturacion","Importe de las opciones facturadas (CLP)",
                                                     "Importe descuentos sobre plano tarifario (CLP)","Importe de las opciones descontadas (CLP)","Acceso fix"))
          SFOpciones<-subset(SFOpciones,select = c("Acceso","Estado acceso","Producto","Tipo de producto","Centro de facturacion","Importe de las opciones facturadas (CLP)",
                                                   "Importe descuentos sobre plano tarifario (CLP)","Importe de las opciones descontadas (CLP)","Acceso fix"))
          SF_Final<-subset(SF_Final,select = c("Acceso","Estado acceso","Producto","Tipo de producto","Centro de facturacion","Importe de las opciones facturadas (CLP)",
                                               "Importe descuentos sobre plano tarifario (CLP)","Importe de las opciones descontadas (CLP)","Acceso fix"))
          SF_Apartados<-subset(SF_Apartados,select = c("Acceso","Estado acceso","Producto","Tipo de producto","Centro de facturacion","Importe de las opciones facturadas (CLP)",
                                                       "Importe descuentos sobre plano tarifario (CLP)","Importe de las opciones descontadas (CLP)","Acceso fix","Revisar"))
          dbWriteTable(
            DB,
            "sf_planes",
            SFPlanes,
            field.types = list(
              `Acceso` = "varchar(255)",
              `Estado acceso` = "varchar(255)",
              `Producto` = "varchar(255)",
              `Tipo de producto` = "varchar(255)",
              `Centro de facturacion` = "varchar(255)",
              `Importe de las opciones facturadas (CLP)` = "double(15,2)",
              `Importe descuentos sobre plano tarifario (CLP)` = "double(15,2)",
              `Importe de las opciones descontadas (CLP)` = "double(15,2)",
              `Acceso fix` = "varchar(255)"
            ),
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
          dbWriteTable(
            DB,
            "sf_opciones",
            SFOpciones,
            field.types = list(
              `Acceso` = "varchar(255)",
              `Estado acceso` = "varchar(255)",
              `Producto` = "varchar(255)",
              `Tipo de producto` = "varchar(255)",
              `Centro de facturacion` = "varchar(255)",
              `Importe de las opciones facturadas (CLP)` = "double(15,2)",
              `Importe descuentos sobre plano tarifario (CLP)` = "double(15,2)",
              `Importe de las opciones descontadas (CLP)` = "double(15,2)",
              `Acceso fix` = "varchar(255)"
            ),
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
          dbWriteTable(
            DB,
            "sf_final",
            SF_Final,
            field.types = list(
              `Acceso` = "varchar(255)",
              `Estado acceso` = "varchar(255)",
              `Producto` = "varchar(255)",
              `Tipo de producto` = "varchar(255)",
              `Centro de facturacion` = "varchar(255)",
              `Importe de las opciones facturadas (CLP)` = "double(15,2)",
              `Importe descuentos sobre plano tarifario (CLP)` = "double(15,2)",
              `Importe de las opciones descontadas (CLP)` = "double(15,2)",
              `Acceso fix` = "varchar(255)"
            ),
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
          dbWriteTable(
            DB,
            "sf_apartados",
            SF_Apartados,
            field.types = list(
              `Acceso` = "varchar(255)",
              `Estado acceso` = "varchar(255)",
              `Producto` = "varchar(255)",
              `Tipo de producto` = "varchar(255)",
              `Centro de facturacion` = "varchar(255)",
              `Importe de las opciones facturadas (CLP)` = "double(15,2)",
              `Importe descuentos sobre plano tarifario (CLP)` = "double(15,2)",
              `Importe de las opciones descontadas (CLP)` = "double(15,2)",
              `Acceso fix` = "varchar(255)",
              `Revisar` = "double(15,2)"
            ),
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
          ##################Min Adicional###################
          MIN_ADICIONAL<-subset(MIN_ADICIONAL,select = c(
            "Acceso",
            "Estado acceso",
            "Tipo",
            "Producto",
            "Centro de facturacion",
            "Cuenta cliente",
            "Factura",
            "Total (CLP)",
            "Plano tarifario (CLP)",
            "Voz (CLP)",
            "Voz nacional (CLP)",
            "N. Voz nacional",
            "Voz nacional (seg)",
            "Voz nac. (min)",
            "Voz (min)",
            "Delta minutos",
            "Precio Real",
            "Servicio llamado",
            "Delta"))
          MIN_ADICIONAL<<-MIN_ADICIONAL
          dbWriteTable(
            DB,
            "min_adicional",
            MIN_ADICIONAL,
            field.types = list(
              `Acceso` = "varchar(255)",
              `Estado acceso` = "varchar(255)",
              `Tipo` = "varchar(255)",
              `Producto` = "varchar(255)",
              `Centro de facturacion` = "varchar(255)",
              `Cuenta cliente` = "varchar(255)",
              `Factura` = "varchar(255)",
              `Total (CLP)` = "double(15,2)",
              `Plano tarifario (CLP)` = "double(15,2)",
              `Voz (CLP)` = "double(15,2)",
              `Voz nacional (CLP)` = "double(15,2)",
              `N. Voz nacional` = "double(15,2)",
              `Voz nacional (seg)` = "double(15,2)",
              `Voz nac. (min)` = "double(15,2)",
              `Voz (min)` = "double(15,2)",
              `Delta minutos` = "double(15,2)",
              `Precio Real` = "double(15,2)",
              `Servicio llamado` = "varchar(255)",
              `Delta` = "double(15,2)"
            ),
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
          #################Plan Contrato###################
          PlanContrato<-subset(PlanContrato,select = c("Acceso",
                                                       "Estado acceso",
                                                       "Tipo Contrato",
                                                       "Producto",
                                                       "Centro de facturacion",
                                                       "Cuenta cliente",
                                                       "Factura",
                                                       "Importe de las opciones facturadas (CLP)",
                                                       "Importe descuentos sobre plano tarifario (CLP)",
                                                       "Importe de las opciones descontadas (CLP)",
                                                       "Precio (CLP)",
                                                       "Delta"))
          PlanContrato<<-PlanContrato
          dbWriteTable(
            DB,
            "plancontrato",
            PlanContrato,
            field.types = list(
              `Acceso` = "varchar(255)",
              `Estado acceso` = "varchar(255)",
              `Tipo Contrato` = "varchar(255)",
              `Producto` = "varchar(255)",
              `Centro de facturacion` = "varchar(255)",
              `Cuenta cliente` = "varchar(255)",
              `Factura` = "varchar(255)",
              `Importe de las opciones facturadas (CLP)` = "double(15,2)",
              `Importe descuentos sobre plano tarifario (CLP)` = "double(15,2)",
              `Importe de las opciones descontadas (CLP)` = "double(15,2)",
              `Precio (CLP)` = "double(15,2)",
              `Delta` = "double(15,2)"
            ),
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
          #################Plan Granel#####################
          PlanContratoGranel<-subset(PlanContratoGranel,select = c("Acceso",
                                                                   "Estado acceso",
                                                                   "Tipo Contrato",
                                                                   "Producto",
                                                                   "Centro de facturacion",
                                                                   "Cuenta cliente",
                                                                   "Factura",
                                                                   "Importe de las opciones facturadas (CLP)",
                                                                   "Importe descuentos sobre plano tarifario (CLP)",
                                                                   "Importe de las opciones descontadas (CLP)",
                                                                   "Precio/min (CLP)",
                                                                   "Precio (CLP)",
                                                                   "Voz nac. (min)",
                                                                   "Precio real (CLP)",
                                                                   "Delta"))
          PlanContratoGranel<<-PlanContratoGranel
          dbWriteTable(
            DB,
            "plangranel",
            PlanContratoGranel,
            field.types = list(
              `Acceso` = "varchar(255)",
              `Estado acceso` = "varchar(255)",
              `Tipo Contrato` = "varchar(255)",
              `Producto` = "varchar(255)",
              `Centro de facturacion` = "varchar(255)",
              `Cuenta cliente` = "varchar(255)",
              `Factura` = "varchar(255)",
              `Importe de las opciones facturadas (CLP)` = "double(15,2)",
              `Importe descuentos sobre plano tarifario (CLP)` = "double(15,2)",
              `Importe de las opciones descontadas (CLP)` = "double(15,2)",
              `Precio/min (CLP)` = "double(15,2)",
              `Precio (CLP)` = "double(15,2)",
              `Voz nac. (min)` = "double(15,2)",
              `Precio real (CLP)` = "double(15,2)",
              `Delta` = "double(15,2)"
            ),
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
          dbWriteTable(
            DB,
            "cdr_relevante",
            cdr2,
            field.types = list(
              `Numero de llamada` = "char(11)",
              `Numero de llamada fix` = "int(9)",
              `Numero llamado` = "varchar(20)",
              `Tipo de llamada` = "ENUM('Datos','MMS','SMS','Voz','E-mail','Desconocidos') NOT NULL",
              `Fecha de llamada` = "DATE",
              `Geografia` = "ENUM('A internacional','Local','Regional','Nacional desconocido','Roaming entrante','Roaming saliente','Roaming desconocido','Internacional desconocido','Desconocidos') NOT NULL",
              `Pais emisor` = "varchar(40)",
              `Pais destinatario` = "varchar(40)",
              `Duracion` = "SMALLINT(8) UNSIGNED NOT NULL",
              `Volumen` = "MEDIUMINT(10) UNSIGNED NOT NULL",
              `Precio` = "FLOAT(10,2) NOT NULL",
              `Tarificacion` = "ENUM('En el plan','Más alla del plan','Desconocidos','Fuera de plan') NOT NULL",
              `Servicio llamado` = "varchar(255)",
              `Organizacion Proveedor` = "varchar(255)"
            ),
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
          #################Roaming Voz###########
          dbWriteTable(
            DB,
            "roaming_voz",
            RoaVoz,
            field.types = NULL ,
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
          #################Roaming Datos########
          dbWriteTable(
            DB,
            "roaming_datos",
            RoaDat,
            field.types = NULL ,
            row.names = FALSE,
            overwrite = TRUE,
            append = FALSE,
            allow.keywords = FALSE
          )
        }
      }
    }
    ####################UPLOAD CONSOLIDADO############
    if(!is.null(Consolidado)){
      dbGetQuery(DB, "SET NAMES 'latin1';")
    dbWriteTable(
      DB,
      "consolidado",
      Consolidado,
      field.types = NULL ,
      row.names = FALSE,
      overwrite = TRUE,
      append = FALSE,
      allow.keywords = FALSE
    )
    }
    
    
    #Run the following code if theres a ticket in the RFP excel checkbox
    #####RFP##########
    if (input$excel == TRUE) {
      #open RFP Workbook
      wb <-
        loadWorkbook("Z:\\AUT Informes\\Licitacion SAAM\\RFP TELEFONIA MOVIL.xlsx")
      
      #Insert Nombre licitacion to Workbook
      licitacion <-
        paste0("SOLUCIÓN DE TELECOMUNICACIONES MÓVILES PARA ",
               input$nombre)
      
      writeData(
        wb,
        sheet = "RFP MOVISTAR",
        licitacion,
        startCol = 2,
        startRow = 6
      )
      writeData(
        wb,
        sheet = "RFP ENTEL",
        licitacion,
        startCol = 2,
        startRow = 6
      )
      writeData(
        wb,
        sheet = "RFP PLANES",
        licitacion,
        startCol = 2,
        startRow = 6
      )
      writeData(
        wb,
        sheet = "RFP CATEGORIAS PLANES",
        licitacion,
        startCol = 2,
        startRow = 6
      )
      writeData(
        wb,
        sheet = "RFP EQUIPOS",
        licitacion,
        startCol = 2,
        startRow = 6
      )
      
      #Insert Nombre cliente to Workbook
      cliente <- input$nombre
      
      writeData(
        wb,
        sheet = "RFP MOVISTAR",
        cliente,
        startCol = 2,
        startRow = 7
      )
      writeData(
        wb,
        sheet = "RFP ENTEL",
        cliente,
        startCol = 2,
        startRow = 7
      )
      writeData(
        wb,
        sheet = "RFP PLANES",
        cliente,
        startCol = 2,
        startRow = 7
      )
      writeData(
        wb,
        sheet = "RFP CATEGORIAS PLANES",
        cliente,
        startCol = 2,
        startRow = 7
      )
      writeData(
        wb,
        sheet = "RFP EQUIPOS",
        cliente,
        startCol = 2,
        startRow = 7
      )
      
      #Insert Date to Workbook
      fecha <- input$fecha
      
      writeData(
        wb,
        sheet = "RFP MOVISTAR",
        fecha,
        startCol = 2,
        startRow = 9
      )
      writeData(
        wb,
        sheet = "RFP ENTEL",
        fecha,
        startCol = 2,
        startRow = 9
      )
      writeData(
        wb,
        sheet = "RFP PLANES",
        fecha,
        startCol = 2,
        startRow = 9
      )
      writeData(
        wb,
        sheet = "RFP CATEGORIAS PLANES",
        fecha,
        startCol = 2,
        startRow = 9
      )
      writeData(
        wb,
        sheet = "RFP EQUIPOS",
        fecha,
        startCol = 2,
        startRow = 9
      )
      
      #Insert client image to Workbook
      #Por Hacer#Si no inserta link, ir a buscar el ultimo de la base de datos
      link <- input$link
      z <- tempfile()
      download.file(link, z, mode = "wb")
      
      insertImage(
        wb,
        sheet = "RFP MOVISTAR",
        z,
        startCol = 3,
        startRow = 1
      )
      insertImage(
        wb,
        sheet = "RFP ENTEL",
        z,
        startCol = 3,
        startRow = 1
      )
      insertImage(
        wb,
        sheet = "RFP PLANES",
        z,
        startCol = 3,
        startRow = 1
      )
      insertImage(
        wb,
        sheet = "RFP CATEGORIAS PLANES",
        z,
        startCol = 3,
        startRow = 1
      )
      insertImage(
        wb,
        sheet = "RFP EQUIPOS",
        z,
        startCol = 3,
        startRow = 1,
      )
      
      #Functions in other files
      #source("pj.r", local = TRUE)
      source("pb.r", local = TRUE)
      #source("jp.r", local = TRUE)
      
      
      # Consumo total Voz MOVISTAR
       writeData(wb, sheet = "RFP MOVISTAR", MovTotMin, startCol = 4, startRow = 14)
      #
      # Consumo voz entre usuarios MOVISTAR
       writeData(wb, sheet = "RFP MOVISTAR", MovVozOnNet, startCol = 4, startRow = 16)
      #
      # Consumo voz a todo destino MOVISTAR
       writeData(wb, sheet = "RFP MOVISTAR", MovATodDes, startCol = 4, startRow = 18)
      #
      # BAM o Servicios de Telemetria MOVISTAR
      # writeData(wb, sheet = "RFP MOVISTAR", X, startCol = 4, startRow = 26)
      #
      # Mensajeria SMS MOVISTAR
       writeData(wb, sheet = "RFP MOVISTAR", MovSms, startCol = 4, startRow = 28)
      #
      # Mensajeria MMS MOVISTAR
       writeData(wb, sheet = "RFP MOVISTAR", MovMms, startCol = 4, startRow = 30)
      #
      # Usuarios Roaming On Demand MOVISTAR
      # writeData(wb, sheet = "RFP MOVISTAR", X, startCol = 4, startRow = 32)
      #
      # Roaming Voz MOVISTAR
       writeData(wb, sheet = "RFP MOVISTAR", MovRoaVoz, startCol = 4, startRow = 34)
      #
      # Roaming Datos MOVISTAR
       writeData(wb, sheet = "RFP MOVISTAR", MovRoaDat, startCol = 4, startRow = 36)
      #
      # Roaming Mensajes MOVISTAR
       writeData(wb, sheet = "RFP MOVISTAR", MovRoaSms, startCol = 4, startRow = 38)
      #
      # $/Minuto Actual MOVISTAR
       writeData(wb, sheet = "RFP MOVISTAR", MovMinAct, startCol = 8, startRow = 15)
      #
      # $/Mb Actual MOVISTAR
      # writeData(wb, sheet = "RFP MOVISTAR", X, startCol = 8, startRow = 20)
      
      
      # Consumo total Voz ENTEL
       writeData(wb, sheet = "RFP ENTEL", EntTotMin, startCol = 4, startRow = 14)
      #
      # Consumo voz entre usuarios ENTEL
      writeData(wb, sheet = "RFP ENTEL", EntVozOnNet, startCol = 4, startRow = 16)
      #
      # Consumo voz a todo destino ENTEL
      writeData(wb, sheet = "RFP ENTEL", EntATodDes, startCol = 4, startRow = 18)
      #
      # BAM o Servicios de Telemetria ENTEL
      # writeData(wb, sheet = "RFP ENTEL", X, startCol = 4, startRow = 26)
      #
      # Mensajeria SMS ENTEL
       writeData(wb, sheet = "RFP ENTEL", EntSms, startCol = 4, startRow = 28)
      #
      # Mensajeria MMS ENTEL
       writeData(wb, sheet = "RFP ENTEL", EntMms, startCol = 4, startRow = 30)
      #
      # Usuarios Roaming On Demand ENTEL
      # writeData(wb, sheet = "RFP ENTEL", X, startCol = 4, startRow = 32)
      #
      # Roaming Voz ENTEL
       writeData(wb, sheet = "RFP ENTEL", EntRoaVoz, startCol = 4, startRow = 34)
      #
      # Roaming Datos ENTEL
       writeData(wb, sheet = "RFP ENTEL", EntRoaDat, startCol = 4, startRow = 36)
      #
      # Roaming Mensajes ENTEL
       writeData(wb, sheet = "RFP ENTEL", EntRoaSms, startCol = 4, startRow = 38)
      #
      # Internacional Voz ENTEL
       writeData(wb, sheet = "RFP ENTEL", EntIntVoz, startCol = 4, startRow = 40)
      #
      # $/Minuto promedio ENTEL
       writeData(wb, sheet = "RFP ENTEL", EntMinAct, startCol = 8, startRow = 16)
      #
      # $/Mb Actual ENTEL
       writeData(wb, sheet = "RFP ENTEL", EntMbAct, startCol = 8, startRow = 21)
      
      #Save Workbook
      saveWorkbook(
        wb,
        paste0(
          "Z:\\AUT Informes\\Licitacion SAAM\\RFP TELEFONIA MOVIL ",
          cliente,
          ".xlsx"
        ),
        overwrite = T
      )
    }
##############END CONECTION ###########   
    #Upload of divisa
    {
      divisa<-data.frame(divisa)
      dbWriteTable(
        DB,
        "divisa",
        divisa,
        field.types = NULL,
        row.names = FALSE,
        overwrite = TRUE,
        append = FALSE,
        allow.keywords = FALSE
      )}
    
    #Kill open connections
    killDbConnections()
    
    #Update the excecute button to "finished" status
    updateButton(session,
                 "execute",
                 style = "success",
                 icon = icon("check"))
    
   
  })
  })