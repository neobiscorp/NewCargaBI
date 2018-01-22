{
  ###########################CARGAR EL DATASET FINAL IGM
##PARTE COMENTADA QUE NO HAY QUE SACAR EN CASO DE QUE HAYA ALGUNA FALLA EN ALGUN LADO 
#   library(shiny)      #Dashboard
#   library(RMySQL)     #Secundary MySQL connection, Kill connections
#   library(DBI)        #Primary MySQL connection
#   library(openxlsx)   #Read xlsx files
#   library(data.table) #For merge, rbind and dataframe works
#   killDbConnections <- function () {
#     all_cons <- dbListConnections(MySQL())
#     print(all_cons)
#     for (con in all_cons)
#       +  dbDisconnect(con)
#     print(paste(length(all_cons), " connections killed."))
#   }
#   DB <- dbConnect(
#     MySQL(),
#     user = "root",
#     password = "",
#     dbname = paste0("igm")
#   )
  print("Base de datos conectada")
#Se crea la tabla sinusos que contiene los datos que no tienen uso de los tres ultimos periodos
SinUsos <-
  subset(
    SinUsos,
    select = c(
      "Acceso",
      "Fecha",
      "Total (CLP)",
      "Meses",
      "usocant"
    )
  )
SinUsos2<<-SinUsos
#Se sube a la base de datos
  dbWriteTable(
    DB,
    "sinusos",
    SinUsos,
    field.types = list(
      `Acceso` = "varchar(255)",
      `Total (CLP)` = "double(15,2)",
      `usocant` = "double(15,2)",
      `Meses` = "double(15,2)",
      `Fecha` = "varchar(255)"
    ),
    row.names = FALSE,
    overwrite = TRUE,
    append = FALSE,
    allow.keywords = FALSE
  )

  #####Tabla UAA_users #####
  #Se eliminan las columnas sobrantes y se cambia el nombre de la columna modelo por equipo
  UAADPT_users2<-UAADP_usos
  UAADPT_users2[,'Modelo']<- UAADPT_users2[["Equipo"]]
  UAADPT_users2[["Equipo"]]<-NULL
  UAADPT_users2[["Acceso fix"]]<-NULL
  UAADPT_users2[["Estado"]]<-NULL
  UAADPT_users2[["Mes"]]<-NULL
  UAADPT_users2[["Tipo de producto"]]<- NULL

  
  #Se sube la tabla de datos con la condicion de que agrega la cantidad de management org que corresponde
  if (is.null(UAADPT_users2[["MANAGEMENTORG1"]])==TRUE){
  dbWriteTable(
    DB,
    "usos",
    UAADPT_users2,
    field.types = list(
      
      `Acceso` = "varchar(255)",
      `Usuario` = "varchar(255)",
      `Centro de facturacion` = "varchar(255)",
      `Modelo` = "varchar(255)",
      `Tipo` = "varchar(255)",
      `Proveedor` = "varchar(255)",
      `Proveedor Nivel 2` = "varchar(255)",
      `Proveedor Nivel 3` = "varchar(255)",
      `Total (CLP)` = "double(15,2)",
      `Plano tarifario (CLP)` = "double(15,2)",
      `Uso (CLP)` = "double(15,2)",
      `Servicios (CLP)` = "double(15,2)",
      `Descuentos (CLP)` = "double(15,2)",
      `Descuento de Plano tarifario (CLP)` = "double(15,2)",
      `Voz (CLP)` = "double(15,2)",
      `Voz nacional (CLP)` = "double(15,2)",
      `Voz inter (CLP)` = "double(15,2)",
      `Datos (CLP)` = "double(15,2)",
      `Datos nacional (CLP)` = "double(15,2)",
      `Datos inter (CLP)` = "double(15,2)",
      `SMS/MMS (CLP)` = "double(15,2)",
      `SMS/MMS nacional (CLP)` = "double(15,2)",
      `SMS/MMS inter (CLP)` = "double(15,2)",
      `Voz (seg)` = "double(15,2)",
      `Voz nacional (seg)` = "double(15,2)",
      `Voz inter (seg)` = "double(15,2)",
      `Datos (KB)` = "double(15,2)",
      `Datos nacional (KB)` = "double(15,2)",
      `Datos inter (KB)` = "double(15,2)",
      `N. SMS/MMS` = "double(15,2)",
      `Fecha` = "varchar(255)"
      
      
    ),
    row.names = FALSE,
    overwrite = TRUE,
    append = FALSE,
    allow.keywords = FALSE
  )
  }
  if (is.null(UAADPT_users2[["MANAGEMENTORG2"]])==TRUE&
      is.null(UAADPT_users2[["MANAGEMENTORG1"]])==FALSE){
    dbWriteTable(
      DB,
      "usos",
      UAADPT_users2,
      field.types = list(
        
        
        `Acceso` = "varchar(255)",
        `Usuario` = "varchar(255)",
        `Centro de facturacion` = "varchar(255)",
        `Modelo` = "varchar(255)",
        `Tipo` = "varchar(255)",
        `Proveedor` = "varchar(255)",
        `Proveedor Nivel 2` = "varchar(255)",
        `Proveedor Nivel 3` = "varchar(255)",
        `MANAGEMENTORG1` = "varchar(255)",
        `Total (CLP)` = "double(15,2)",
        `Plano tarifario (CLP)` = "double(15,2)",
        `Uso (CLP)` = "double(15,2)",
        `Servicios (CLP)` = "double(15,2)",
        `Descuentos (CLP)` = "double(15,2)",
        `Descuento de Plano tarifario (CLP)` = "double(15,2)",
        `Voz (CLP)` = "double(15,2)",
        `Voz nacional (CLP)` = "double(15,2)",
        `Voz inter (CLP)` = "double(15,2)",
        `Datos (CLP)` = "double(15,2)",
        `Datos nacional (CLP)` = "double(15,2)",
        `Datos inter (CLP)` = "double(15,2)",
        `SMS/MMS (CLP)` = "double(15,2)",
        `SMS/MMS nacional (CLP)` = "double(15,2)",
        `SMS/MMS inter (CLP)` = "double(15,2)",
        `Voz (seg)` = "double(15,2)",
        `Voz nacional (seg)` = "double(15,2)",
        `Voz inter (seg)` = "double(15,2)",
        `Datos (KB)` = "double(15,2)",
        `Datos nacional (KB)` = "double(15,2)",
        `Datos inter (KB)` = "double(15,2)",
        `N. SMS/MMS` = "double(15,2)",
        `Fecha Renovacion` = "date",
        `Fecha` = "varchar(255)"
      ),
      row.names = FALSE,
      overwrite = TRUE,
      append = FALSE,
      allow.keywords = FALSE
    )  
  }
  if (is.null(UAADPT_users2[["MANAGEMENTORG3"]])==TRUE&
      is.null(UAADPT_users2[["MANAGEMENTORG2"]])==FALSE){
    
    dbWriteTable(
      DB,
      "usos",
      UAADPT_users2,
      field.types = list(
        
        `Acceso` = "varchar(255)",
        `Usuario` = "varchar(255)",
        `Centro de facturacion` = "varchar(255)",
        `Modelo` = "varchar(255)",
        `Tipo` = "varchar(255)",
        `Proveedor` = "varchar(255)",
        `Proveedor Nivel 2` = "varchar(255)",
        `Proveedor Nivel 3` = "varchar(255)",
        `MANAGEMENTORG1` = "varchar(255)",
        `MANAGEMENTORG2` = "varchar(255)",
        `Total (CLP)` = "double(15,2)",
        `Plano tarifario (CLP)` = "double(15,2)",
        `Uso (CLP)` = "double(15,2)",
        `Servicios (CLP)` = "double(15,2)",
        `Descuentos (CLP)` = "double(15,2)",
        `Descuento de Plano tarifario (CLP)` = "double(15,2)",
        `Voz (CLP)` = "double(15,2)",
        `Voz nacional (CLP)` = "double(15,2)",
        `Voz inter (CLP)` = "double(15,2)",
        `Datos (CLP)` = "double(15,2)",
        `Datos nacional (CLP)` = "double(15,2)",
        `Datos inter (CLP)` = "double(15,2)",
        `SMS/MMS (CLP)` = "double(15,2)",
        `SMS/MMS nacional (CLP)` = "double(15,2)",
        `SMS/MMS inter (CLP)` = "double(15,2)",
        `Voz (seg)` = "double(15,2)",
        `Voz nacional (seg)` = "double(15,2)",
        `Voz inter (seg)` = "double(15,2)",
        `Datos (KB)` = "double(15,2)",
        `Datos nacional (KB)` = "double(15,2)",
        `Datos inter (KB)` = "double(15,2)",
        `N. SMS/MMS` = "double(15,2)",
        `Fecha Renovacion` = "date",
        `Fecha` = "varchar(255)"
      ),
      row.names = FALSE,
      overwrite = TRUE,
      append = FALSE,
      allow.keywords = FALSE
    )  
  }
  if (is.null(UAADPT_users2[["MANAGEMENTORG4"]])==TRUE&
      is.null(UAADPT_users2[["MANAGEMENTORG3"]])==FALSE){
    dbWriteTable(
      DB,
      "usos",
      UAADPT_users2,
      field.types = list(
        
        `Acceso` = "varchar(255)",
        `Usuario` = "varchar(255)",
        `Centro de facturacion` = "varchar(255)",
        `Modelo` = "varchar(255)",
        `Tipo` = "varchar(255)",
        `Proveedor` = "varchar(255)",
        `Proveedor Nivel 2` = "varchar(255)",
        `Proveedor Nivel 3` = "varchar(255)",
        `MANAGEMENTORG1` = "varchar(255)",
        `MANAGEMENTORG2` = "varchar(255)",
        `MANAGEMENTORG3` = "varchar(255)",
        `Total (CLP)` = "double(15,2)",
        `Plano tarifario (CLP)` = "double(15,2)",
        `Uso (CLP)` = "double(15,2)",
        `Servicios (CLP)` = "double(15,2)",
        `Descuentos (CLP)` = "double(15,2)",
        `Descuento de Plano tarifario (CLP)` = "double(15,2)",
        `Voz (CLP)` = "double(15,2)",
        `Voz nacional (CLP)` = "double(15,2)",
        `Voz inter (CLP)` = "double(15,2)",
        `Datos (CLP)` = "double(15,2)",
        `Datos nacional (CLP)` = "double(15,2)",
        `Datos inter (CLP)` = "double(15,2)",
        `SMS/MMS (CLP)` = "double(15,2)",
        `SMS/MMS nacional (CLP)` = "double(15,2)",
        `SMS/MMS inter (CLP)` = "double(15,2)",
        `Voz (seg)` = "double(15,2)",
        `Voz nacional (seg)` = "double(15,2)",
        `Voz inter (seg)` = "double(15,2)",
        `Datos (KB)` = "double(15,2)",
        `Datos nacional (KB)` = "double(15,2)",
        `Datos inter (KB)` = "double(15,2)",
        `N. SMS/MMS` = "double(15,2)",
        `Fecha Renovacion` = "date",
        `Fecha` = "varchar(255)"
      ),
      row.names = FALSE,
      overwrite = TRUE,
      append = FALSE,
      allow.keywords = FALSE
    )  
  }
  if (is.null(UAADPT_users2[["MANAGEMENTORG5"]])==TRUE&
      is.null(UAADPT_users2[["MANAGEMENTORG4"]])==FALSE){
    dbWriteTable(
        DB,
        "usos",
        UAADPT_users2,
        field.types = list(
          
          `Acceso` = "varchar(255)",
          `Usuario` = "varchar(255)",
          `Centro de facturacion` = "varchar(255)",
          `Modelo` = "varchar(255)",
          `Tipo` = "varchar(255)",
          `Proveedor` = "varchar(255)",
          `Proveedor Nivel 2` = "varchar(255)",
          `Proveedor Nivel 3` = "varchar(255)",
          `MANAGEMENTORG1` = "varchar(255)",
          `MANAGEMENTORG2` = "varchar(255)",
          `MANAGEMENTORG3` = "varchar(255)",
          `MANAGEMENTORG4` = "varchar(255)",
          `Total (CLP)` = "double(15,2)",
          `Plano tarifario (CLP)` = "double(15,2)",
          `Uso (CLP)` = "double(15,2)",
          `Servicios (CLP)` = "double(15,2)",
          `Descuentos (CLP)` = "double(15,2)",
          `Descuento de Plano tarifario (CLP)` = "double(15,2)",
          `Voz (CLP)` = "double(15,2)",
          `Voz nacional (CLP)` = "double(15,2)",
          `Voz inter (CLP)` = "double(15,2)",
          `Datos (CLP)` = "double(15,2)",
          `Datos nacional (CLP)` = "double(15,2)",
          `Datos inter (CLP)` = "double(15,2)",
          `SMS/MMS (CLP)` = "double(15,2)",
          `SMS/MMS nacional (CLP)` = "double(15,2)",
          `SMS/MMS inter (CLP)` = "double(15,2)",
          `Voz (seg)` = "double(15,2)",
          `Voz nacional (seg)` = "double(15,2)",
          `Voz inter (seg)` = "double(15,2)",
          `Datos (KB)` = "double(15,2)",
          `Datos nacional (KB)` = "double(15,2)",
          `Datos inter (KB)` = "double(15,2)",
          `N. SMS/MMS` = "double(15,2)",
          `Fecha Renovacion` = "date",
          `Fecha` = "varchar(255)"
        ),
        row.names = FALSE,
        overwrite = TRUE,
        append = FALSE,
        allow.keywords = FALSE
      )  }
  if (is.null(UAADPT_users2[["MANAGEMENTORG6"]])==TRUE&
      is.null(UAADPT_users2[["MANAGEMENTORG5"]])==FALSE){dbWriteTable(
        DB,
        "usos",
        UAADPT_users2,
        field.types = list(
          
          `Acceso` = "varchar(255)",
          `Usuario` = "varchar(255)",
          `Centro de facturacion` = "varchar(255)",
          `Modelo` = "varchar(255)",
          `Tipo` = "varchar(255)",
          `Proveedor` = "varchar(255)",
          `Proveedor Nivel 2` = "varchar(255)",
          `Proveedor Nivel 3` = "varchar(255)",
          `MANAGEMENTORG1` = "varchar(255)",
          `MANAGEMENTORG2` = "varchar(255)",
          `MANAGEMENTORG3` = "varchar(255)",
          `MANAGEMENTORG4` = "varchar(255)",
          `MANAGEMENTORG5` = "varchar(255)",
          `Total (CLP)` = "double(15,2)",
          `Plano tarifario (CLP)` = "double(15,2)",
          `Uso (CLP)` = "double(15,2)",
          `Servicios (CLP)` = "double(15,2)",
          `Descuentos (CLP)` = "double(15,2)",
          `Descuento de Plano tarifario (CLP)` = "double(15,2)",
          `Voz (CLP)` = "double(15,2)",
          `Voz nacional (CLP)` = "double(15,2)",
          `Voz inter (CLP)` = "double(15,2)",
          `Datos (CLP)` = "double(15,2)",
          `Datos nacional (CLP)` = "double(15,2)",
          `Datos inter (CLP)` = "double(15,2)",
          `SMS/MMS (CLP)` = "double(15,2)",
          `SMS/MMS nacional (CLP)` = "double(15,2)",
          `SMS/MMS inter (CLP)` = "double(15,2)",
          `Voz (seg)` = "double(15,2)",
          `Voz nacional (seg)` = "double(15,2)",
          `Voz inter (seg)` = "double(15,2)",
          `Datos (KB)` = "double(15,2)",
          `Datos nacional (KB)` = "double(15,2)",
          `Datos inter (KB)` = "double(15,2)",
          `N. SMS/MMS` = "double(15,2)",
          `Fecha Renovacion` = "date",
          `Fecha` = "varchar(255)"
        ),
        row.names = FALSE,
        overwrite = TRUE,
        append = FALSE,
        allow.keywords = FALSE
      )  }
  if (is.null(UAADPT_users2[["MANAGEMENTORG7"]])==TRUE&
      is.null(UAADPT_users2[["MANAGEMENTORG6"]])==FALSE){dbWriteTable(
        DB,
        "usos",
        UAADPT_users2,
        field.types = list(
          
          `Acceso` = "varchar(255)",
          `Usuario` = "varchar(255)",
          `Centro de facturacion` = "varchar(255)",
          `Modelo` = "varchar(255)",
          `Tipo` = "varchar(255)",
          `Proveedor` = "varchar(255)",
          `Proveedor Nivel 2` = "varchar(255)",
          `Proveedor Nivel 3` = "varchar(255)",
          `MANAGEMENTORG1` = "varchar(255)",
          `MANAGEMENTORG2` = "varchar(255)",
          `MANAGEMENTORG3` = "varchar(255)",
          `MANAGEMENTORG4` = "varchar(255)",
          `MANAGEMENTORG5` = "varchar(255)",
          `MANAGEMENTORG6` = "varchar(255)",
          `Total (CLP)` = "double(15,2)",
          `Plano tarifario (CLP)` = "double(15,2)",
          `Uso (CLP)` = "double(15,2)",
          `Servicios (CLP)` = "double(15,2)",
          `Descuentos (CLP)` = "double(15,2)",
          `Descuento de Plano tarifario (CLP)` = "double(15,2)",
          `Voz (CLP)` = "double(15,2)",
          `Voz nacional (CLP)` = "double(15,2)",
          `Voz inter (CLP)` = "double(15,2)",
          `Datos (CLP)` = "double(15,2)",
          `Datos nacional (CLP)` = "double(15,2)",
          `Datos inter (CLP)` = "double(15,2)",
          `SMS/MMS (CLP)` = "double(15,2)",
          `SMS/MMS nacional (CLP)` = "double(15,2)",
          `SMS/MMS inter (CLP)` = "double(15,2)",
          `Voz (seg)` = "double(15,2)",
          `Voz nacional (seg)` = "double(15,2)",
          `Voz inter (seg)` = "double(15,2)",
          `Datos (KB)` = "double(15,2)",
          `Datos nacional (KB)` = "double(15,2)",
          `Datos inter (KB)` = "double(15,2)",
          `N. SMS/MMS` = "double(15,2)",
          `Fecha Renovacion` = "date",
          `Fecha` = "varchar(255)"
        ),
        row.names = FALSE,
        overwrite = TRUE,
        append = FALSE,
        allow.keywords = FALSE
      )  }
  if (is.null(UAADPT_users2[["MANAGEMENTORG8"]])==TRUE&
      is.null(UAADPT_users2[["MANAGEMENTORG7"]])==FALSE){dbWriteTable(
        DB,
        "usos",
        UAADPT_users2,
        field.types = list(
          
          `Acceso` = "varchar(255)",
          `Usuario` = "varchar(255)",
          `Centro de facturacion` = "varchar(255)",
          `Modelo` = "varchar(255)",
          `Tipo` = "varchar(255)",
          `Proveedor` = "varchar(255)",
          `Proveedor Nivel 2` = "varchar(255)",
          `Proveedor Nivel 3` = "varchar(255)",
          `MANAGEMENTORG1` = "varchar(255)",
          `MANAGEMENTORG2` = "varchar(255)",
          `MANAGEMENTORG3` = "varchar(255)",
          `MANAGEMENTORG4` = "varchar(255)",
          `MANAGEMENTORG5` = "varchar(255)",
          `MANAGEMENTORG6` = "varchar(255)",
          `MANAGEMENTORG7` = "varchar(255)",
          `Total (CLP)` = "double(15,2)",
          `Plano tarifario (CLP)` = "double(15,2)",
          `Uso (CLP)` = "double(15,2)",
          `Servicios (CLP)` = "double(15,2)",
          `Descuentos (CLP)` = "double(15,2)",
          `Descuento de Plano tarifario (CLP)` = "double(15,2)",
          `Voz (CLP)` = "double(15,2)",
          `Voz nacional (CLP)` = "double(15,2)",
          `Voz inter (CLP)` = "double(15,2)",
          `Datos (CLP)` = "double(15,2)",
          `Datos nacional (CLP)` = "double(15,2)",
          `Datos inter (CLP)` = "double(15,2)",
          `SMS/MMS (CLP)` = "double(15,2)",
          `SMS/MMS nacional (CLP)` = "double(15,2)",
          `SMS/MMS inter (CLP)` = "double(15,2)",
          `Voz (seg)` = "double(15,2)",
          `Voz nacional (seg)` = "double(15,2)",
          `Voz inter (seg)` = "double(15,2)",
          `Datos (KB)` = "double(15,2)",
          `Datos nacional (KB)` = "double(15,2)",
          `Datos inter (KB)` = "double(15,2)",
          `N. SMS/MMS` = "double(15,2)",
          `Fecha Renovacion` = "date",
          `Fecha` = "varchar(255)"
        ),
        row.names = FALSE,
        overwrite = TRUE,
        append = FALSE,
        allow.keywords = FALSE
      )  }
  if (is.null(UAADPT_users2[["MANAGEMENTORG8"]])==FALSE){dbWriteTable(
    DB,
    "usos",
    UAADPT_users2,
    field.types = list(
      
      `Acceso` = "varchar(255)",
      `Usuario` = "varchar(255)",
      `Centro de facturacion` = "varchar(255)",
      `Modelo` = "varchar(255)",
      `Tipo` = "varchar(255)",
      `Proveedor` = "varchar(255)",
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
      `Total (CLP)` = "double(15,2)",
      `Plano tarifario (CLP)` = "double(15,2)",
      `Uso (CLP)` = "double(15,2)",
      `Servicios (CLP)` = "double(15,2)",
      `Descuentos (CLP)` = "double(15,2)",
      `Descuento de Plano tarifario (CLP)` = "double(15,2)",
      `Voz (CLP)` = "double(15,2)",
      `Voz nacional (CLP)` = "double(15,2)",
      `Voz inter (CLP)` = "double(15,2)",
      `Datos (CLP)` = "double(15,2)",
      `Datos nacional (CLP)` = "double(15,2)",
      `Datos inter (CLP)` = "double(15,2)",
      `SMS/MMS (CLP)` = "double(15,2)",
      `SMS/MMS nacional (CLP)` = "double(15,2)",
      `SMS/MMS inter (CLP)` = "double(15,2)",
      `Voz (seg)` = "double(15,2)",
      `Voz nacional (seg)` = "double(15,2)",
      `Voz inter (seg)` = "double(15,2)",
      `Datos (KB)` = "double(15,2)",
      `Datos nacional (KB)` = "double(15,2)",
      `Datos inter (KB)` = "double(15,2)",
      `N. SMS/MMS` = "double(15,2)",
      `Fecha Renovacion` = "date",
      `Fecha` = "varchar(255)"
    ),
    row.names = FALSE,
    overwrite = TRUE,
    append = FALSE,
    allow.keywords = FALSE
  )  }
  usofinal<<-UAADPT_users2
  print("DATOS SUBIDOS")
  
  ######Tabla UTP_accesses######
 
}
