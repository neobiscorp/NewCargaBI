{
  
  #Se borran las columnas innecesarias que causaran distorcion sobre los nombres finales de las columnas
ACCESSES2<-ACCESSES
ACCESSES2[["Acceso"]]<-NULL
ACCESSES2[["Proveedor"]]<-NULL
ACCESSES2[["Tipo"]]<-NULL
ACCESSES2[["Estado"]]<-NULL
# se une uso con ACCESSES 
Consolidado <- merge(uso,
                      ACCESSES2,
                      by.x = "Acceso fix",
                      by.y = "Acceso fix",
                      all.x = TRUE)

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
 
  ASDAS<<-names(SinUsos)
  SinUsos[,'usocant']<-SinUsos[,'Voz (sec)']+SinUsos[,'Datos (KB)']+SinUsos[,'N.° SMS/MMS']
  SinUsos <-
    subset(
      SinUsos,
      select = c(
        "Acceso",
        "Fecha",
        "Total",
        "usocant"
      )
    )
  
  SinUsos<-subset(SinUsos,SinUsos[["usocant"]]==0)

  month1 <- sapply(SinUsos[,'Fecha'], substr, 6, 7)
    month <- as.numeric(month1)
    rm(month1)
    year1<-sapply(SinUsos[,'Fecha'],substr,1,4)
    year<-as.numeric(year1)
    rm(year1)
    AAA<-data.frame(SinUsos[,'Fecha'])
    AAA[,'month']<-month
    AAA[,'year']<-year
    BBB<-subset(AAA,
                AAA["year"]==max(year))
    rm(month,year)
    fin1<-max(BBB[["month"]])
    fin2<-max(BBB[["year"]])
    rm(BBB)
    
    
  
  for (i in 1:length(SinUsos[["Acceso"]])){
  SinUsos[["Meses"]][i]<-(fin2-AAA[["year"]][i])*12+fin1-AAA[["month"]][i]+1
  }
    rm(AAA)
  SinUsos<-subset(SinUsos,SinUsos[["Meses"]]<=3) 
  SinUsos<<-subset(SinUsos,SinUsos[["Meses"]]<=3)

  ########Excepciones############
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

  Consolidado<<-Consolidado
}

