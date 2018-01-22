{
  SFOpciones<-subset(SFACTURADOS,
                     SFACTURADOS[["Tipo de producto"]]=="Option")
  SFOpciones<<-SFOpciones
  SFPlanesA<-subset(SFACTURADOS,
                    SFACTURADOS[["Tipo de producto"]]=="Plano tarifario" & 
                      SFACTURADOS[["Estado acceso"]]=="Activo")
  SFPlanesDb<-subset(SFACTURADOS,
                     SFACTURADOS[["Tipo de producto"]]=="Plano tarifario" & 
                       SFACTURADOS[["Estado acceso"]]=="Dado de baja")
  SFPlanes<-rbind(SFPlanesDb,SFPlanesA)
  SFPlanes<<-SFPlanes
  SFPlanes2<-SFPlanes

  ######################Ajuste de Servicios Facturados################
  a<-duplicated(SFPlanes2[["Acceso"]],fromLast = FALSE)
  b<-duplicated(SFPlanes2[["Acceso"]],fromLast = TRUE)
  SFPlanes2[["Duplicados"]]<-a
  SFPlanes2[["Duplicados2"]]<-b
  SFduplicados<-subset(SFPlanes2,SFPlanes2[["Duplicados"]]=="TRUE"|
                         SFPlanes2[["Duplicados2"]]=="TRUE")
  SF_no_duplicados<-subset(SFPlanes2,SFPlanes2[["Duplicados"]]=="FALSE"&
                             SFPlanes2[["Duplicados2"]]=="FALSE")
  if(length(SFduplicados[["Acceso"]])>0){
    SF_a_evaluar<- merge(SFduplicados,
                         MOVISTAR_PLANES,
                         by = "Producto",
                         all.x = TRUE)
    SF_fueradecontratoSC<-subset(SF_a_evaluar,
                                 is.na(SF_a_evaluar[["Tipo"]])==TRUE &
                                   SF_a_evaluar[["Importe de las opciones descontadas (CLP)"]]==0)
    SF_fueradecontratoCC<-subset(SF_a_evaluar,
                                 is.na(SF_a_evaluar[["Tipo"]])==TRUE &
                                   SF_a_evaluar[["Importe de las opciones descontadas (CLP)"]]!=0)
    SF_en_contrato<-subset(SF_a_evaluar,
                           is.na(SF_a_evaluar[["Tipo"]])==FALSE)
    SF_CPduplicados<-rbind(SF_en_contrato,SF_fueradecontratoCC)
    SF_CPduplicados[["Duplicado"]]<-NULL
    SF_CPduplicados[["Duplicado2"]]<-NULL
    a<-duplicated(SF_CPduplicados[["Acceso"]],fromLast = FALSE)
    b<-duplicated(SF_CPduplicados[["Acceso"]],fromLast = TRUE)
    SF_CPduplicados[["Duplicados"]]<-a
    SF_CPduplicados[["Duplicados2"]]<-b
    SFduplicados2<-subset(SF_CPduplicados,SF_CPduplicados[["Duplicados"]]=="TRUE"|SF_CPduplicados[["Duplicados2"]]=="TRUE")
    SFduplicadosbuenos<-subset(SF_CPduplicados,SF_CPduplicados[["Duplicados"]]=="FALSE"&SF_CPduplicados[["Duplicados2"]]=="FALSE")
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
        CdF[i]<-SFUnico[["Centro de facturacion"]][1]
        IOF[i]<-sum(SFUnico[["Importe de las opciones facturadas (CLP)"]][1:length(SFUnico[["Acceso"]])])
        IDPT[i]<-sum(SFUnico[["Importe descuentos sobre plano tarifario (CLP)"]][1:length(SFUnico[["Acceso"]])])
        IOD[i]<-sum(SFUnico[["Importe de las opciones descontadas (CLP)"]][1:length(SFUnico[["Acceso"]])])
        Accf[i]<-SFUnico[["Acceso fix"]][1]
        
      }
      Accesos<-(unique(SFduplicados2["Acceso"]))
      Accesos["Estado acceso"]<-EstAcc
      Accesos["Producto"]<-Prod
      Accesos["Tipo de producto"]<-TdProd
      Accesos["Centro de facturacion"]<-CdF
      Accesos["Importe de las opciones facturadas (CLP)"]<-IOF
      Accesos["Importe descuentos sobre plano tarifario (CLP)"]<-IDPT
      Accesos["Importe de las opciones descontadas (CLP)"]<-IOD
      Accesos["Acceso fix"]<-Accf
      
      SFUnicos<-Accesos
      
    }
    else {
      SFUnicos<-subset(SFACTURADOS,SFACTURADOS[["Acceso"]]==1&SFACTURADOS[["Acceso"]]!=1)
      
    }
    SF_no_duplicados[["Duplicados"]]<-NULL
    SF_no_duplicados[["Duplicados2"]]<-NULL
    SFduplicadosbuenos2<-subset(SFduplicadosbuenos,select = c("Acceso","Estado acceso","Producto","Tipo de producto","Centro de facturacion","Importe de las opciones facturadas (CLP)",
                                                              "Importe descuentos sobre plano tarifario (CLP)","Importe de las opciones descontadas (CLP)","Acceso fix"))
    SF_Final<-rbind(SFUnicos,SF_no_duplicados,SFduplicadosbuenos2)
    SFduplicados2<-subset(SFduplicados2,select = c("Acceso","Estado acceso","Producto","Tipo de producto","Centro de facturacion","Importe de las opciones facturadas (CLP)",
                                                   "Importe descuentos sobre plano tarifario (CLP)","Importe de las opciones descontadas (CLP)","Acceso fix"))
    SF_fueradecontratoSC<-subset(SF_fueradecontratoSC,select = c("Acceso","Estado acceso","Producto","Tipo de producto","Centro de facturacion","Importe de las opciones facturadas (CLP)",
                                                                 "Importe descuentos sobre plano tarifario (CLP)","Importe de las opciones descontadas (CLP)","Acceso fix"))
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
  rm(SF_a_evaluar,SF_CPduplicados,SF_en_contrato,SF_fueradecontratoCC,SF_no_duplicados,SFduplicados,SFduplicados2,SFduplicadosbuenos,SFduplicadosbuenos2,SFPlanes2,SF_fueradecontratoSC,SFUnicos,SFPlanesA,SFPlanesDb)
  SFPlanes_final<-SF_Final
  if(length(SFOpciones[["Acceso"]])>0){
    SF_Final[["Duplicados"]]<-NULL
    SF_Final[["Duplicados2"]]<-NULL
    SF_Final<-rbind(SF_Final,SFOpciones)
  }
  SF_Final<<-SF_Final
  SF_Apartados<<-SF_Apartados
  SFPlanes_final<<-SFPlanes_final
}