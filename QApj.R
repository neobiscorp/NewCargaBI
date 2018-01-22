{
  #Preparacion de datos
  {
    Meses<-function(x){ 
      month1 <- sapply(x, substr, 6, 7)
      month <- as.numeric(month1)
      rm(month1)
      year1<-sapply(x,substr,1,4)
      year<-as.numeric(year1)
      rm(year1)
      AAA<-data.frame(x)
      AAA[,'month']<-month
      AAA[,'year']<-year
      BBB<-subset(AAA,
                  AAA["year"]==min(year))
      rm(month,year)
      inicio1<-min(BBB[["month"]])
      inicio2<-min(BBB[["year"]])
      rm(BBB)
      return((AAA[[,"year"]]-inicio2)*12+AAA[[,"month"]]-inicio1+1)
      rm(AAA)
    }#Meses(i,x) donde x es una columna de Fechas ordenadas por YYYY/MM/DD o YYYY-MM-DD HH:mm:ss e "i" es una fila en especifica de la columna, el resultado obtenido es la cantidad de meses transcurridos desde el mes inicial hasta el mes del dato "i"
    #######Preparacion de datos del cdr que estan en server.R########
    # cdr_accesses <-
    #   merge (cdr, ACCESSES, by.x = "Numero de llamada fix", by.y = "Acceso fix", all.x = TRUE)
    # 
    # cdr_accesses <-
    #   merge (
    #     cdr_accesses,
    #     ACCESSES,
    #     by.x = "Numero llamado",
    #     by.y = "Acceso fix",
    #     all.x = TRUE
    #   )
    # cdr_accesses[, "NET"] <-
    #   ifelse (cdr_accesses["Proveedor Nivel 2.x"] == cdr_accesses["Proveedor Nivel 2.y"] &
    #             cdr_accesses["Proveedor Nivel 3.x"] == cdr_accesses["Proveedor Nivel 3.y"],
    #           1,
    #           0) #Se crea un mecanismo para definir cuales llamadas son realizadas dentro de un mismo cliente
    # 
    # mes1 <- sapply(cdr_accesses["Fecha de llamada"], substr, 6, 7)
    # mes <- as.numeric(mes1)
    # cdr_accesses["Mes"] <- mes #Se agrega el mes como columna a parte de la fecha para poder utilizar el mes dentro de otras formulas
    # 
    # rm(mes1, mes) #Con rm se borran los exedentes de datos, genera una optimizacion de tiempo a la hora de procesar los datos
    ################################## 
    # cdr_accesses <-
    #   subset(cdr_accesses, cdr_accesses["Mes"] != min(cdr_accesses["Mes"])) # Se extraen los datos sobrantes de los periodos anteriores de los que se van a revisar
    
    cdr_accesses[,"PdivD"] <- cdr_accesses[,"Precio"]/cdr_accesses[,"Duracion"]*60 #generacion de variable Precio/Duracion que se usara mas adelante *60 para que este en minutos
    cdr_accesses[,"PdivV"] <- cdr_accesses[,"Precio"]/cdr_accesses[,"Volumen"]*1024 #generacion de variable Precio/Volumen que se usara mas adelante *1024 para que este en Mb
    MesesT<-c()
    
      MesesT<-Meses(cdr_accesses[["Fecha de llamada"]])
    
    cdr_accesses[,"MesesT"]<-MesesT
    cdr_accesses["Mes"]<-NULL
    #Estas variables pueden presentar valores infinitos o no acordes a la realidad pero no pasan a estar incluidos en algun dato relevante debido a que los filtros usados
    #tienen la precision necesaria para evitar tenerlos dentro de alguna formula
    #CREACION usos con accesos unicos
    
    
    uso2 <- uso
    a<-duplicated(uso2[["Acceso"]],fromLast = TRUE )
    uso2[["duplicado"]] <-a
    uso2<-subset(uso2,uso2[["duplicado"]] == FALSE)
    PLAN2 <-subset(PLAN,PLAN["Tipo de producto"]=="Plano tarifario")
    usounicojoinPLANt <-
      merge(
        uso2,
        PLAN2,
        by.x = "Acceso fix",
        by.y = "Acceso fix",
        all.x = TRUE
      )
    usoplantjointipo <-
      merge(
        usounicojoinPLANt,
        TIPO,
        by.x = "Producto",
        by.y = "Producto",
        all.x = TRUE
      )
    usoplantjointipo2<-usoplantjointipo
    usoplantjointipo2["Proveedor.x"] <- NULL
    usoplantjointipo2["Proveedor.y"] <- NULL
    
    UTP_accesses<-
      merge (
        usoplantjointipo2,
        ACCESSES,
        by.x = "Acceso fix",
        by.y = "Acceso fix",
        all.x = TRUE
      ) #Conjunto de uniones de tablas que permitiran definir la cantidad de dispositivos de Banda Ancha Movil (BAM) entre los dispositivos
    #UTP_accesses es la union de las tablas uso, TIPO, PLAN, ACCESSES, dejando como id el Accesos unicos de la tabla usos
    
    rm(uso2,usounicojoinPLANt,a,PLAN2,usoplantjointipo2)
    #Datos a sacar se dejan en un formato rellenable para que no haya limitantes en la cantidad de centros de factura que presentara la empresa
    #estos datos a sacar presentan la explicacion de que son dentro de la Documentacion
    nM<-c() #numero de meses que usa el centro de facturacion i de movistar, donde si nM tiene menos meses el promedio mensual se seguira calculando sin problemas
    nE<-c() #lo mismo de nM pero para entel
    datEn<-c()#cantidad de datos segun centro de facturacion para proveedor entel
    datM<-c()# lo mismo pero para movistar
    EntATodDes<-c()
    EntBAM<-c()
    EntIntVoz<-c()
    EntMbAct<-c()
    EntMinAct<-c()
    EntMms<-c()
    EntRoaDat<-c()
    EntRoaSms<-c()
    EntRoaVoz<-c()
    EntSms<-c()
    EntTotMin<-c()
    EntVozOnNet<-c()
    
    MovATodDes<-c()
    MovBAM<-c()
    #MovIntVoz<-c() solo esta en entel
    MovMbAct<-c()
    MovMinAct<-c()
    MovMms<-c()
    MovRoaDat<-c()
    MovRoaSms<-c()
    MovRoaVoz<-c()
    MovSms<-c()
    MovTotMin<-c()
    MovVozOnNet<-c()
    
    
    nomemp<-unique(CUENTAS["Empresa"]) #Se genera una columna que solo contenga los nombres sin repetir de los centros de facturacion
    m<-lapply(nomemp,as.character) #se deja en un formato compatible con otras formulas
    CC<-as.numeric(lengths(nomemp))  #se cuentan la cantidad de nombres de empresas que hay de los centros de facturacion
    
    
    
    
    UTP_accesses <-
      merge (
        UTP_accesses,
        CUENTAS,
        by.x = "Proveedor Nivel 3",
        by.y = "Cuenta Cliente",
        all.x = TRUE
      )
    UTP_accesses2<-UTP_accesses #se deja guardado un original de UTP_accesses para no perder datos al filtrar
  }
  #BAM por Empresa
  for(i in 0:CC+1){
    UTP_accesses<-UTP_accesses2 #se recuperan los datos filtrados mas adelante
    if (i==CC+1){ #en el ultimo loop del for se usan los datos generales
      print("datos generales")
    }else{    
      print(i)
      print(m[["Empresa"]][i]) #Se ve que empresa se esta usando
      UTP_accesses<-
        subset(UTP_accesses,
               UTP_accesses["Empresa"]==m[["Empresa"]][i] 
        ) #Se filtra segun que empresa para sacar RFP de los distintos clientes
    }
    #BAM O SERVICIOS DE TELEMETRÍA MOVISTAR
    {MovBAMm <- subset(UTP_accesses,UTP_accesses[["Tipo"]] == "BAM"
                       &UTP_accesses[["Proveedor.x"]]=="Movistar CL"
    )
      MovBAM[i] <-as.numeric(length(MovBAMm[["Acceso fix"]])) #La forma de saber la cantidad de BAM es dejar los accesos como unicos, asignarles un Tipo segun el nombre del producto, luego generar los filtros de Tipo = BAM y por proveedor
      rm(MovBAMm)}
    #BAM O SERVICIOS DE TELEMETRÍA ENTEL
    EntBAMm <- subset(UTP_accesses,UTP_accesses[["Tipo"]] == "BAM"
                      &UTP_accesses[["Proveedor.x"]]=="Entel PCS (CL)")
    EntBAM[i] <-as.numeric(length(EntBAMm[["Acceso fix"]]))# se usa lo mismo que en MovBAM pero con proveedor = Entel
    rm(EntBAMm)
  }
  print(MovBAM)
  print(EntBAM)
  
  #Eleccion de Centro de Facturacion
  {
    cdr_accesses <-
      merge (
        cdr_accesses,
        CUENTAS,
        by.x = "Proveedor Nivel 3.x",
        by.y = "Cuenta Cliente",
        all.x = TRUE
      ) #Se hace une cdr_accesses con CUENTAS para poder filtrar segun el nombre de la empresa
    cdr_accesses2<-cdr_accesses #se guardan los datos de cdr_accesses para ser reutilizados despues
    nomemp<-unique(CUENTAS["Empresa"])
    
    m<-lapply(nomemp,as.character)
    
    CC<-as.numeric(lengths(nomemp))
    
    
  }
  #DATOS A SACAR
  for(i in 0:CC+1){ #se hace lo mismo que en BAM pero esta vez se filtra el cdr_accesses segun el nombre de la empresa
    cdr_accesses<-cdr_accesses2
    if (i==CC+1){
      print("datos generales")
    }else{
      print(i)
      print(m[["Empresa"]][i])
      cdr_accesses<-
        subset(cdr_accesses,
               cdr_accesses["Empresa"]==m[["Empresa"]][i]
        )
    }
    
    #######MOVISTAR######
    
    {
      cdr_movistar <-
        subset(cdr_accesses, cdr_accesses[["Proveedor.x"]] == "Movistar CL") #dado que en esta parte solo usaremos proveedor movistar, lo filtraremos ahora para no hacerlo por cada dato a sacar
      datM[i]<-as.numeric(lengths(cdr_movistar["Tipo de llamada"])) #guarda la cantidad de datos existentes como un numero
      if(datM[[i]]!=0){ #Optimizacion de procesamiento, si no hay datos disponibles de movistar los resultados obtenidos seran NA en todo otro caso se obtendran resultados
        #CONSUMO TOTAL VOZ
        print("Datos Movistar")
        print(datM[[i]])
        movistarvoz <-
          subset(cdr_movistar,
                 cdr_movistar["Tipo de llamada"] == "Voz" &
                   cdr_movistar["Duracion"] > 0)  #Se guarda el cdr filtrado por tipo de llamada = voz como movistar voz
        nM[i]<-as.numeric(length(unique(cdr_movistar[["Mes"]])))
        MovTotMin[i] <- as.numeric((sum(movistarvoz["Duracion"]) / 60) / nM[[i]]) #Manera de obtener El consumo total Voz, debido a que la duracion esta en segundos hay que dividir por 60 y por los meses para obtener un promedio mensual
        print("MovTotMin")
        print(MovTotMin[[i]])
        #CONSUMO VOZ ENTRE USUARIOS SAAM SA
        movistarvozonnet <-
          subset(movistarvoz,
                 movistarvoz["NET"] == 1&
                   (movistarvoz["Geografia"] == "Local"|
                      movistarvoz["Geografia"]=="Nacional desconocido")
          )
        datM2<-as.numeric(lengths(movistarvozonnet["Tipo de llamada"]))
        if(datM2!=0){# Los datos en este punto por los filtros pueden llegar a faltar y producir errores por lo que se repite la mecanica de evitar que se calcule si no hay datos que calcular
          MovVozOnNet[i] <- as.numeric((sum(movistarvozonnet["Duracion"]) / 60) / nM[[i]])
        }else{MovVozOnNet[i]<-0}
        rm(movistarvozonnet)
        print("MovVozOnNet")
        print(MovVozOnNet[[i]])
        
        #CONSUMO VOZ A TODO DESTINO
        MovATodDes[i] <- MovTotMin[[i]] - MovVozOnNet[[i]] # a todo destino es a los que no estan dentro de la misma red, por lo que se resta el On Net de los minutos totales
        print("MovATodDes")
        print(MovATodDes[[i]])
        
        
        #SMARTPHONES GAMA ALTA
        #SMARTPHONES GAMA MEDIA
        #---
        
        #MENSAJERÍA SMS
        movistarSMS <-
          subset(
            cdr_movistar,
            cdr_movistar["Tipo de llamada"] == "SMS"
            &
              cdr_movistar["Precio"] > 0 &
              (
                cdr_movistar["Geografia"] == "Local" |
                  cdr_movistar["Geografia"] == "Nacional desconocido"
              )
          )
        MovSms[i]<-as.numeric(sapply(movistarSMS["Precio"], median)) #El precio se calcula como la mediana de los precios de los mensajes debido a que los valores irregulares pueden ser bastantes 
        print("MovSms") #Los valores irregulares son en su mayoria son de mensajes que no deberian mandarse desde un celular de la empresa, es decir, promociones, cupones, o con costo de compra de algo
        print(MovSms)
        rm (movistarSMS)
        
        #MENSAJERÍA MMS
        movistarMMS <-
          subset(
            cdr_movistar,
            cdr_movistar["Tipo de llamada"] == "MMS"
            &
              cdr_movistar["Precio"] > 0 &
              (
                cdr_movistar["Geografia"] == "Local" |
                  cdr_movistar["Geografia"] == "Nacional desconocido"
              )
          )
        
        MovMms[i]<-as.numeric(sapply(movistarMMS["Precio"], median)) #bajo la misma logica que en SMS se usa mediana para los MMS
        print(MovMms)
        rm (movistarMMS)
        
        #USUARIOS ROAMING ON DEMAND
        #ROAMING VOZ
        mroam <-
          subset(
            cdr_movistar,
            cdr_movistar["Geografia"] == "Roaming saliente" |
              cdr_movistar["Geografia"] == "Roaming entrante"
          ) #Filtro para revisar solo los Roaming
        mroamvoz <-
          subset(mroam,
                 mroam["Tipo de llamada"] == "Voz" &
                   mroam["Duracion"] > 0)
        datM2<-as.numeric(lengths(mroamvoz["Tipo de llamada"])) 
        if(datM2!=0){     #misma logica pasada se ocupa despues de un filtro fuerte
          MovRoaVoz[i]<-(sum(mroamvoz["Duracion"]) / 60 / nM[[i]]) # los segundos se dividen por 60 y se obtiene un promedio mensual
          print(MovRoaVoz)
        }
        rm(mroamvoz)
        #ROAMING DATOS
        mroamdat <- subset(mroam,
                           mroam["Tipo de llamada"] == "Datos" &
                             mroam["Volumen"] > 0)
        datM2<-as.numeric(lengths(mroamdat["Tipo de llamada"]))
        if(datM2!=0){
          MovRoaDat[i]<-(sum(mroamdat["Volumen"]) / 1024 / nM[[i]]) #los bytes se dividen por 1024 para obtener la cantidad en Mb y se obtiene el promedio mensual
          print(MovRoaDat)
        }
        rm(mroamdat)
        #ROAMING MENSAJES
        mroamsms <- subset(mroam,
                           mroam["Tipo de llamada"] == "SMS" &
                             mroam["Precio"] > 0)
        
        MovRoaSms[i]<-as.numeric(sapply(mroamsms["Precio"], median))
        print(MovRoaSms)
        rm(mroamsms)
        
        #---
        #$/minuto actual
        movistarvoz <-
          subset(
            movistarvoz,
            movistarvoz["Precio"] > 0  &
              (
                movistarvoz["Geografia"] == "Nacional desconocido" |
                  movistarvoz["Geografia"] == "Local"
              )
          )
        
        MovMinAct[i]<-as.numeric(sapply(movistarvoz["PdivD"],mean)) #cuando se pregunta el precio se usan los que tienen precio mayor a 0 para diferenciar de lo que esta dentro de los planes
        print(MovMinAct)
        
        #$/Mb Actual
        movistardatos <- subset(cdr_movistar,
                                cdr_movistar["Tipo de llamada"] == "Datos"
                                & (cdr_movistar["Geografia"] == "Nacional Desconocido" |
                                     cdr_movistar["Geografia"] == "Local")
                                & cdr_movistar["Volumen"] > 0
                                & cdr_movistar["Precio"] > 0) #Filtros necesarios para sacar el precio promedio por Mb de datos
        
        #movistardatos[, "Volumen"] <- movistardatos["Volumen"] / 1024
        MovMbAct[i] <- as.numeric(sapply(movistardatos["PdivV"],mean)) #Uso de promedio para obtener el precio por Mb de movistar
        
        print("MovMbAct")
        print(MovMbAct)
        #Remoción tablas y variables
        rm(cdr_movistar, mroam, movistarvoz,movistardatos) #Se borran los datos que son innecesarios mantener
      }
    }
    
    #######ENTEL######
    #Se usan los mismos fitros y formulas que en movistar a excepcion de "A internacional" que no se presenta en movistar y se presenta en la linea 444 - 450
    {
      cdr_entel <-
        subset(cdr_accesses, cdr_accesses[["Proveedor.x"]] == "Entel PCS (CL)")
      datEn[i]<-as.numeric(lengths(cdr_entel["Tipo de llamada"]))
      if(datEn[[i]]!=0){
        #CONSUMO TOTAL VOZ
        print("Datos entel")
        print(datEn[[i]])
        entelvoz <-
          subset(cdr_entel,
                 cdr_entel["Tipo de llamada"] == "Voz" &
                   cdr_entel["Duracion"] > 0)
        nE[i]<-as.numeric(length(unique(cdr_entel[["Mes"]])))
        EntTotMin[i] <- (sum(entelvoz["Duracion"]) / 60) / nE[[i]]
        print("EntTotMin")
        print(EntTotMin)
        
        #CONSUMO VOZ ENTRE USUARIOS SAAM SA
        entelvozonnet <-
          subset(entelvoz,
                 entelvoz["NET"] == 1&
                   (entelvoz["Geografia"] == "Local" |
                      entelvoz["Geografia"] == "Nacional desconocido"))
        
        EntVozOnNet[i] <- (sum(entelvozonnet["Duracion"]) / 60) / nE[[i]]
        rm(entelvozonnet)
        print(EntVozOnNet)
        
        #CONSUMO VOZ A TODO DESTINO
        EntATodDes[i] <- EntTotMin[[i]] - EntVozOnNet[[i]]
        print(EntATodDes)
        #Paises mas llamados
        
        #SMARTPHONES GAMA ALTA
        #SMARTPHONES GAMA MEDIA
        #---
        
        #MENSAJERÍA SMS
        entelSMS <-
          subset(
            cdr_entel,
            cdr_entel["Tipo de llamada"] == "SMS"
            &
              cdr_entel["Precio"] > 0 &
              (cdr_entel["Geografia"] == "Local" |
                 cdr_entel["Geografia"] == "Nacional desconocido")
          )
        EntSms[i] <- as.numeric(sapply(entelSMS["Precio"], median))
        print("EntSms")
        print(EntSms)
        rm (entelSMS)
        
        #MENSAJERÍA MMS
        entelMMS <-
          subset(
            cdr_entel,
            cdr_entel["Tipo de llamada"] == "MMS"
            &
              cdr_entel["Precio"] > 0 &
              (cdr_entel["Geografia"] == "Local" |
                 cdr_entel["Geografia"] == "Nacional desconocido")
          )
        
        EntMms[i] <- as.numeric(sapply(entelMMS["Precio"], median))
        print(EntMms)
        rm (entelMMS)
        
        #USUARIOS ROAMING ON DEMAND
        #ROAMING VOZ
        mroam <-
          subset(cdr_entel,
                 cdr_entel["Geografia"] == "Roaming saliente" |
                   cdr_entel["Geografia"] == "Roaming entrante")
        mroamvoz <-
          subset(mroam,
                 mroam["Tipo de llamada"] == "Voz" &
                   mroam["Duracion"] > 0)
        
        EntRoaVoz[i] <- (sum(mroamvoz["Duracion"]) / 60 / nE[[i]])
        print(EntRoaVoz)
        rm(mroamvoz)
        #ROAMING DATOS
        mroamdat <- subset(mroam,
                           mroam["Tipo de llamada"] == "Datos" &
                             mroam["Volumen"] > 0)
        datM2<-as.numeric(lengths(mroamdat["Tipo de llamada"]))
        if(datM2!=0){
          EntRoaDat[i] <- (sum(mroamdat["Volumen"]) / 1024 / nE[[i]])
          print(EntRoaDat)
        }
        rm(mroamdat)
        #ROAMING MENSAJES
        mroamsms <- subset(mroam,
                           mroam["Tipo de llamada"] == "SMS" &
                             mroam["Precio"] > 0)
        datM2<-as.numeric(lengths(mroamsms["Tipo de llamada"]))
        if(datM2!=0){
          EntRoaSms[i] <- as.numeric(sapply(mroamsms["Precio"], median))
          print(EntRoaSms)
        }
        rm(mroamsms)
        
        #Internacional Voz
        entelvozint <-
          subset(entelvoz, entelvoz["Geografia"] == "A internacional")
        entelvozint["Duracion"]<-entelvozint["Duracion"]/60
        EntIntVoz[i] <- as.numeric(sapply(entelvozint["Duracion"],sum))/nE[[i]] #se genera desde el filtro de Voz-Entel y Geografia "A internacional" un promedio de duracion mensual
        print (EntIntVoz)
        
        #---
        #$/minuto actual
        entelvoz <-
          subset(entelvoz,
                 entelvoz["Precio"] > 0  &
                   #entelvoz["Mes"] == max(entelvoz["Mes"]) &
                   (entelvoz["Geografia"] == "Nacional desconocido" |
                      entelvoz["Geografia"] == "Local"))
        EntMinAct[i]<-as.numeric(sapply(entelvoz["PdivD"],mean))
        #EntMinAct <-
        # (sum(entelvoz["Precio"]) / (sum(entelvoz["Duracion"]) / 60))
        print("EntMinAct")
        print(EntMinAct)
        
        #$/Mb Actual
        enteldatos <- subset(cdr_entel,
                             cdr_entel["Tipo de llamada"] == "Datos"
                             & (cdr_entel["Geografia"] == "Nacional Desconocido" |
                                  cdr_entel["Geografia"] == "Local")
                             & cdr_entel["Volumen"] > 0
                             & cdr_entel["Precio"] > 0)
        
        
        EntMbAct[i] <- as.numeric(sapply(enteldatos["PdivV"],mean))
        
        print(EntMbAct)
        #Remoción tablas y variables
        rm(cdr_entel, mroam, entelvoz, enteldatos,entelvozint)
      }
      print(i)
    }
  }
  cdr_accesses<-cdr_accesses2
  rm(cdr_accesses2)
  #Paises mas llamados
  {
    cdr_movistar <-
      subset(cdr_accesses, cdr_accesses[["Proveedor.x"]] == "Movistar CL")
    movistarpais<-
      subset(cdr_movistar,
             cdr_movistar["Tipo de llamada"]== "Voz"&
               (cdr_movistar["Geografia"] == "A internacional"|
                  cdr_movistar["Geografia"]=="Roaming saliente")&
               cdr_movistar["Duracion"]>0&
               cdr_movistar["Pais destinatario"]!="Chile") #Filtro para conocer las llamadas fuera de chile de parte de movistar
    PaisesMov<-as.character(unique(movistarpais[["Pais destinatario"]])) #se crea una variable que contenga los paises a los que se llamaron sin repetirse
    PaisesMov2<-c()
    PaisesMov3<-c()
    for(j in 1:length(PaisesMov)){ #Segun la cantidad de paises llamados hay que ver cuantos minutos fueron destinados a ese pais
      
      MovPaisDur<-subset(movistarpais,movistarpais["Pais destinatario"]==PaisesMov[[j]]) #se filtra segun el pais que queremos ver
      
      PaisesMov2[j]<-sum(MovPaisDur[["Duracion"]]/60) # se le asigna una duracion
      PaisesMov3[j]<-length(MovPaisDur[["Geografia"]]) # se ve la cantidad de llamadas que se hicieron
    }
    
    MovPais<-data.frame(PaisesMov) #se genera un formato para juntar los datos obtenidos
    
    MovPais["Duracion"]<-PaisesMov2 #se agrega la duracion al pais correspondiente
    MovPais["Cantidad de llamadas"]<-PaisesMov3 #se agrega la cantidad de llamadas al pais correspondiente
    MovPais3<-MovPais #generamos la lista de paises desordenada
    MovPais<-MovPais3[order(-MovPais3[["Duracion"]]),] #Manera de dejar ordenado de manera decendiente
    
    rm(PaisesMov,PaisesMov2,PaisesMov3,MovPaisDur,movistarpais,MovPais3)
    #Se hace de la misma manera la lista de paises ahora con el proveedor entel
    cdr_entel <-
      subset(cdr_accesses, cdr_accesses[["Proveedor.x"]] == "Entel PCS (CL)")
    entelpais<-
      subset(cdr_entel,
             cdr_entel["Tipo de llamada"]== "Voz"&
               (cdr_entel["Geografia"] == "A internacional"|
                  cdr_entel["Geografia"]=="Roaming saliente")&
               cdr_entel["Duracion"]>0&
               cdr_entel["Pais destinatario"]!="Chile")
    PaisesEnt<-as.character(unique(entelpais[["Pais destinatario"]]))
    PaisesEnt2<-c()
    PaisesEnt3<-c()
    for(j in 1:length(PaisesEnt)){
      print(PaisesEnt[[j]])
      EntPaisDur<-subset(entelpais,entelpais["Pais destinatario"]==PaisesEnt[[j]])
      
      PaisesEnt2[j]<-sum(EntPaisDur[["Duracion"]]/60)
      PaisesEnt3[j]<-length(EntPaisDur[["Geografia"]])
    }
    
    EntPais<-data.frame(PaisesEnt)
    
    EntPais["Duracion"]<-PaisesEnt2
    EntPais["Cantidad de llamadas"]<-PaisesEnt3
    EntPais3<-EntPais
    EntPais<-EntPais3[order(-EntPais3[["Duracion"]]),]
    
    rm(PaisesEnt,PaisesEnt2,PaisesEnt3,EntPaisDur,entelpais,EntPais3)
  }
  rm(cdr_entel,cdr_movistar)
  
  
  
  
  
  
  #Accesos unicos por tipo
  #Se quiere generar los consumos promedios mensuales por acceso
  uso2<-uso
  mes1 <- sapply(uso2["Fecha"], substr, 6, 7)
  mes <- as.numeric(mes1)
  uso2["Mes"] <- mes
  rm(mes1,mes)
  Vnacional<-as.numeric(uso2[["N. Voz nacional"]])
  Vnacional<-as.numeric(sapply(uso2["N. Voz nacional"], substr,1,nchar(uso2["N. Voz nacional"])))
  uso2["Vnacional"] <- Vnacional
  uso2[,"N. VozTotal"] <- uso2[["Vnacional"]] +as.numeric(uso2[,"N. Voz inter."])
  accesosunicos2<-as.list(unique(uso2["Acceso"]))
  rm(Vnacional)
  
  i<-1
  Vtotal2<-c()
  Vnacional2<-c()
  Vinter2<-c()
  VhaciaInt2<-c()
  VRoam2<-c()
  VRoamEnt2<-c()
  VRoamSal2<-c()
  nmeses<-c()
  for(i in 1:lengths(accesosunicos2)){
    AccMes<-subset(uso2,uso2["Acceso"] == as.character(accesosunicos2[["Acceso"]][i])) #Generamos un uso unico al cual le agregamos todos los datos de los distintos meses
    nmeses[i]<-as.numeric(lengths(AccMes["Acceso"])) #se ve cuantos meses ha sido ocupado el acceso correspondiente
    #Se obtienen los datos necesarios desde AccMes que contiene los mismos titulos de uso junto con los que se agregaron de 557 al 562
    Vtotal2[i]<-sum(AccMes["N. VozTotal"])/nmeses[[i]]/60
    Vnacional2[i]<- sum(AccMes["Vnacional"])/nmeses[[i]]/60
    Vinter2[i]<-sum(AccMes["N. Voz inter."])/nmeses[[i]]/60
    VhaciaInt2[i]<-sum(AccMes["N. Voz hacia Int"])/nmeses[[i]]/60
    VRoam2[i]<-sum(AccMes["N. Voz roaming"])/nmeses[[i]]/60
    VRoamEnt2[i]<-sum(AccMes["N. Voz Roaming In"])/nmeses[[i]]/60
    VRoamSal2[i]<-sum(AccMes["N. Voz Roaming Out"])/nmeses[[i]]/60
    
  }
  accesosunicos<-(unique(uso2["Acceso"])) #Se usa un formato compatible para la compilacion de datos
  #Se unen los datos recopilados a la nueva tabla accesosunicos
  accesosunicos["VTotal"]<-Vtotal2
  accesosunicos["Vnacional"]<-Vnacional2
  accesosunicos["Vinter"]<-Vinter2
  accesosunicos["VhaciaInt"]<-VhaciaInt2
  accesosunicos["VRoaming"]<-VRoam2
  accesosunicos["VRoaming Entrante"]<-VRoamEnt2
  accesosunicos["VRoaming Saliente"]<-VRoamSal2
  accesosunicos["Cantidad Meses"]<-nmeses
  
  rm(Vtotal2,VRoamSal2,VRoamEnt2,VRoam2,Vnacional2,Vinter2,VhaciaInt2,nmeses,accesosunicos2,AccMes,uso2,UTP_accesses2)
  
  
  
}
