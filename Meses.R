Meses<-function(i,x){ #Meses(i,x) donde x es una columna de Fechas ordenadas por YYYY/MM/DD e "i" es una fila en especifica de la columna, el resultado obtenido es la cantidad de meses transcurridos desde el mes inicial hasta el mes del dato "i"
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
              AAA["year"]==max(year))
  rm(month,year)
  fin1<-max(BBB[["month"]])
  fin2<-max(BBB[["year"]])
  rm(BBB)
  return((AAA[["year"]][i]-fin2)*12+fin1-AAA[["month"]][i]+1)
  rm(AAA)
}



