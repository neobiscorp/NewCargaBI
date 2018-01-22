  upload<-function(x){
  killDbConnections <- function () {
    all_cons <- dbListConnections(MySQL())
    print(all_cons)
    for (con in all_cons)
      +  dbDisconnect(con)
    print(paste(length(all_cons), " connections killed."))
  }
  killDbConnections()
  client <- "tester"
  DB <- dbConnect(
    MySQL(),
    user = "root",
    password = "",
    dbname = paste0(client)
  )
  #dbGetQuery(DB, "SET NAMES 'utf8';")
  dbGetQuery(DB, "SET NAMES 'latin1';")
  dbWriteTable(
    DB,
    "tester",
    x,
    field.types = NULL,
    row.names = FALSE,
    overwrite = TRUE,
    append = FALSE,
    allow.keywords = FALSE
  )
  killDbConnections()
  }
upload(uso)

