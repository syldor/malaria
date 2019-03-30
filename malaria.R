library("data.table")
library("plyr")
library("openxlsx")

setwd("/home/syldor/Projects/malaria")
results <- as.data.table(read.csv2('main_data.csv', header = TRUE, sep=','))
questions <- as.data.table(read.csv2('questions.csv', header = TRUE, sep=',', stringsAsFactors = FALSE))

results[results == 'Other (specify):____________']<-""

for (idx in seq(1:8)) {
  setnames(results,paste0("T_GI_", idx), paste0("GI", idx))
}

for (letter in c('a', 'b', 'c', 'd', 'e', 'f', 'g', 'h')) {
  setnames(results,paste0("HR2_", letter), paste0("HR2", letter))
}
setnames(results,'HR2_i1_Specify', 'HR2i1Specify')
setnames(results,'HR2_i1', 'HR2i1')

setnames(results,'HR2_i2_Specify', 'HR2i2Specify')
setnames(results,'HR2_i2', 'HR2i2')
setnames(results,'HR2_j1_Specify', 'HR2j1Specify')
setnames(results,'HR2_j1', 'HR2j1')
setnames(results,'HR2_j2_Specify', 'HR2j2Specify')
setnames(results,'HR2_j2', 'HR2j2')

setnames(results,"GI3", "District")

districts_list <- data.frame(District = results[, "District"])


multipleAnswersQ <- function(question, results, questions) {
  districts_list <- data.frame(District = results[, "District"])
  codes <- names( results )
  headers <- c(question, codes[grepl(paste0(question, "_"), codes)])
  headers <- names(results)[names(results) %in% headers]
  resultsForQuestion <-results[, names(results) %in% c("District", headers), with=FALSE]
  resultsLong <- melt(resultsForQuestion, id.vars = c("District"), measure.vars = headers)
  resultsConcatenate <- resultsLong[value != "", .(value = paste(value, collapse="\n")), by = District]
  names(resultsConcatenate)[names(resultsConcatenate) == 'value'] <- questions[Code==question, Question][1]
  

  finalResult <- merge(districts_list, resultsConcatenate, by="District", all.x=TRUE)
  finalResult[is.na(finalResult)] <- "-"
  return(finalResult)
}


getQuestionsOfModule <- function(module) {
  codes <- names(results)
  headers <- codes[grepl(module, codes)]
  headers2 <- strsplit(headers, "_")
  headers3 <- lapply(headers2, function(l) l[[1]])
  headers4 <- unique(headers3)  
  headers5 <- unlist(headers4, use.names = FALSE)
  return(headers5)
}

drawTable <- function(data, wb, idx, sheetNumber) {
  numberOfDistricts <- 10
  padding <- 1
  header <- 1
  
  tableSize <- numberOfDistricts + padding + header
  headerRow <- (idx - 1) * tableSize + padding + header
  
  headerStyle <- createStyle(fontSize = 14, fontColour = "#FFFFFF", halign = "center",
                             fgFill = "#4F81BD", border="TopBottom", borderColour = "#4F81BD", wrapText = TRUE)
  
  tableStyle <- createStyle(border= "TopBottomLeftRight", valign="top", wrapText = TRUE)
  
  tableRange <- seq(from= (idx - 1) * tableSize + padding + header, to=(idx - 1) * tableSize + padding + header + numberOfDistricts)
  
  addStyle(wb, sheet = sheetNumber, tableStyle, rows=tableRange, cols=1:5, gridExpand = TRUE)
  addStyle(wb, sheet = sheetNumber, headerStyle, rows=headerRow, cols=1:5, gridExpand = TRUE)
  
  writeData(wb, sheet = sheetNumber, x = data, startCol = 1, startRow = headerRow)
}

group_questions <- function(res, idx, districts_list) {
  df <- districts_list
  start <-  min(4* (idx-1)+1, length(res))
  end <- min(4*idx, length(res))
  for (i in seq(from=start, to=end)) {
    df <- merge(df, res[[i]], by="District", all.x=TRUE)
  }
  return(df)
}

generate_module_sheets <- function(idx, modules_list, wb, results, questions) {
  module <- modules_list[idx]
  sheetNumber <- idx
  ACMultipleQList <- getQuestionsOfModule(module)
  res <- lapply(ACMultipleQList, multipleAnswersQ, results = results, questions = questions)
  
  grouped_res <- list()
  
  numberOfGroups <- ceiling(length(res)/ 4)
  for (idx in seq(from = 1, to=numberOfGroups)) {
    grouped_res[[idx]] <- group_questions(res, idx, districts_list)
  }
  
  addWorksheet(wb, sheetName = module)
  
  setColWidths(wb, sheet = sheetNumber, cols = 1, widths = 15)
  setColWidths(wb, sheet = sheetNumber, cols = 2:20, widths = 60)
  
  for (idx in seq(from = 1, to=length(grouped_res))) {
    drawTable(grouped_res[[idx]], wb, idx, sheetNumber)
  }
  return(wb)
}

wb <- createWorkbook()
modules_list <- c('GI', 'AC', 'WP', 'HR', 'TR', 'KDA', 'SV', 'SC', 'VC', 'SR', 'FR', 'CC')


for (moduleIdx in seq(from = 1, to=length(modules_list))) {
  wb <- generate_module_sheets(moduleIdx, modules_list, wb, results, questions)
}


saveWorkbook(wb, "/home/syldor/VB-shared/AC.xlsx", overwrite = TRUE)




