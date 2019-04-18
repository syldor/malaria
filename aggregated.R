library("data.table")
library("dplyr")
library("tidyr")
library("openxlsx")

setwd("/home/syldor/Projects/malaria")
results <- as.data.table(read.csv2('main_data.csv', header = TRUE, sep=','))
questions <- as.data.table(read.csv2('src/questions.csv', header = TRUE, sep=',', stringsAsFactors = FALSE))

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


setnames(results, "VC26_4_OTH", "VC16_4_OTH")
setnames(results, "SR1_1", "SR2_1")
setnames(results, "T_CC3_1", "CC3_1")
setnames(results, "T_CC3_2", "CC3_2")
setnames(results, "T_CC3_3", "CC3_3")



districts_list <- data.frame(District = results[, "District"])


multipleAnswersQ <- function(question, results, questions) {
  districts_list <- data.frame(District = results[, "District"])
  codes <- names( results )
  headers <- c(question, codes[grepl(paste0(question, "_"), codes)])
  headers <- names(results)[names(results) %in% headers]
  resultsForQuestion <-results[, names(results) %in% c("District", headers), with=FALSE]
  resultsLong <- melt(resultsForQuestion, id.vars = c("District"), measure.vars = headers)
  names(resultsLong)[names(resultsLong) == 'value'] <- questions[Code==question, Question][1]

  return(resultsLong)
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



drawTableAgg <- function(data, wb, rowIdx, sheetNumber) {
  
  headerRow <- rowIdx
  
  headerStyle <- createStyle(fontSize = 14, fontColour = "#FFFFFF", halign = "center",
                             fgFill = "#4F81BD", border="TopBottom", borderColour = "#4F81BD", wrapText = TRUE)
  
  tableStyle <- createStyle(border= "TopBottomLeftRight", valign="top", wrapText = TRUE)
  
  tableRange <- seq(from= rowIdx, to=rowIdx + nrow(data))
  
  addStyle(wb, sheet = sheetNumber, tableStyle, rows=tableRange, cols=1:5, gridExpand = TRUE)
  addStyle(wb, sheet = sheetNumber, headerStyle, rows=headerRow, cols=1:5, gridExpand = TRUE)
  
  writeData(wb, sheet = sheetNumber, x = data, startCol = 1, startRow = headerRow)
}

generate_module_agg_sheets <- function(idx, modules_list, wb, results, questions) {
  module <- modules_list[idx]
  sheetNumber <- idx
  ACMultipleQList <- getQuestionsOfModule(module)
  res <- lapply(ACMultipleQList, multipleAnswersQ, results = results, questions = questions)
  
  grouped_res <- list()
  
  for (idx in seq(from = 1, to=length(res))) {
    df <- res[[idx]]
    
    question <- names(df)[3]
    names(df) <- c("district", "variable", "question")
    df <- df[df$question != ""]
    df %<>% group_by(question) %>% summarise(n = n())
    names(df) <- c(question, "answers")
    grouped_res[[idx]] <- df
  }
  
  addWorksheet(wb, sheetName = module)
  
  setColWidths(wb, sheet = sheetNumber, cols = 1, widths = 60)
  setColWidths(wb, sheet = sheetNumber, cols = 2:20, widths = 20)
  rowIdx = 1
  for (idx in seq(from = 1, to=length(grouped_res))) {
    drawTableAgg(grouped_res[[idx]], wb, rowIdx, sheetNumber)
    rowIdx = rowIdx + 1 + nrow(grouped_res[[idx]])
  }
  return(wb)
}

## Aggregated Data


wb <- createWorkbook()
modules_list <- c('GI', 'AC', 'WP', 'HR', 'TR', 'KDA', 'SV', 'SC', 'VC', 'SR', 'FR', 'CC')

ACMultipleQList <- getQuestionsOfModule('GI')
res <- lapply(ACMultipleQList, multipleAnswersQ, results = results, questions = questions)


for (moduleIdx in seq(from = 1, to=length(modules_list))) {
  wb <- generate_module_agg_sheets(moduleIdx, modules_list, wb, results, questions)
}

saveWorkbook(wb, "/home/syldor/VB-shared/malaria_survey_agg.xlsx", overwrite = TRUE)


module <- modules_list[2]
ACMultipleQList <- getQuestionsOfModule('AC')

question <- 'AC5'
districts_list <- data.frame(District = results[, "District"])
codes <- names( results )
headers <- c(question, codes[grepl(paste0(question, "_"), codes)])
headers <- names(results)[names(results) %in% headers]
resultsForQuestion <-results[, names(results) %in% c("District", headers), with=FALSE]
resultsLong <- melt(resultsForQuestion, id.vars = c("District"), measure.vars = headers)
names(resultsLong)[names(resultsLong) == 'value'] <- questions[Code==question, Question][1]
resultsLong

question <- names(resultsLong)[3]
names(resultsLong) <- c("district", "variable", "question")

resultsLong <- resultsLong[resultsLong$question != ""]


resultsLong %<>% group_by(question) %>% summarise(n = n())
resultsLong




