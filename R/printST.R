
#' Print summary overall results from a Single Trial analysis in a workbook
#'
#' @description Prints a summary and overall results from a Single Trial
#' analysis in a *.xlsx file using openxlsx. The workbook is comprised of two worksheets.
#' The first one contains basic trial stats and summaries. The second one contains
#' a model summary.
#' This function considers the following column names: Trial, Genotype, Replicate, Row, Column. If
#' any of these columns do not posses these respective names, please be sure of renaming them.
#'
#'
#' @param asreml.obj An asreml object with the linear mixed model fitted for the analysis.
#' @return A *.xlsx file comprised of two spreadsheet containing the information from the list
#'
#' @import openxlsx
#' @importFrom magrittr %>%
#' @importFrom stringr str_split_fixed
#' @import dplyr
#'
#' @export
#'

printST <- function(asreml.obj,
                    .output = "ST_output.xlsx"){

  resultswb <- createWorkbook(.output)

  # Get the data from the model
  mydf <- get(asreml.obj$call$data)

  tt <- table(mydf$Genotype) %>% as.data.frame
  colnames(tt) <- c("Genotype","Frequency")

  sumtab <- mydf %>%
    summarise(Rows = n_distinct(Row),
              Columms = n_distinct(Column),
              Reps = n_distinct(Replicate),
              Genos = n_distinct(Genotype)) %>%
    as.data.frame()

  addWorksheet(resultswb, sheetName = 'FileChecks')
  nrow <- 1
  writeData(resultswb,'Data file: ****COMPLETE****', sheet = 'FileChecks',
                 startRow = nrow, startCol = 1, colNames = F)

  addStyle(resultswb, sheet = 'FileChecks', cols = 1L, rows = nrow,
           style = createStyle(textDecoration = "bold", fontSize = 16))
  nrow <- nrow + 1

  writeData(resultswb, paste('Trial',levels(mydf$Trial)), sheet = 'FileChecks',
                 startRow = nrow, startCol = 1, colNames = F)
  nrow <- nrow + 1

  writeData(resultswb, 'Number of Genotypes:', sheet = 'FileChecks',
                 startRow = nrow, startCol = 1, colNames = F)

  writeData(resultswb, n_distinct(mydf$Genotype),
                 sheet = 'FileChecks',
                 startRow = nrow, startCol = 2, colNames = F)
  nrow <- nrow + 2

  writeData(resultswb, "Summary stats:", sheet = 'FileChecks',
                 startRow = nrow, startCol = 1, colNames = F)
  addStyle(resultswb, sheet = 'FileChecks', cols = 1L, rows = nrow,
           style = createStyle(textDecoration = "bold", fontSize = 14))
  nrow <- nrow + 1

  writeData(resultswb, sumtab, sheet = 'FileChecks',
                 startRow = nrow, startCol = 1, colNames = T, rowNames = F,
            headerStyle = createStyle(textDecoration = "bold",
                                      fontSize = 12,
                                      borderStyle = "medium", borderColour = "black",
                                      fgFill = "lightgrey"))
   nrow <- nrow + nrow(sumtab) + 2

  writeData(resultswb, "Genotype frequency:", sheet = 'FileChecks',
                 startRow = nrow, startCol = 1, colNames = F)
  addStyle(resultswb, sheet = 'FileChecks', cols = 1L, rows = nrow,
           style = createStyle(textDecoration = "bold", fontSize = 14))
  nrow <- nrow + 1

  writeData(resultswb, tt, sheet = 'FileChecks',
                 startRow = nrow, startCol = 1, colNames = T, rowNames = F,
            headerStyle = createStyle(textDecoration = "bold",
                                      fontSize = 12,
                                      borderStyle = "medium", borderColour = "black",
                                      fgFill = "lightgrey"))
  nrow <- nrow + nrow(tt) + 2


  setColWidths(resultswb, sheet = 'FileChecks',
               cols = 1:(ncol(sumtab)), widths = 18)

  # ----------------------------------------------------------------------------
  ## End of checks spreadsheet ##

  sum.nice <- summary(asreml.obj)$varcomp%>%
    tibble::rownames_to_column("Component")

  wald.tab <- wald(asreml.obj, ssType = "conditional", denDF = "algebraic")$Wald %>%
    tibble::rownames_to_column("Term")
  wald.tab$Pr <- round(wald.tab$Pr,4)


  addWorksheet(resultswb, sheetName = 'ModelSummary')
  nrow <- 1

  writeData(resultswb, "Summary of the final model",
            sheet = 'ModelSummary',
            startRow = nrow, startCol = 1, colNames = F)
  addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = nrow,
           style = createStyle(textDecoration = "bold", fontSize = 16))
  nrow <- nrow + 2

  writeData(resultswb, "Summary of the variance components",
            sheet = 'ModelSummary',
            startRow = nrow, startCol = 1, colNames = F)
  addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = nrow,
           style = createStyle(textDecoration = "bold", fontSize = 14))
  nrow <- nrow + 1

  writeData(resultswb, sum.nice, sheet = 'ModelSummary',
            startRow = nrow, startCol = 1, colNames = T, rowNames = F,
            headerStyle = createStyle(textDecoration = "bold",
                                      fontSize = 12,
                                      borderStyle = "medium", borderColour = "black",
                                      fgFill = "lightgrey"))
  addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = (nrow):(nrow+nrow(sum.nice)),
           style = createStyle(textDecoration = "bold",
                               fontSize = 12,
                               borderStyle = "medium", borderColour = "black",
                               fgFill = "lightgrey"), gridExpand = TRUE)
  nrow <- nrow + nrow(sum.nice) + 2

  writeData(resultswb, "Wald table of fixed effects",
            sheet = 'ModelSummary',
            startRow = nrow, startCol = 1, colNames = F)
  addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = nrow,
           style = createStyle(textDecoration = "bold", fontSize = 14))

  nrow <- nrow + 1
  writeData(resultswb, wald.tab, sheet = 'ModelSummary',
            startRow = nrow, startCol = 1, colNames = T, rowNames = F,
            headerStyle = createStyle(textDecoration = "bold",
                                      fontSize = 12,
                                      borderStyle = "medium", borderColour = "black",
                                      fgFill = "lightgrey"))
  addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = (nrow):(nrow+nrow(wald.tab)),
           style = createStyle(textDecoration = "bold",
                               fontSize = 12,
                               borderStyle = "medium", borderColour = "black",
                               fgFill = "lightgrey"), gridExpand = TRUE)
  nrow <- nrow + nrow(wald.tab) + 1
  writeData(resultswb, "Note: Wald table was computed using ssType = \"conditional\" and denDF = \"algebraic\".",
            sheet = 'ModelSummary',
            startRow = nrow, startCol = 1, colNames = F)
  addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = nrow,
           style = createStyle(fontSize = 10))


  setColWidths(resultswb, sheet = 'ModelSummary',
               cols = 1:(ncol(wald.tab)), widths = 18)
  saveWorkbook(resultswb, file = .output, overwrite = TRUE)
}


