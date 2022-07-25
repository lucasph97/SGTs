
#' Print a table in an worksheet
#'
#' @description Prints a table in a *.xlsx file with a nice format.
#'
#' @param .tab The table to be printed
#' @param .position The position in the spreadsheet where the table will be printed. Set by default as 1,1.
#' @param .output The name of the output file. Set by default as "coolTab_output.xlsx".
#' @param .sheet The name of the sheet where the table will be printed. Set by default as "Sheet".
#' @param .title The table title. Set by default as "This is my cool tab".
#' @param .reformat Set column and row names to title format; upper case for first letter. Set by default as "FALSE".
#'
#' @return
#'
#' @importFrom magrittr %>%
#' @importFrom stringr str_to_title
#' @import openxlsx
#'
#' @export
#'
printTab <- function(.tab, .position = c(1,1),
                     .output = "coolTab_output.xlsx",
                     .sheet = "Sheet",
                     .title = NULL,
                     .reformat = FALSE){

  # Create or load file
  if(file.exists(.output)){
    wb <- loadWorkbook(file = .output)
    if(!(.sheet %in% getSheetNames(.output))){
      addWorksheet(wb, sheetName = .sheet)
    }
  }else if(!file.exists(.output)){
    wb <- createWorkbook(.output)
    addWorksheet(wb, sheetName = .sheet)
  }

  # Set as data.frame if it is not and tidy colnames and rownames
  if(!is.data.frame(.tab)){
    .tab <- as.data.frame(.tab)
  }

  if(.reformat == TRUE){
    colnames(.tab) <- sapply(colnames(.tab), str_to_title)

    if(!is.null(rownames(.tab))){
      rownames(.tab) <- sapply(rownames(.tab), str_to_title)
    }
  }

  if(!is.null(.title)){
  # Print title
  writeData(wb, .title, sheet = .sheet, colNames = FALSE,
            startRow = .position[1], startCol = .position[2])
  addStyle(wb, sheet = .sheet, rows = .position[1], cols = .position[2],
           style = createStyle(halign = 'left',
                               valign = "center",
                               textDecoration = "bold",
                               fontSize = 14))
  .position[1] <- .position[1] + 1
  }

  if(is.null(rownames(.tab)) |
    mean(rownames(.tab) == c(1:nrow(.tab)))==1){

    writeData(wb, .tab, sheet = .sheet, colNames = TRUE, rowNames = FALSE,
              startRow = (.position[1]), startCol = .position[2],
              borders = "surrounding", borderStyle = "thick",
              headerStyle = createStyle(textDecoration = "bold",
                                        fontSize = 12,
                                        fgFill = "lightgrey"))

  }else if(!is.null(rownames(.tab))){
    writeData(wb, .tab, sheet = .sheet, colNames = TRUE, rowNames = TRUE,
              startRow = (.position[1]), startCol = .position[2],
              borders = "surrounding", borderStyle = "thick",
              headerStyle = createStyle(textDecoration = "bold",
                                        fontSize = 12,
                                        fgFill = "lightgrey"))

  }
  setColWidths(wb, sheet = .sheet, cols = 2:ncol(.tab), widths = "auto")

  saveWorkbook(wb, file = .output, overwrite = TRUE)

}

