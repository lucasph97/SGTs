
#' Summary stats of recorded traits
#'
#' @description Provides a data frame with minimum, median and maximum values from traits of
#' interest and the number of trials that the trait was measured for each genotype, together with pedigree
#' information, if any. The output from this function can be merged with predicted means from other analysed
#' traits. Once the table of summary statistics is complete, with or without additional predicted means,
#' everything should be ready to use printSumStats, the function to print this table in a nice way in an
#' *.xlsx file.
#'
#' @param .tab A data frame object containing the table with the summary stats of the traits of interest
#' for each genotype.
#' @param .output An *.xlsx file to write the outputs. The default is "SumStats_output.xlsx".
#'
#' @import openxlsx
#' @importFrom magrittr %>%
#' @importFrom stringr str_split_fixed
#' @import dplyr
#'
#' @export
#'
#'

printStats <- function(.tab, .output = "SumStats_output.xlsx"){

  resultswb <- createWorkbook(.output)


  ## Put all check summaries into one Excel spreadsheet

  if("Pedigree" %in% colnames(.tab)){

    addWorksheet(resultswb, sheetName = "Summary")

    writeData(resultswb, "Genotype", sheet = 'Summary', colNames = FALSE,
              startRow = 1, startCol = 1)
    mergeCells(resultswb, sheet = "Summary", rows = 1:2, cols = 1L)

    writeData(resultswb, "Pedigree", sheet = 'Summary',
              startRow = 1, startCol = 2, colNames = FALSE)

    mergeCells(resultswb, sheet = "Summary", rows = 1:2, cols = 2L)

    writeData(resultswb, select(.tab,c("Genotype","Pedigree")), sheet = 'Summary',
              startRow = 3, startCol = 1, colNames = FALSE)

    setColWidths(resultswb, sheet = 'Summary', cols = c(1,2), widths = "auto")

    addStyle(resultswb, sheet = 'Summary', cols = c(1L,2L), rows = c(1L),
             style = createStyle(halign = 'left',
                                 valign = "center",
                                 textDecoration = "bold",
                                 fgFill = "lightgrey"))

    # saveWorkbook(resultswb, file = .output, overwrite = TRUE)

    # Get names for the header
    temp <- colnames(.tab)[c(3:ncol(.tab))]

    headers <- lapply(stringi::stri_split_regex(stringi::stri_reverse(temp),
                                                pattern = '[:punct:]', n = 2),
                      stringi::stri_reverse)

    headers <- setNames(data.table::transpose(headers)[2:1],
                        c('output1', 'output2')) %>%
      as.data.frame %>% t
    rownames(headers) <- NULL

    writeData(resultswb, headers, sheet = 'Summary',
              startRow = 1, startCol = 3, colNames = FALSE)

    # For practical purposes
    headers <- cbind("Genotype","Pedigree",headers)

    traits <- unique(headers[1,])

    # From colors()
    nicecolors <- c("aquamarine","beige","coral","deepskyblue",
                    "honeydew","lavender","lemonchiffon","lightblue","lightpink",
                    "mintcream","mistyrose","orchid","palegreen","paleturquoise",
                    "seagreen","salmon","skyblue")

    for(i in 3:(length(traits))){
      pos <- which((headers[1,] %in% traits[i]) & !(headers[2,] %in% "Summary"))

      mergeCells(resultswb, sheet = "Summary", rows = 1L, cols = c(pos[1]:tail(pos,1)))

      addStyle(resultswb, sheet = 'Summary', cols = pos[1]:(tail(pos,1)+1), rows = 1:2,
               style = createStyle(halign = 'center',
                                   valign = "center",
                                   textDecoration = "bold",
                                   fgFill = nicecolors[i-2]), gridExpand = T)

      if("Summary" %in% headers[2,pos[1]:(tail(pos,1)+1)]){
        groupColumns(resultswb, sheet = "Summary", cols = pos[1]:tail(pos,1))
      } else if(!("Summary" %in% headers[2,pos[1]:(tail(pos,1)+1)])){
        .tab[,pos[1]:(tail(pos,1))] <- round(.tab[,pos[1]:(tail(pos,1))],2)
      }

    }

    freezePane(resultswb, "Summary",
      firstActiveRow = 3L,
      firstActiveCol = 3)

  }else if(!("Pedigree" %in% colnames(.tab))){
    addWorksheet(resultswb, sheetName = "Summary")

    writeData(resultswb, "Genotype", sheet = 'Summary', colNames = FALSE,
              startRow = 1, startCol = 1)
    mergeCells(resultswb, sheet = "Summary", rows = 1:2, cols = 1L)

    writeData(resultswb, select(.tab,c("Genotype")), sheet = 'Summary',
              startRow = 3, startCol = 1, colNames = FALSE)

    setColWidths(resultswb, sheet = 'Summary', cols = 1L, widths = "auto")

    addStyle(resultswb, sheet = 'Summary', cols = 1L, rows = c(1L),
             style = createStyle(halign = 'left',
                                 valign = "center",
                                 textDecoration = "bold",
                                 fgFill = "lightgrey"))

    # saveWorkbook(resultswb, file = .output, overwrite = TRUE)

    # Get names for the header
    temp <- colnames(.tab)[c(2:ncol(.tab))]

    headers <- lapply(stringi::stri_split_regex(stringi::stri_reverse(temp), pattern = '[:punct:]', n = 2),
                      stringi::stri_reverse)

    headers <- setNames(data.table::transpose(headers)[2:1], c('output1', 'output2')) %>%
      as.data.frame %>% t
    rownames(headers) <- NULL

    writeData(resultswb, headers, sheet = 'Summary',
              startRow = 1, startCol = 2, colNames = FALSE)

    # For practical purposes
    headers <- cbind("Genotype",headers)

    traits <- unique(headers[1,])

    # From colors()
    nicecolors <- c("aquamarine","beige","coral","deepskyblue",
                    "honeydew","lavender","lemonchiffon","lightblue","lightpink",
                    "mintcream","mistyrose","orchid","palegreen","paleturquoise",
                    "seagreen","salmon","skyblue")

    for(i in 2:(length(traits))){
      pos <- which((headers[1,] %in% traits[i]) & !(headers[2,] %in% "Summary"))

      mergeCells(resultswb, sheet = "Summary", rows = 1L, cols = c(pos[1]:tail(pos,1)))

      addStyle(resultswb, sheet = 'Summary', cols = pos[1]:(tail(pos,1)+1), rows = 1:2,
               style = createStyle(halign = 'center',
                                   valign = "center",
                                   textDecoration = "bold",
                                   fgFill = nicecolors[i-2]), gridExpand = T)

      if("Summary" %in% headers[2,pos[1]:(tail(pos,1)+1)]){
        groupColumns(resultswb, sheet = "Summary", cols = pos[1]:tail(pos,1))
      } else if(!("Summary" %in% headers[2,pos[1]:(tail(pos,1)+1)])){
        .tab[,pos[1]:(tail(pos,1))] <- round(.tab[,pos[1]:(tail(pos,1))],2)
      }

    }

    freezePane(resultswb, "Summary",
               firstActiveRow = 2,
               firstActiveCol = 2)
  }
  writeData(resultswb, select(.tab,-c("Genotype")), sheet = 'Summary',
                              startRow = 3, startCol = 2, colNames = FALSE)

            addStyle(resultswb, sheet = 'Summary', cols = 2:ncol(.tab), rows = 3:nrow(.tab),
                     style = createStyle(halign = 'center',
                                         valign = "center"), gridExpand = T)

            saveWorkbook(resultswb, file = .output, overwrite = TRUE)

}
