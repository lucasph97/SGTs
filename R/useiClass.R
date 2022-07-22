
#' Use Interaction Classes
#'
#' @description Prints the results from the use of interaction classes, iClass method proposed by Smith et al. 2021.
#' in a *.xlsx file using openxlsx. The function also returns a list containing the results.
#'
#' @param .fa.asreml.object The fa.asreml object from ASExtras4 package
#' @param .factors The factors that will be used to create the interaction classes. Set by default as the
#' order of the FA model, namely, all factors are considered. It is not necessary to specify contiguous factors.
#' For example, it is possible to specify factors 1, 3 and 7, without considering factors 2, 4, 5 and 6.
#' @param .print If you wish to print the outputs in an excel file, please set as TRUE. Set as FALSE by default.
#' @param .output An *.xlsx file to write the outputs. The default is "iClass_output.xlsx".
#'
#' @return A list containing
#' @return - $iClasses the A data frame containing a table of the composition of the iClasses
#' @return - $iClassOP_wide A data frame containing the iClass Overall Performance measures in wide format
#' @return - $iClassOP_long A data frame containing the iClass Overall Performance measures in long format
#' @return A spreadsheet containing the information from the list. The spreadsheet can be printed in a new or
#' existing *.xlsx file.
#'
#' @import openxlsx
#' @importFrom magrittr %>%
#' @importFrom stringr str_split_fixed
#' @import dplyr
#'
#' @export
#'

useiClass <- function(.fa.asreml.obj,
                    .factors = NULL,
                    .print = FALSE,
                    .output = "iClass_output.xlsx"){

  gen.comp.name <- names(.fa.asreml.obj$gammas)
  gen.comp.str <- stringr::str_split_fixed(string = gen.comp.name,
                                           pattern = "[\\|,:,+.\\(,\\)]+", n = 2)[[1]][1]

  fa.order <- stringr::str_extract(gen.comp.name, "[0-9]+") %>% as.numeric

  trial_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+", string = gen.comp.name)[[1]][2]
  geno_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+", string = gen.comp.name)[[1]][4]

  if(is.null(.factors)){
    .factors <- 1:fa.order
  } else if(length(.factors) == 1){
    .factors <- 1:.factors
  }

  factors_names <- paste0("fac_",.factors)

  ## Rotated loadings
  RotLambda <- .fa.asreml.obj[[1]][[1]]$'rotated loads' %>% as.data.frame %>%
    tibble::rownames_to_column(trial_name) %>% select(all_of(c(trial_name, factors_names)))

  iC_mat <- matrix(NA, nrow = nrow(RotLambda), ncol = ncol(RotLambda)) %>%
    as.data.frame
  colnames(iC_mat) <- colnames(RotLambda)
  iC_mat[,1] <- RotLambda[,1]


  for (j in 2:ncol(RotLambda)){
    for(i in 1:nrow(RotLambda)){
      if(RotLambda[i,j] > 0){
        iC_mat[i,j] <- "p"
      } else if (RotLambda[i,j] < 0){
        iC_mat[i,j] <- "n"
      }
    }
  }

  RotLambda_iC <- as.data.frame(RotLambda)

  RotLambda_iC$iClass <- as.factor(do.call(paste0,iC_mat[, 2:ncol(RotLambda_iC)]))

  RotLambda_iC <- RotLambda_iC[order(RotLambda_iC$iClass, decreasing = T),]

  # iClass Overall Performance measures

  # Mean lambda sub rw - Mean of the loadings for factor r across the environments
  # in iClass w

  iC_meanLambda <- RotLambda_iC[,-1] %>% group_by(iClass) %>%
    summarise_all(~mean(.))

  ## Rotated Scores
  RotScores <- matrix(.fa.asreml.obj[[2]][[1]]$scores$blupr, ncol = fa.order, byrow = F) %>%
    as.data.frame()

  dimnames(RotScores) <- list(unique((.fa.asreml.obj[[2]][[1]]$scores %>%
                                        dplyr::select(., geno_name))[,1]),
                              unique((.fa.asreml.obj[[2]][[1]]$scores %>%
                                        dplyr::select(., all_of(trial_name)))[,1]))

  RotScores <- RotScores %>%
    select(paste0("Comp", .factors))

  ## iClass OP
  iC_meanLambda.matrix <- iC_meanLambda %>%
    tibble::column_to_rownames(var = "iClass") %>% as.matrix()

  iClassOP.wide <- iC_meanLambda.matrix%*%t(RotScores) %>% t %>% as.data.frame() %>%
    tibble::rownames_to_column(geno_name) %>% relocate(geno_name)

  iClassOP.long <- tidyr::pivot_longer(data = iClassOP.wide,
                                       cols = colnames(iClassOP.wide)[-1],
                                       names_to = "iClass", values_to = "OP")

  # Compoisition of iClasses
  iClasses <- RotLambda_iC %>% select(all_of(trial_name),iClass)
  iClasses_table <- cbind(c("iClass", paste0(trial_name, "s")), table(iClasses$iClass))

    if(.print == TRUE){
    # Print the results in a *.xlsx file ----
    # Create or load file
    if(file.exists(.output)){
      wb <- loadWorkbook(file = .output)
    }else{
      wb <- createWorkbook(.output)
    }
    # iClasses composition
    addWorksheet(wb, sheetName = "iClasses")

    writeData(wb,'iClass method output', sheet = "iClasses",
              startRow = 1, startCol = 1, colNames = F)

    addStyle(wb, sheet = "iClasses", cols = 1L, rows = 1L,
             style = createStyle(textDecoration = "bold", fontSize = 16))

    writeData(wb, paste0("The number of iClasses was ", n_distinct(iClasses$iClass)),
              sheet = "iClasses", startRow = 2, startCol = 1, colNames = F)

    writeData(wb, iClasses, sheet = "iClasses", startRow = 4, startCol = 1,
              colNames = T, rowNames = F,
              headerStyle = createStyle(textDecoration = "bold",
                                        fontSize = 12,
                                        borderStyle = "medium",
                                        borderColour = "black",
                                        fgFill = "lightgrey"))
    # iClassOP in wide format
    addWorksheet(wb, sheetName = "iClassOP_wide")

    writeData(wb, iClassOP.wide, sheet = "iClassOP_wide", startRow = 1, startCol = 1,
              colNames = T, rowNames = F,
              headerStyle = createStyle(textDecoration = "bold",
                                        fontSize = 12,
                                        borderStyle = "medium",
                                        borderColour = "black",
                                        fgFill = "lightgrey"))
    # iClassOP in long format
    addWorksheet(wb, sheetName = "iClassOP_long")

    writeData(wb, iClassOP.long, sheet = "iClassOP_long", startRow = 1, startCol = 1,
              colNames = T, rowNames = F,
              headerStyle = createStyle(textDecoration = "bold",
                                        fontSize = 12,
                                        borderStyle = "medium",
                                        borderColour = "black",
                                        fgFill = "lightgrey"))
    #### Save file----
    saveWorkbook(wb, file = .output, overwrite = TRUE)
  }
  #### Save outputs in a list to return ----
  sum_list <- list()
  sum_list$iClasses <- iClasses
  sum_list$iClassOPs_wide <- iClassOP.wide
  sum_list$iClassOPs_long <- iClassOP.wide

  return(sum_list)
}
