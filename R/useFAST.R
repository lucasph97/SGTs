
#' Use Factor Analytic Selection Tools
#'
#' @description Prints the results from the use of Factor Analytic Selection Tools, FAST, in a *.xlsx file
#' using openxlsx. The function also returns a list containing the results.
#'
#' @param .asreml.obj An asreml object with the linear mixed model fitted for the analysis.
#' @param .fa.asreml.object A fa.asreml object from ASExtras4 package. Set as NULL by default. This
#' object is necesary only if a FA model was fitted
#' @param  .aii.pedicure Set as NULL by default. If molecular marker data was used to compute the kinship
#' matrix during the analysis, the mean value of the diagonal elements of the kinship matrix prior to
#' the scaling must be introduced in this argument. The mean value of the diagonal elements of the kinship
#' matrix is used to compute the additive genetic variance.
#' @param .print If you wish to print the outputs in an excel file, please set as TRUE. Set as FALSE by default.
#' @param .output An *.xlsx file to write the outputs. The default is "FAST_output.xlsx".
#' @param .sheet The name of the sheet where the table will be printed. Set by default as "FAST-additive" if
#' .nonadd = FALSE and as FAST-total if .nonadd = TRUE.
#'
#' @return A list containing
#' @return - $OP the Overall Performance measure
#' @return - $RMSD the Root Mean Square Deviation measure
#' @return - $R_Fi Responsiveness for factor i. There will be as much responsiveness objects within the list as
#' factors -1 there are.
#' @return - $adjBLUP the adjusted BLUPs.
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

useFAST <- function(.asreml.obj,
                    .fa.asreml.obj = NULL,
                    .output = "FAST_output.xlsx",
                    .print = FALSE,
                    .sheet,
                    .aii.pedicure = NULL){
  # Get the data from the model
  mydf <- get(.asreml.obj$call$data)

  sum.nice <- summary(.asreml.obj, vparameters = TRUE)$vparameters

  # Was any genetic information being considered for the analysis?
  if(any(grepl("vm", names(sum.nice)))){
    knownG <- TRUE
  } else{
    knownG <- FALSE
  }

  # Using genetic information ----
  if(knownG == TRUE){
    ### 1.1.1. Additive component name
    add.comp.name <- names(sum.nice)[grep("vm", names(sum.nice))]
    add.comp.str <- stringr::str_split_fixed(string = add.comp.name,
                                             pattern = "\\(", n = 2)[[1]][1]

    add.fa.order <- stringr::str_extract(add.comp.name, "[0-9]+") %>% as.numeric

    relationship_mat.name <-  stringr::str_split_fixed(add.comp.name, ",", n = 3)[[1]][3] %>%
      substr(2, nchar(.)-1)

    # Get Trial and Genotype names from data
    trial_name <- stringr::str_split(pattern = "[\\,:|,+.\\(,\\)]+", string = add.comp.name)[[1]][2]

    geno_name <- stringr::str_split(pattern = "[\\,:|,+.\\(,\\)]+", string = add.comp.name)[[1]][5]

    # Was the non-additive genetic covariance accounted for?
    if(sum(names(.asreml.obj$G.param) %>% stringr::str_detect(string = .,"fa")) == 1){
      nonadd = FALSE
    } else if(sum(names(.asreml.obj$G.param) %>% stringr::str_detect(string = .,"fa")) == 2){
      nonadd = TRUE
    }

    # FA fitted only for additive genetic variance ----
    if(nonadd == FALSE){

      # Number of trials
      ntrials <- dplyr::n_distinct(dplyr::select(get(.asreml.obj$call$data),
                                                 all_of(trial_name)))

      ## Rotated loadings
      AddRotLambda <- .fa.asreml.obj[[1]][[1]]$'rotated loads' %>% as.data.frame %>%
        tibble::rownames_to_column(trial_name)

      ## Rotated Scores
      AddRotScores <- matrix(.fa.asreml.obj$blups[[1]]$scores$blupr, ncol = add.fa.order, byrow = F)

      dimnames(AddRotScores) <- list(unique((.fa.asreml.obj$blups[[1]]$scores %>%
                                               dplyr::select(., all_of(geno_name)))[,1]),
                                     unique((.fa.asreml.obj$blups[[1]]$scores %>%
                                               dplyr::select(., all_of(trial_name)))[,1]))

      AddRotScores.df <- AddRotScores %>% as.data.frame %>%
        tibble::rownames_to_column(geno_name)

      ### OP and RMSD (factor 1) ----
      AddRotScores.df$OP <- mean(AddRotLambda[,2])*AddRotScores.df[,2]

      ## Blups (data frame, only reg blups)
      addblups.df <- .fa.asreml.obj$blups[[1]]$blups[,c(2,3,1,4)]

      addblups.df$adjblup1 <- addblups.df$regblup - as.vector(AddRotScores[,1] %*% t(AddRotLambda[,2]))

      AddRotScores.df$RMSD <- as.vector(sqrt(tapply(addblups.df$adjblup1^2,
                                                    dplyr::select(addblups.df, all_of(geno_name))[,1], mean)))

      AddRotScores.df <- AddRotScores.df[order(AddRotScores.df$OP,
                                               rev(AddRotScores.df$RMSD),
                                               decreasing = TRUE),]


      ## Responsiveness and adjusted BLUPs (remaining factors) ----
      if(add.fa.order == 2){
        pos <- mean(AddRotLambda[AddRotLambda[, 3] > 0, 3])
        neg <- mean(AddRotLambda[AddRotLambda[, 3] < 0, 3])

        AddRotScores.df$RESP <- (pos-neg) * AddRotScores.df[, paste0("Comp",2)]
        names(AddRotScores.df)[names(AddRotScores.df) == "RESP"] <- paste0("RESP", 2)

        addblups.df$adjblup2 <- addblups.df$regblup - as.vector(RotScores[,2] %*% t(AddRotLambda[,3]))

      } else if (add.fa.order > 2){
        for (i in 3:(add.fa.order + 1)) {
          pos <- mean(AddRotLambda[AddRotLambda[, i] > 0, i])
          neg <- mean(AddRotLambda[AddRotLambda[, i] < 0, i])

          AddRotScores.df$RESP <- (pos-neg) * AddRotScores.df[, paste0("Comp", i-1)]
          names(AddRotScores.df)[names(AddRotScores.df) == "RESP"] <- paste0("RESP", i-1)
        }
        for (i in 3:(add.fa.order)) {
          addblups.df$adjblup <- addblups.df$regblup - as.vector(AddRotScores[,2:(i-1)] %*% t(AddRotLambda[,3:i]))
          names(addblups.df)[names(addblups.df) == "adjblup"] <- paste0("adjblup", i-1)
        }
      }

      addblups.df <- merge(addblups.df, AddRotLambda)
        if(.print == TRUE){
        ## Print the results in a *.xlsx file ----
        .sheet = "FAST-additive"
        # Create or load file
        if(file.exists(.output)){
          wb <- loadWorkbook(file = .output)
          if(!(.sheet %in% getSheetNames(.output))){
            addWorksheet(wb, sheetName = .sheet)
          }
        }else{
          wb <- createWorkbook(.output)
          addWorksheet(wb, sheetName = .sheet)
        }
        #### Title----
        writeData(wb,'Factor Analytic Selection Tools (FAST) Output', sheet = .sheet,
                  startRow = 1, startCol = 1, colNames = F)
        # Title style
        addStyle(wb, sheet = .sheet, cols = 1L, rows = 1:2,
                 style = createStyle(textDecoration = "bold", fontSize = 16))

        #### Table of results----
        writeData(wb, AddRotScores.df, sheet = .sheet, startRow = 2, startCol = 1,
                  colNames = T, rowNames = F,
                  headerStyle = createStyle(textDecoration = "bold",
                                            fontSize = 12,
                                            borderStyle = "medium",
                                            borderColour = "black",
                                            fgFill = "lightgrey"))

        # Set color scale for OP and RMSD measures
        conditionalFormatting(wb, sheet = .sheet,
                              cols = which(colnames(AddRotScores.df) == "OP"),
                              rows = 3:(3 + nrow(AddRotScores.df)),
                              style = c("#F8696B","white","#63BE7B"),
                              type = "colourScale")

        conditionalFormatting(wb, sheet = .sheet,
                              cols = which(colnames(AddRotScores.df) == "RMSD"),
                              rows = 3:(3 + nrow(AddRotScores.df)),
                              style = c("#63BE7B","white","#F8696B"),
                              type = "colourScale")
        # Set column width
        setColWidths(wb, cols = 1:ncol(AddRotScores.df), sheet = .sheet,
                     widths = "auto", ignoreMergedCells = TRUE)
        #### Save file----
        saveWorkbook(wb, file = .output, overwrite = TRUE)
      }
      #### Save outputs in a list to return ----
      sum_list <- list()
      sum_list$FAST_additive <- AddRotScores.df
      sum_list$Adj_additive_BLUPs <- addblups.df
    }
    # FA fitted for both additive and non-additive genetic variances ----
    if(nonadd == TRUE){

      # Number of trials
      ntrials <- dplyr::n_distinct(dplyr::select(get(.asreml.obj$call$data),
                                                 all_of(trial_name)))

      ## FA terms names and order
      add.comp.name <- names(sum.nice)[grep("vm", names(sum.nice))]
      add.comp.str <- stringr::str_split_fixed(string = add.comp.name,
                                               pattern = "\\(", n = 2)[[1]][1]

      add.fa.order <- stringr::str_extract(add.comp.name, "[0-9]+") %>% as.numeric

      relationship_mat.name <-  stringr::str_split_fixed(add.comp.name, ",", n = 3)[[1]][3] %>%
        substr(2, nchar(.)-1)

      # Get Trial and Genotype names from data
      trial_name <- stringr::str_split(pattern = "[\\,:|,+.\\(,\\)]+", string = add.comp.name)[[1]][2]

      geno_name <- stringr::str_split(pattern = "[\\,:|,+.\\(,\\)]+", string = add.comp.name)[[1]][5]

      nonadd.comp.name <- names(sum.nice)[grep("ide", names(sum.nice))]
      nonadd.fa.order <- stringr::str_extract(nonadd.comp.name, "[0-9]+") %>% as.numeric

      # It is necessary to obtain unrotated loadings to scale them and rotate them afterwards
      # (Smith and Cullis 2018)

      # Get diagonal of relationship matrix to scale the additive loadings and scores

      # If we are using pedigree, we can obtain the aii value from there
      if("inbreeding" %in% names(attributes(get(relationship_mat.name)))){
        aii <-  mean(1 + attr(get(relationship_mat.name), "inbreeding"))
        # However if we are using markers, we need to input the aii value as the matrix
        # that can be obtained from .asreml.obj is already scaled by that value
      }else{
        aii <- .aii.pedicure
      }

      ## Loadings----
      ### Additive
      AddLambda <-matrix(sum.nice[[grep("vm", names(sum.nice))]], nrow = ntrials)[,-1]
      #### Scaled matrix
      AddLambda <- sqrt(aii)*AddLambda

      ### Non additive
      NonAddLambda <- matrix(sum.nice[[grep("ide",names(sum.nice))]], nrow = ntrials)[,-1]

      ### Total
      TotalLambda <- cbind(AddLambda, NonAddLambda)

      ## Rotated loadings
      ss <- svd(TotalLambda)
      TotalLambda <- -TotalLambda %*% ss$v # why the minus ?

      dimnames(TotalLambda) <- list(c(unique(dplyr::select(get(.asreml.obj$call$data),
                                                           trial_name)))[[1]],
                                    c(paste0("Add_Comp_",1:ncol(AddLambda)),
                                      paste0("NonAdd_Comp_",1:ncol(AddLambda))))

      ## Rotated Scores
      ### Additive
      AddRotScores <- matrix(.fa.asreml.obj[[2]][[add.comp.name]]$scores$blupr,
                             ncol = add.fa.order, byrow = F)
      AddScaledRotScores <- AddRotScores/sqrt(aii)

      dimnames(AddScaledRotScores) <- list(unique((.fa.asreml.obj[[2]][[add.comp.name]]$scores %>%
                                                     dplyr::select(., all_of(geno_name)))[,1]),
                                           unique((.fa.asreml.obj[[2]][[add.comp.name]]$scores %>%
                                                     dplyr::select(., trial_name))[,1]))
      AddScaledRotScores.df <- AddScaledRotScores %>% as.data.frame %>%
        tibble::rownames_to_column(geno_name)

      ### Non additive
      NonAddRotScores <- matrix(.fa.asreml.obj[[2]][[nonadd.comp.name]]$scores$blupr,
                                ncol = nonadd.fa.order, byrow = F)

      dimnames(NonAddRotScores) <- list(unique((.fa.asreml.obj[[2]][[nonadd.comp.name]]$scores %>%
                                                  dplyr::select(., all_of(geno_name)))[,1]),
                                        unique((.fa.asreml.obj[[2]][[nonadd.comp.name]]$scores %>%
                                                  dplyr::select(., trial_name))[,1]))
      NonAddRotScores.df <- NonAddRotScores %>% as.data.frame %>%
        tibble::rownames_to_column(geno_name)

      ### Total
      TotScores <- cbind(AddScaledRotScores, NonAddRotScores)

      # rotated scores
      TotRotScores <- -TotScores %*% ss$v

      dimnames(TotRotScores) <- list(unique((.fa.asreml.obj$blups[[1]]$scores %>%
                                               dplyr::select(., all_of(geno_name)))[,1]),
                                     c(paste0("Add_Comp_",1:ncol(AddLambda)),
                                       paste0("NonAdd_Comp_",1:ncol(AddLambda))))

      # Rotated Scores (data frame, rows = genotypes, columns = factors) to add the FAST measurements by genotype
      TotRotScores.df <- data.frame(Genotype = rownames(TotRotScores), TotRotScores)

      # Blups (data frame, only reg blups sorted by )
      totblups.df <- cbind(.fa.asreml.obj[[2]][[add.comp.name]]$blups[,c(2,3,1,4)],
                           .fa.asreml.obj[[2]][[nonadd.comp.name]]$blups[,c(1,4)])

      colnames(totblups.df)[3:4] <- paste0("Add_", colnames(totblups.df)[3:4] )
      colnames(totblups.df)[5:6] <- paste0("NonAdd_", colnames(totblups.df)[5:6] )

      totblups.df$Total_blup <- totblups.df$Add_blup + totblups.df$NonAdd_blup
      totblups.df$Total_regblup <- totblups.df$Add_regblup + totblups.df$NonAdd_regblup

      ## OP and RMSD (factor 1) ----
      #%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%-----
      #%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%-----
      # CHECK THIS CHECK THIS CHECK THIS CHECK THIS CHECK THIS CHECK THIS-----
      #%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%-----
      #%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%-----
      TotRotScores.df$OP <- mean(TotalLambda[,1]) * TotRotScores.df$Add_Comp_1 +
        mean(TotalLambda[,(1+ncol(AddLambda))]) * TotRotScores.df$NonAdd_Comp_1

      totblups.df$adjblup1 <- totblups.df$Total_regblup - as.vector(TotRotScores[,1] %*% t(TotalLambda[,1])) -
        as.vector(TotRotScores[,1] %*% t(TotalLambda[,(1+ncol(AddLambda))]))

      TotRotScores.df$RMSD <- as.vector(sqrt(tapply(totblups.df$adjblup1^2,
                                                    dplyr::select(totblups.df, geno_name)[,1],
                                                    mean)))

      TotRotScores.df <- TotRotScores.df[order(TotRotScores.df$OP,rev(TotRotScores.df$RMSD), decreasing = TRUE),]

      ## Responsiveness and adjusted blups (remaining factors) ----
      if (ncol(TotalLambda) > 1){
        for (i in 2:(ncol(TotalLambda))) {
          pos <- mean(TotalLambda[TotalLambda[, i] > 0, i])
          neg <- mean(TotalLambda[TotalLambda[, i] < 0, i])

          TotRotScores.df$RESP <- (pos-neg) * TotRotScores.df[, paste0("Comp", i)]
          names(TotRotScores.df)[names(TotRotScores.df) == "RESP"] <- paste0("RESP", i)
        }
        for (i in 2:(ncol(TotalLambda)-1)) {
          totblups.df$adjblup <- totblups.df$regblup - as.vector(TotRotScores[,1:i] %*% t(TotalLambda[,1:i]))
          names(totblups.df)[names(totblups.df) == "adjblup"] <- paste0("adjblup", i)
        }
      }

      # FAST measurements responsiveness and adjusted blups (remaining factors)
      if (ncol(TotalLambda) > 1){
        for (f in 2:ncol(TotalLambda)) {
          pos <- mean(TotalLambda[TotalLambda[, f] > 0, f])
          neg <- mean(TotalLambda[TotalLambda[, f] < 0, f])
          TotRotScores.df$RESP <- (pos-neg) * TotRotScores.df[, paste0("Comp_", f)]
          names(TotRotScores.df)[names(TotRotScores.df) == "RESP"] <- paste0("RESP", f)
        }
        for (f in 2:(ncol(TotalLambda)-1)) {
          totblups.df$adjblup <- totblups.df$regblup - as.vector(TotRotScores[,1:f] %*% t(TotalLambda[,1:f]))
          names(totblups.df)[names(totblups.df) == "adjblup"] <- paste0("adjblup", f)
        }
      }

      ## Print the results in a *.xlsx file ----
      .sheet = "FAST-total"
      # Create or load file
      if(file.exists(.output)){
        wb <- loadWorkbook(file = .output)
        if(!(.sheet %in% getSheetNames(.output))){
          addWorksheet(wb, sheetName = .sheet)
        }
      }else{
        wb <- createWorkbook(.output)
        addWorksheet(wb, sheetName = .sheet)
      }
      #### Title----
      writeData(wb,'Factor Analytic Selection Tools (FAST) Output', sheet = .sheet,
                startRow = 1, startCol = 1, colNames = F)
      # Title style
      addStyle(wb, sheet = .sheet, cols = 1L, rows = 1:2,
               style = createStyle(textDecoration = "bold", fontSize = 16))

      #### Table of results----
      writeData(wb, TotRotScores.df, sheet = .sheet, startRow = 2, startCol = 1,
                colNames = T, rowNames = F,
                headerStyle = createStyle(textDecoration = "bold",
                                          fontSize = 12,
                                          borderStyle = "medium",
                                          borderColour = "black",
                                          fgFill = "lightgrey"))

      # Set color scale for OP and RMSD measures
      conditionalFormatting(wb, sheet = .sheet,
                            cols = which(colnames(TotRotScores.df) == "OP"),
                            rows = 3:(3 + nrow(TotRotScores.df)),
                            style = c("#F8696B","white","#63BE7B"),
                            type = "colourScale")

      conditionalFormatting(wb, sheet = .sheet,
                            cols = which(colnames(TotRotScores.df) == "RMSD"),
                            rows = 3:(3 + nrow(TotRotScores.df)),
                            style = c("#63BE7B","white","#F8696B"),
                            type = "colourScale")
      # Set column width
      setColWidths(wb, cols = 1:ncol(TotRotScores.df), sheet = .sheet,
                   widths = "auto", ignoreMergedCells = TRUE)
      #### Save file----
      saveWorkbook(wb, file = .output, overwrite = TRUE)

      ### Save outputs in a list to return ----
      sum_list <- list()
      sum_list$FAST_total <- TotRotScores.df
      sum_list$Adj_total_BLUPs <- totblups.df
    }

    # Without genetic information
  }else if(knownG == FALSE){
    gen.comp.name <- names(sum.nice)[grep("fa", names(sum.nice))]
    gen.comp.str <- stringr::str_split_fixed(string = gen.comp.name,
                                             pattern = "[\\|,:,+.\\(,\\)]+", n = 2)[[1]][1]

    fa.order <- stringr::str_extract(gen.comp.name, "[0-9]+") %>% as.numeric

    trial_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+", string = gen.comp.name)[[1]][2]
    geno_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+", string = gen.comp.name)[[1]][4]

    ## Rotated loadings
    RotLambda <- .fa.asreml.obj[[1]][[1]]$'rotated loads' %>% as.data.frame %>%
      tibble::rownames_to_column(trial_name)

    ## Rotated Scores
    RotScores <- matrix(.fa.asreml.obj[[2]][[1]]$scores$blupr, ncol = fa.order, byrow = F)

    dimnames(RotScores) <- list(unique((.fa.asreml.obj[[2]][[1]]$scores %>%
                                          dplyr::select(., geno_name))[,1]),
                                unique((.fa.asreml.obj[[2]][[1]]$scores %>%
                                          dplyr::select(., all_of(trial_name)))[,1]))

    RotScores.df <- RotScores %>% as.data.frame %>%
      tibble::rownames_to_column(geno_name)

    ## OP and RMSD (factor 1) -----
    RotScores.df$OP <- mean(RotLambda[,2])*RotScores.df[,2]

    ## Blups (data frame, only reg blups)
    blups.df <- .fa.asreml.obj[[2]][[1]]$blups[,c(2,3,1,4)]

    blups.df$adjblup1 <- blups.df$regblup - as.vector(RotScores[,1] %*% t(RotLambda[,2]))

    RotScores.df$RMSD <- as.vector(sqrt(tapply(blups.df$adjblup1^2,
                                               dplyr::select(blups.df, geno_name)[,1],
                                               mean)))

    RotScores.df <- RotScores.df[order(RotScores.df$OP, rev(RotScores.df$RMSD),
                                       decreasing = TRUE),]


    # Responsiveness and adjusted BLUPs (remaining factors) ----
    if (fa.order == 2){
      pos <- mean(RotLambda[RotLambda[, 3] > 0, 3])
      neg <- mean(RotLambda[RotLambda[, 3] < 0, 3])

      RotScores.df$RESP <- (pos-neg) * RotScores.df[, paste0("Comp",2)]
      names(RotScores.df)[names(RotScores.df) == "RESP"] <- paste0("RESP", 2)

      blups.df$adjblup2 <- blups.df$regblup - as.vector(RotScores[,2] %*% t(RotLambda[,3]))

    } else if (fa.order > 2){
      for (i in 3:(fa.order + 1)) {
        pos <- mean(RotLambda[RotLambda[, i] > 0, i])
        neg <- mean(RotLambda[RotLambda[, i] < 0, i])

        RotScores.df$RESP <- (pos-neg) * RotScores.df[, paste0("Comp", i-1)]
        names(RotScores.df)[names(RotScores.df) == "RESP"] <- paste0("RESP", i-1)
      }
      for (i in 3:(fa.order + 1)) {
        blups.df$adjblup <- blups.df$regblup - as.vector(RotScores[,2:(i-1)] %*% t(RotLambda[,3:i]))
        names(blups.df)[names(blups.df) == "adjblup"] <- paste0("adjblup", i-1)
      }
    }

    blups.df <- merge(blups.df, RotLambda)

    # Print the results in a *.xlsx file ----
    .sheet = "FAST"
    # Create or load file
    if(file.exists(.output)){
      wb <- loadWorkbook(file = .output)
      if(!(.sheet %in% getSheetNames(.output))){
        addWorksheet(wb, sheetName = .sheet)
      }
    }else{
      wb <- createWorkbook(.output)
      addWorksheet(wb, sheetName = .sheet)
    }
    #### Title----
    writeData(wb,'Factor Analytic Selection Tools (FAST) Output', sheet = .sheet,
              startRow = 1, startCol = 1, colNames = F)
    # Title style
    addStyle(wb, sheet = .sheet, cols = 1L, rows = 1:2,
             style = createStyle(textDecoration = "bold", fontSize = 16))

    #### Table of results----
    writeData(wb, RotScores.df, sheet = .sheet, startRow = 2, startCol = 1,
              colNames = T, rowNames = F,
              headerStyle = createStyle(textDecoration = "bold",
                                        fontSize = 12,
                                        borderStyle = "medium",
                                        borderColour = "black",
                                        fgFill = "lightgrey"))

    # Set color scale for OP and RMSD measures
    conditionalFormatting(wb, sheet = .sheet,
                          cols = which(colnames(RotScores.df) == "OP"),
                          rows = 3:(3 + nrow(RotScores.df)),
                          style = c("#F8696B","white","#63BE7B"),
                          type = "colourScale")

    conditionalFormatting(wb, sheet = .sheet,
                          cols = which(colnames(RotScores.df) == "RMSD"),
                          rows = 3:(3 + nrow(RotScores.df)),
                          style = c("#63BE7B","white","#F8696B"),
                          type = "colourScale")
    # Set column width
    setColWidths(wb, cols = 1:ncol(RotScores.df), sheet = .sheet,
                 widths = "auto", ignoreMergedCells = TRUE)
    #### Save file----
    saveWorkbook(wb, file = .output, overwrite = TRUE)

    #### Save outputs in a list to return ----
    sum_list <- list()
    sum_list$FAST <- RotScores.df
    sum_list$Adj_BLUPs <- blups.df
  }
  return(sum_list)
}
