
#' Print summary overall results from a Multi-Environment Trial analysis in a workbook
#'
#' @description Prints a summary and overall results from a Multi-Environment Trial
#' analysis in a *.xlsx file using openxlsx. The required inputs depend on the model that was fitted and the
#' type of genetic information used, if any. The workbook is comprised of two worksheets. The first one contains
#' basic stats by trial, genotype concurrence by trial and a genotype by trial table. The second one contains
#' a model summary. The function also returns a list containing all the mentioned summaries.
#'
#' @param .asreml.obj An asreml object with the linear mixed model fitted for the analysis.
#' @param .fa.asreml.objectect A fa.asreml object from ASExtras4 package. Set as NULL by default. This
#' object is necesary only if a FA model was fitted
#' @param .aii.pedicure Set as NULL by default. If molecular marker data was used to compute the kinship
#' matrix during the analysis, the mean value of the diagonal elements of the kinship matrix prior to
#' the scaling must be introduced in this argument. The mean value of the diagonal elements of the kinship
#' matrix is used to compute the additive genetic variance.
#' @param .print If you wish to print the outputs in an excel file, please set as TRUE. Set as FALSE by default.
#' @param .output An *.xlsx file to write the outputs. The default is "MET_output.xlsx".
#'
#' @return A list containing
#' @return - $Trials_Summary a summary of the trials
#' @return - $Results_Summary a summary of the additive, non-additive and total genetic variance, error variance
#' @return - $Additive_Cmat and $Total_Cmat the additive and total genetic correlation matrices, if genetic correlation was accounted for
#' @return - $Additive_Heatmap and $Total_Heatmap the heatmaps for additive and total genetic correlation matrices, if genetic correlation was accounted for
#' @return - $Concurrence a genotype concurrence by trial table
#' @return - $GxE_tab a genotype by trial table
#' @return - $FA_sum the percentage of variance accounted for each trial by each
#' factor and in total if a FA model was fitted
#' @return A *.xlsx file comprised of two spreadsheet containing the information from the list
#'
#' @import openxlsx
#' @import ggplot2
#' @importFrom magrittr %>%
#' @importFrom stringr str_split_fixed
#' @import dplyr
#'
#' @export
#'

printMET <- function(.asreml.obj,
                      .fa.asreml.obj = NULL,
                      .aii.pedicure = NULL,
                      .print = FALSE,
                      .output = "MET_output.xlsx"){

  # Get the data from the model
  mydf <- get(.asreml.obj$call$data)
  sum.nice <- summary(.asreml.obj, vparameters = TRUE)$vparameters
  # 1. Get Trial/Environment and Genotype/Variety names
  if(any(grepl("vm", names(sum.nice)))){
    knownG <- TRUE
  } else{
    knownG <- FALSE
  }

  ## 1.1. knownG == T. Using genetic information
  if(knownG == TRUE){
    # Determine what structures were used for additive and non-additive:

    #   - diag():vm() + diag():ide()
    #   - us():vm() + diag():ide()
    #   - fa():vm() + diag():ide()
    #   - fa():vm() + fa():ide()

    ### 1.1.1. Additive component name
    add.comp.name <- stringr::str_split(as.character(.asreml.obj$call$random),
                                             pattern = "\\+")[[2]] %>%
      stringr::str_subset(., pattern = "vm") %>%
      trimws(., "both")

    add.comp.str <- stringr::str_split_fixed(string = add.comp.name,
                                             pattern = "\\(", n = 2)[[1]][1]
    ### 1.1.2. Non-additive component name
    nonadd.comp.name <- stringr::str_split(as.character(.asreml.obj$call$random),
                                           pattern = "\\+")[[2]] %>%
      stringr::str_subset(., pattern = "ide") %>%
      trimws(., "both")

    nonadd.comp.str <- stringr::str_split_fixed(string = nonadd.comp.name,
                                                pattern = "\\(", n = 2)[[1]][1]

    # Get names for Trial/Environment, Genotype/Variety and Relationship Matrix
    ### 1.1.3. GxE models
    #### 1.1.3.1. If the GxE model for additive part was diag or us
    if(add.comp.str != "fa" & add.comp.str != "us"){
      add.fa.order <- NULL
      relationship_mat.name <- (stringr::str_split(pattern = ":", string = add.comp.name)[[1]][2] %>%
                                  stringr::str_split(pattern = "[\\,:|,\\(,\\)]",.))[[1]][3] %>%
        trimws(., which = "both")

      # Get Trial and Genotype names from data
      trial_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+",
                                       string = add.comp.name)[[1]][2]

      geno_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+",
                                      string = add.comp.name)[[1]][4]
      add.comp.str <- "diag"
      genetic_model <- "Additive DIAG + Non-Additive DIAG"

    }else if(add.comp.str == "us"){
      trial_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+",
                                       string = add.comp.name)[[1]][2]
      geno_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+",
                                      string = add.comp.name)[[1]][4]

      relationship_mat.name <- (stringr::str_split(pattern = ":", string = add.comp.name)[[1]][2] %>%
                                  stringr::str_split(pattern = "[\\,:|,\\(,\\)]",.))[[1]][3] %>%
        trimws(., which = "both")
      genetic_model <- "Additive US + Non-Additive DIAG"
      #### 1.1.3.2. If the GxE model for additive part was fa
    }else if(add.comp.str == "fa"){

      add.fa.order <- stringr::str_extract(add.comp.name, "[0-9]+") %>% as.numeric

      relationship_mat.name <- (stringr::str_split(pattern = ":", string = add.comp.name)[[1]][2] %>%
                                   stringr::str_split(pattern = "[\\,:|,\\(,\\)]",.))[[1]][3] %>%
        trimws(., which = "both")

      # Get Trial and Genotype names from data
      trial_name <- stringr::str_split(pattern = "[\\,:|,+.\\(,\\)]+", string = add.comp.name)[[1]][2]

      geno_name <- stringr::str_split(pattern = "[\\,:|,+.\\(,\\)]+", string = add.comp.name)[[1]][5]

      #### 1.1.3.3. If the GxE model for non-additive part was fa
      if(nonadd.comp.str == "fa"){
        nonadd.fa.order <- stringr::str_extract(nonadd.comp.name, "[0-9]+") %>% as.numeric
        genetic_model <- paste0("Additive FA(", add.fa.order,
                                ") + Non-Additive FA(", nonadd.fa.order,
                                ")")
      }else{
        nonadd.comp.str <- "diag"
        genetic_model <- paste0("Additive FA(", add.fa.order,
                                ") + Non-Additive DIAG")
      }
    }
    ## 1.2. knownG = F. Without genetic information
  }else if(knownG == FALSE){

    # Determine what structures was used for genetic effects:

    #   - diag()
    #   - us()
    #   - fa()

    ### 1.2.1. GxE component name
    ## Here it is assumed that this term was the first term in the "random" part
    ## of the asreml model, as it is not possible to determine it using vm or
    ## ide

    gen.comp.name <- stringr::str_split(as.character(.asreml.obj$call$random),
                                                         pattern = "\\+")[[2]][1] %>%
      trimws(., "both")


    gen.comp.str <- stringr::str_split_fixed(string = gen.comp.name,
                                             pattern = "[\\|,:,+.\\(,\\)]+", n = 2)[[1]][1]

    # Get names for Trial/Environment, Genotype/Variety and Relationship Matrix
    ## Was the structure diag, us or fa?

    #### 1.2.2. If the GxE model was diag
    if(gen.comp.str != "us" & gen.comp.str != "fa"){
      gen.comp.str <- "diag"
      # Get Trial and Genotype names from data
      trial_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+", string = gen.comp.name)[[1]][2]

      geno_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+", string = gen.comp.name)[[1]][3]

      genetic_model <- "DIAG"
      #### 1.2.3. If the GxE model was us  ----
    }else if(gen.comp.str == "us"){
      # Get Trial and Genotype names from data
      trial_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+", string = gen.comp.name)[[1]][2]

      geno_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+", string = gen.comp.name)[[1]][3]
      genetic_model <- "US"
      #### 1.2.4. If the GxE model was fa  ----
    }else if(gen.comp.str == "fa"){
      fa.order <- stringr::str_extract(gen.comp.name, "[0-9]+") %>% as.numeric
      genetic_model <- paste0("FA(", fa.order, ")")

      trial_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+", string = gen.comp.name)[[1]][2]
      geno_name <- stringr::str_split(pattern = "[\\|,:,+.\\(,\\)]+", string = gen.comp.name)[[1]][4]

    }
  }
  # 2. Get check summaries
  ## 2.1. Concurrence table
  tt <- table(dplyr::select(mydf, all_of(geno_name), all_of(trial_name)))
  tt[tt>1] <- 1
  tt.conc <- t(tt)%*%tt %>% as.data.frame %>%
    tibble::rownames_to_column(all_of(trial_name))
  ## 2.2. GxE table
  ttgxe <- table(dplyr::select(mydf, all_of(geno_name), all_of(trial_name)))
  ttgxe[ttgxe>1] <- 1
  ttgxe <- as.data.frame.matrix(ttgxe)  %>%
    tibble::rownames_to_column(all_of(geno_name))

  ## 2.3. Trials summary table
  sumtab <- mydf %>% group_by(dplyr::select(mydf, all_of(trial_name))) %>%
    summarise(Rows = n_distinct(Row),
              Columms = n_distinct(Column),
              Reps = n_distinct(Replicate),
              Genos = n_distinct(get(geno_name)),
              Mean = round(mean(get(.asreml.obj$formulae$fixed[[2]]), na.rm = T),3),
              MVs = sum(is.na(get(.asreml.obj$formulae$fixed[[2]])))) %>%
    as.data.frame()

  ## 2.4. Print check summaries
  if(.print == TRUE){
    resultswb <- createWorkbook(.output)

    addWorksheet(resultswb, sheetName = 'FileChecks')
    row_pos <- 1
    writeData(resultswb,'Data file: ****COMPLETE****', sheet = 'FileChecks',
              startRow = row_pos, startCol = 1, colNames = F)

    addStyle(resultswb, sheet = 'FileChecks', cols = 1L, rows = row_pos,
             style = createStyle(textDecoration = "bold", fontSize = 16))
    row_pos <- row_pos + 1

    writeData(resultswb, 'Number of Trials:', sheet = 'FileChecks',
              startRow = row_pos, startCol = 1, colNames = F)

    writeData(resultswb, n_distinct(tt.conc$Trial),
              sheet = 'FileChecks',
              startRow = row_pos, startCol = 2, colNames = F)
    row_pos <- row_pos + 1

    writeData(resultswb, 'Number of Genotypes:', sheet = 'FileChecks',
              startRow = row_pos, startCol = 1, colNames = F)

    writeData(resultswb, n_distinct(ttgxe$Genotype),
              sheet = 'FileChecks',
              startRow = row_pos, startCol = 2, colNames = F)
    row_pos <- row_pos + 2

    writeData(resultswb, "Summary stats per Trial:", sheet = 'FileChecks',
              startRow = row_pos, startCol = 1, colNames = F)
    addStyle(resultswb, sheet = 'FileChecks', cols = 1L, rows = row_pos,
             style = createStyle(textDecoration = "bold", fontSize = 14))
    row_pos <- row_pos + 1

    writeData(resultswb, sumtab, sheet = 'FileChecks',
              startRow = row_pos, startCol = 1, colNames = T, rowNames = F,
              headerStyle = createStyle(textDecoration = "bold",
                                        fontSize = 12,
                                        borderStyle = "medium", borderColour = "black",
                                        fgFill = "lightgrey"))
    addStyle(resultswb, sheet = 'FileChecks', cols = 1L, rows = (row_pos):(row_pos + nrow(sumtab)),
             style = createStyle(textDecoration = "bold",
                                 fontSize = 12,
                                 borderStyle = "medium", borderColour = "black",
                                 fgFill = "lightgrey"),
             gridExpand = T)

    row_pos <- row_pos + nrow(sumtab) + 2

    writeData(resultswb, "Genotype concurrences by Trial:", sheet = 'FileChecks',
              startRow = row_pos, startCol = 1, colNames = F)
    addStyle(resultswb, sheet = 'FileChecks', cols = 1L, rows = row_pos,
             style = createStyle(textDecoration = "bold", fontSize = 14))
    row_pos <- row_pos + 1

    writeData(resultswb, tt.conc, sheet = 'FileChecks',
              startRow = row_pos, startCol = 1, colNames = T, rowNames = F,
              headerStyle = createStyle(textDecoration = "bold",
                                        fontSize = 12,
                                        borderStyle = "medium", borderColour = "black",
                                        fgFill = "lightgrey"))

    addStyle(resultswb, sheet = 'FileChecks', cols = 1L, rows = (row_pos):(row_pos + nrow(tt.conc)),
             style = createStyle(textDecoration = "bold",
                                 fontSize = 12,
                                 borderStyle = "medium", borderColour = "black",
                                 fgFill = "lightgrey"),
             gridExpand = T)

    row_pos <- row_pos + nrow(tt.conc) + 2

    writeData(resultswb, "Genotype by Trial table:", sheet = 'FileChecks',
              startRow = row_pos, startCol = 1, colNames = F)
    addStyle(resultswb, sheet = 'FileChecks', cols = 1L, rows = row_pos,
             style = createStyle(textDecoration = "bold", fontSize = 14))
    row_pos <- row_pos + 1

    writeData(resultswb, ttgxe, sheet = 'FileChecks',
              startRow = row_pos, startCol = 1, colNames = T, rowNames = F,
              headerStyle = createStyle(textDecoration = "bold",
                                        fontSize = 12,
                                        borderStyle = "medium", borderColour = "black",
                                        fgFill = "lightgrey"))
    addStyle(resultswb, sheet = 'FileChecks', cols = 1L, rows = (row_pos):(row_pos + nrow(ttgxe)),
             style = createStyle(textDecoration = "bold",
                                 fontSize = 12,
                                 borderStyle = "medium", borderColour = "black",
                                 fgFill = "lightgrey"),
             gridExpand = T)

    setColWidths(resultswb, sheet = 'FileChecks',
                 cols = 1:(ncol(ttgxe)), widths = 18)
    setColWidths(resultswb, sheet = 'FileChecks',
                 cols = (ncol(ttgxe)+1):7, widths = 18)
  }

  # 3. Get model summaries
  ## 3.1. Variance Components
  ### 3.1.1. knownG = T. Using genetic information
  if(knownG == TRUE){
    #### 3.1.1.1. If the GxE model for additive part was fa
    if(add.comp.str == "fa" & nonadd.comp.str == "diag"){

      vaf.all <- round(.fa.asreml.obj[[1]][[1]]$'total',1)
      pct.expl <- round(.fa.asreml.obj[[1]][[1]]$'site %vaf',1)

      pct.expl <- as.data.frame(cbind(rownames(pct.expl),pct.expl))
      rownames(pct.expl) <- NULL

      for(i in 2:(ncol(pct.expl)-1)){
        colnames(pct.expl)[i] <- paste("Factor",i-1)
      }
      colnames(pct.expl)[1] <- trial_name
      colnames(pct.expl)[ncol(pct.expl)] <- "All"
      pct.expl[,-1] <- sapply(pct.expl[-1],as.numeric)


      add.rot.loads <- round(.fa.asreml.obj$gammas[[add.comp.name]]$'rotated loads', 3)
      add.pct.vaf <- round(.fa.asreml.obj$gammas[[add.comp.name]]$'site %vaf', 2)
      colnames(add.pct.vaf) <- paste('vaf', colnames(add.pct.vaf), sep = '_')


      # To determine the additive Gmat, determine if the genetic information comes
      # from pedigree or marker data:
      if("inbreeding" %in% names(attributes(get(relationship_mat.name)))){
        aii <-  mean(1 + attr(get(relationship_mat.name),"inbreeding"))
        add.Gmat <- aii*.fa.asreml.obj$gammas[[add.comp.name]]$Gmat
      }else{
        aii <- .aii.pedicure
        add.Gmat <- aii*.fa.asreml.obj$gammas[[add.comp.name]]$Gmat

      }
      #### 3.1.1.2. If the GxE model for additive part was us ----
    }else if(add.comp.str == "us"){
      add.comp.name <- names(sum.nice) %>%
        stringr::str_subset(., pattern = "vm")
      if("inbreeding" %in% names(attributes(get(relationship_mat.name)))){
        aii <-  mean(1 + attr(get(relationship_mat.name),"inbreeding"))
        add.Gmat <- aii*sum.nice[[add.comp.name]]
      }else{
        aii <- .aii.pedicure
        add.Gmat <- aii*sum.nice[[add.comp.name]]
      }
      #### 3.1.1.3. If the GxE model for additive part was diag ----
    }else if(add.comp.str == "diag"){
      add.comp.name <- names(sum.nice) %>%
        stringr::str_subset(., pattern = "vm")
      if("inbreeding" %in% names(attributes(get(relationship_mat.name)))){
        aii <-  mean(1 + attr(get(relationship_mat.name),"inbreeding"))
        add.Gmat <- aii*sum.nice[[add.comp.name]]
      }else{
        aii <- .aii.pedicure
        add.Gmat <- aii*sum.nice[[add.comp.name]]
      }
      add.Gmat <- diag(add.Gmat)
      temp <- names(sum.nice[[add.comp.name]])
      temp <- str_split_fixed(temp, "_", 2)[,2]
      dimnames(add.Gmat) <- list(temp,temp)
    }

    #no fa order for nonadd
    # aux2 <- paste0(paste0("fa(Trial, ", nonadd.fa.order), "):ide(Genotype)")
    #s3mung19.nonadd.rot.loads <- round(mets2mung21.met.sum$gammas[[relationship_mat.name]]$'rotated loads',3)
    #s3mung19.nonadd.pct.vaf <- round(mets2mung21.met.sum$gammas[[relationship_mat.name]]$'site %vaf',2)
    #colnames(s3mung19.nonadd.pct.vaf)=paste('vaf',colnames(s3mung19.nonadd.pct.vaf),sep='_')
    nonadd.comp.name <- names(sum.nice) %>%
      stringr::str_subset(., pattern = "ide")
    nonadd.Gmat <- diag(sum.nice[names(sum.nice) %in% nonadd.comp.name][[1]])
    #nonadd.Gmat <- mets2mung21.met.sum$gammas[[relationship_mat.name]]$Gmat

    tot.Gmat <- add.Gmat + nonadd.Gmat

    evar <- round(.asreml.obj$vparameters[grep("!R$", names(.asreml.obj$vparameters))],4)

    Trial.sum <- cbind(diag(add.Gmat), # additive genetic var
                       diag(nonadd.Gmat), # non-additive genetic var
                       diag(tot.Gmat), # total genetic var
                       evar) %>% # residual var
      round(.,4) %>% cbind(sumtab$Mean,.) %>%
      as.data.frame %>%
      setNames(., c('Mean', 'G.Add', 'G.NonAdd', 'G.Total', 'Error')) %>%
      mutate(`Additive pct.` = round(100*G.Add/G.Total,2),
             `Non-Additive pct.` = round(100*G.NonAdd/G.Total,2)) %>%
      tibble::rownames_to_column(trial_name)

    Trial.sum[,-1] <- sapply(Trial.sum[-1],as.numeric)

    if(add.comp.str != "diag"){
      add.Cmat <- round(cov2cor(add.Gmat),3)
      add.Cmat <- add.Cmat %>% as.data.frame %>%
        tibble::rownames_to_column(trial_name)
      add.Cmat[,-1] <- sapply(add.Cmat[-1], as.numeric)

      tot.Cmat <- round(cov2cor(tot.Gmat),3)
      tot.Cmat <- tot.Cmat %>% as.data.frame %>%
        tibble::rownames_to_column(trial_name)
      tot.Cmat[,-1] <- sapply(tot.Cmat[-1], as.numeric)

      # Heatmap with ggplot

      dis.mat <- 1 - add.Cmat[,-1]
      agnes.add <- cluster::agnes(x = dis.mat, diss = T)

      Cmat_long <- reshape2::melt(data = add.Cmat, id.vars = trial_name)
      colnames(Cmat_long) <- c("Var1","Var2","value")

      levels <- add.Cmat[,1]
      levels <- as.factor(levels)
      levels <- levels[agnes.add$order]

      # Refacotrise environment levels according to the ordered levels
      Cmat_long$Var1 <- factor(Cmat_long$Var1, levels = levels)
      Cmat_long$Var2 <- factor(Cmat_long$Var2, levels = levels)
      # Reverse order of factor levels on the y-axis
      Cmat_long$Var2 <- fct_rev(Cmat_long$Var2)

      # Produce heatmap using ggplot2 (MM's)

      #The colour scale below is the same one used in fa.asreml
      hh <- rev(rainbow(256, start = 0, end = 2/3))
      # Set diagonal value equal to NA
      Cmat_long$value[Cmat_long$value == 1] <- NA

      # Heatmap with fa.asreml color scale
      add.heatmap <- ggplot(Cmat_long , aes(Var1, Var2)) +
        geom_tile(aes(fill = value)) +
        xlab(trial_name) +
        ylab(trial_name) +
        guides(fill = guide_colourbar(barwidth = 1, barheight = 18)) +
        theme_bw() +
        theme(panel.grid.minor = element_blank(),
              axis.line.x = element_line(colour = "black"),
              legend.title = element_text(size = 17),
              legend.text = element_text(size = 15),
              axis.line.y = element_line(colour = "black"),
              axis.text.x = element_text(size = 12, angle = 90, vjust = 0.5, hjust = 1),
              axis.title.y = element_text(size = 20),
              axis.title.x = element_text(size = 20),
              axis.text.y = element_text(size = 12),
              strip.text.x = element_text(size = 15),
              strip.text.y = element_text(size = 15)) +
        scale_fill_gradientn(colours = hh, na.value='white',
                             name = "Additive\nGenetic\ncorrelation", limits = c(-1,1))

      dis.mat <- 1 - tot.Cmat[,-1]
      agnes.tot <- cluster::agnes(x = dis.mat, diss = T)

      # Heatmap with ggplot
      Cmat_long <- reshape2::melt(data = tot.Cmat, id.vars = "Trial")
      colnames(Cmat_long) <- c("Var1","Var2","value")

      levels <- tot.Cmat[,1]
      levels <- as.factor(levels)
      levels <- levels[agnes.tot$order]

      # Refacotrise environment levels according to the ordered levels
      Cmat_long$Var1 <- factor(Cmat_long$Var1, levels = levels)
      Cmat_long$Var2 <- factor(Cmat_long$Var2, levels = levels)
      # Reverse order of factor levels on the y-axis
      Cmat_long$Var2 <- fct_rev(Cmat_long$Var2)

      # Produce heatmap using ggplot2 (MM's)

      #The colour scale below is the same one used in fa.asreml
      hh <- rev(rainbow(256, start = 0, end = 2/3))
      # Set diagonal value equal to NA
      Cmat_long$value[Cmat_long$value == 1] <- NA

      # Heatmap with fa.asreml color scale
      tot.heatmap <- ggplot(Cmat_long , aes(Var1, Var2)) +
        geom_tile(aes(fill = value)) +
        xlab(trial_name) +
        ylab(trial_name) +
        guides(fill = guide_colourbar(barwidth = 1, barheight = 18)) +
        theme_bw() +
        theme(panel.grid.minor = element_blank(),
              axis.line.x = element_line(colour = "black"),
              legend.title = element_text(size = 17),
              legend.text = element_text(size = 15),
              axis.line.y = element_line(colour = "black"),
              axis.text.x = element_text(size = 12, angle = 90, vjust = 0.5, hjust = 1),
              axis.title.y = element_text(size = 20),
              axis.title.x = element_text(size = 20),
              axis.text.y = element_text(size = 12),
              strip.text.x = element_text(size = 15),
              strip.text.y = element_text(size = 15)) +
        scale_fill_gradientn(colours = hh, na.value='white',
                             name = "Total\nGenetic\ncorrelation", limits = c(-1,1))
    }
    ### 3.1.2. knownG = F. Without genetic information
  } else if(knownG == FALSE){
    #### 3.1.2.1. If the GxE model  was fa ----
    if(gen.comp.str == "fa"){

      vaf.all <- round(.fa.asreml.obj[[1]][[1]]$'total',1)
      pct.expl <- round(.fa.asreml.obj[[1]][[1]]$'site %vaf',1)

      pct.expl <- as.data.frame(cbind(rownames(pct.expl), pct.expl))
      rownames(pct.expl) <- NULL

      for(i in 2:(ncol(pct.expl)-1)){
        colnames(pct.expl)[i] <- paste("Factor",i-1)
      }
      colnames(pct.expl)[1] <- trial_name
      colnames(pct.expl)[ncol(pct.expl)] <- "All"
      pct.expl[,-1] <- sapply(pct.expl[-1],as.numeric)


      rot.loads <- round(.fa.asreml.obj$gammas[[gen.comp.name]]$'rotated loads', 3)
      pct.vaf <- round(.fa.asreml.obj$gammas[[gen.comp.name]]$'site %vaf', 2)
      colnames(pct.vaf) <- paste('vaf', colnames(pct.vaf), sep = '_')

      Gmat <- .fa.asreml.obj$gammas[[gen.comp.name]]$Gmat

      #### 3.1.2.2. If the GxE model was us
    }else if(gen.comp.str == "us"){
      Gmat <- sum.nice[[gen.comp.name]]


      #### 3.1.2.3. If the GxE model was diag
    }else if(gen.comp.str == "diag"){
      gen.comp.name <- names(sum.nice)[1]
      Gmat <- diag(sum.nice[[gen.comp.name]])

      temp <- names(sum.nice[[gen.comp.name]])
      temp <- str_split_fixed(temp, "_", 2)[,2]
      dimnames(Gmat) <- list(temp,temp)
    }

    evar <- round(.asreml.obj$vparameters[grep("!R$", names(.asreml.obj$vparameters))],4)

    Trial.sum <- cbind(diag(Gmat), # genetic var
                       evar) %>% # residual var
      round(.,4) %>% cbind(sumtab$Mean,.) %>%
      as.data.frame %>%
      setNames(., c('Mean', 'G.Var', 'Error')) %>%
      tibble::rownames_to_column(trial_name)

    Trial.sum[,-1] <- sapply(Trial.sum[-1],as.numeric)

    if(gen.comp.str != "diag"){
      Cmat <- round(cov2cor(Gmat),3)
      Cmat <- Cmat %>% as.data.frame %>%
        tibble::rownames_to_column(trial_name)
      Cmat[,-1] <- sapply(Cmat[-1], as.numeric)

      # Heatmap with ggplot

      dis.mat <- 1 - Cmat[,-1]
      agnes.gen <- cluster::agnes(x = dis.mat, diss = T)

      Cmat_long <- reshape2::melt(data = Cmat, id.vars = trial_name)
      colnames(Cmat_long) <- c("Var1","Var2","value")

      levels <- Cmat[,1]
      levels <- as.factor(levels)
      levels <- levels[agnes.gen$order]

      # Refacotrise environment levels according to the ordered levels
      Cmat_long$Var1 <- factor(Cmat_long$Var1, levels = levels)
      Cmat_long$Var2 <- factor(Cmat_long$Var2, levels = levels)
      # Reverse order of factor levels on the y-axis
      Cmat_long$Var2 <- fct_rev(Cmat_long$Var2)

      # Produce heatmap using ggplot2 (MM's)

      #The colour scale below is the same one used in fa.asreml
      hh <- rev(rainbow(256, start = 0, end = 2/3))
      # Set diagonal value equal to NA
      Cmat_long$value[Cmat_long$value == 1] <- NA

      # Heatmap with fa.asreml color scale
      heatmap <- ggplot(Cmat_long , aes(Var1, Var2)) +
        geom_tile(aes(fill = value)) +
        xlab(trial_name) +
        ylab(trial_name) +
        guides(fill = guide_colourbar(barwidth = 1, barheight = 18)) +
        theme_bw() +
        theme(panel.grid.minor = element_blank(),
              axis.line.x = element_line(colour = "black"),
              legend.title = element_text(size = 17),
              legend.text = element_text(size = 15),
              axis.line.y = element_line(colour = "black"),
              axis.text.x = element_text(size = 12, angle = 90, vjust = 0.5, hjust = 1),
              axis.title.y = element_text(size = 20),
              axis.title.x = element_text(size = 20),
              axis.text.y = element_text(size = 12),
              strip.text.x = element_text(size = 15),
              strip.text.y = element_text(size = 15)) +
        scale_fill_gradientn(colours = hh, na.value='white',
                             name = "Genetic\ncorrelation", limits = c(-1,1))
    }
  }

  # 4. Print model summaries
  if(knownG == TRUE){

    if(.print == TRUE){
      addWorksheet(resultswb, sheetName = 'ModelSummary')
      row_pos <- 1

      writeData(resultswb, paste0("GxE model for yield: ", genetic_model),
                sheet = 'ModelSummary',
                startRow = row_pos, startCol = 1, colNames = F)
      addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = row_pos,
               style = createStyle(textDecoration = "bold", fontSize = 16))
      if(add.comp.str == "fa"){
        row_pos <- row_pos + 1
        writeData(resultswb, paste0("The percentage of additive genetic variance accounted for by the factorial model was ", vaf.all,"%"),
                  sheet = 'ModelSummary',
                  startRow = row_pos, startCol = 1, colNames = F)

        addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = row_pos,
                 style = createStyle(textDecoration = "bold"))
        row_pos <- row_pos + 2

        writeData(resultswb, "Percentage of addititve genetic variance explained by factors at each Trial:",
                  sheet = 'ModelSummary',
                  startRow = row_pos, startCol = 1, colNames = F)
        addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = row_pos,
                 style = createStyle(textDecoration = "bold", fontSize = 14))
        row_pos <- row_pos + 1

        writeData(resultswb, pct.expl,
                  sheet = 'ModelSummary',
                  startRow = row_pos, startCol = 1, colNames = T,
                  headerStyle = createStyle(textDecoration = "bold",
                                            fontSize = 12,
                                            borderStyle = "medium",
                                            borderColour = "black",
                                            fgFill = "lightgrey"))

        addStyle(resultswb, sheet = 'ModelSummary', cols = 1L,
                 rows = (row_pos):(row_pos + nrow(pct.expl)),
                 style = createStyle(textDecoration = "bold",
                                     fontSize = 12,
                                     borderStyle = "medium",
                                     borderColour = "black",
                                     fgFill = "lightgrey"), gridExpand = TRUE)
        row_pos <- row_pos + nrow(pct.expl) + 2
      }

      writeData(resultswb, "Trial information: Mean, Additive, non-additive and total genetic and error variances:",
                sheet = 'ModelSummary',
                startRow = row_pos, startCol = 1, colNames = F)
      addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = row_pos,
               style = createStyle(textDecoration = "bold", fontSize = 14))

      row_pos <- row_pos + 1
      writeData(resultswb, Trial.sum, sheet = 'ModelSummary',
                startRow = row_pos, startCol = 1, colNames = T, rowNames = F,
                headerStyle = createStyle(textDecoration = "bold",
                                          fontSize = 12,
                                          borderStyle = "medium",
                                          borderColour = "black",
                                          fgFill = "lightgrey"))

      addStyle(resultswb, sheet = 'ModelSummary', cols = 1L,
               rows = (row_pos):(row_pos + nrow(Trial.sum)),
               style = createStyle(textDecoration = "bold",
                                   fontSize = 12,
                                   borderStyle = "medium",
                                   borderColour = "black",
                                   fgFill = "lightgrey"), gridExpand = TRUE)
      row_pos <- row_pos + nrow(Trial.sum) + 2

      if(add.comp.str != "diag"){
        writeData(resultswb, "Additive Genetic correlations matrix:",
                  sheet = 'ModelSummary',
                  startRow = row_pos, startCol = 1, colNames = F)
        addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = row_pos,
                 style = createStyle(textDecoration = "bold", fontSize = 14))
        row_pos <- row_pos + 1

        writeData(resultswb, add.Cmat, sheet = 'ModelSummary',
                  startRow = row_pos, startCol = 1, colNames = T, rowNames = F,
                  headerStyle = createStyle(textDecoration = "bold",
                                            fontSize = 12,
                                            borderStyle = "medium", borderColour = "black",
                                            fgFill = "lightgrey"))
        addStyle(resultswb, sheet = 'ModelSummary', cols = 1L,
                 rows = (row_pos):(row_pos + nrow(add.Cmat)),
                 style = createStyle(textDecoration = "bold",
                                     fontSize = 12,
                                     borderStyle = "medium",
                                     borderColour = "black",
                                     fgFill = "lightgrey"), gridExpand = TRUE)

        row_pos <- row_pos + nrow(add.Cmat) + 2

        writeData(resultswb, "Total Genetic correlations matrix:",
                  sheet = 'ModelSummary',
                  startRow = row_pos, startCol = 1, colNames = F)
        addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = row_pos,
                 style = createStyle(textDecoration = "bold", fontSize = 14))
        row_pos <- row_pos + 1

        writeData(resultswb, tot.Cmat,
                  sheet = 'ModelSummary', startRow = row_pos, startCol = 1,
                  colNames = T, rowNames = F,
                  headerStyle = createStyle(textDecoration = "bold",
                                            fontSize = 12,
                                            borderStyle = "medium",
                                            borderColour = "black",
                                            fgFill = "lightgrey"))
        addStyle(resultswb, sheet = 'ModelSummary', cols = 1L,
                 rows = (row_pos):(row_pos + nrow(tot.Cmat)),
                 style = createStyle(textDecoration = "bold",
                                     fontSize = 12,
                                     borderStyle = "medium",
                                     borderColour = "black",
                                     fgFill = "lightgrey"), gridExpand = TRUE)

      }


      if(add.comp.str != "diag"){
        setColWidths(resultswb, sheet = 'ModelSummary',
                     cols = 1:(ncol(add.Cmat)+20), widths = 18)

      }else if(add.comp.str == "diag"){
        setColWidths(resultswb, sheet = 'ModelSummary',
                     cols = 1:8, widths = 18)
      }

      saveWorkbook(resultswb, file = .output, overwrite = TRUE)
    }
    sum_list <- list()
    sum_list$Trials_Summary <- sumtab
    sum_list$Results_Summary <- Trial.sum
    if(add.comp.str != "diag"){
      sum_list$Additive_Cmat <- add.Cmat
      sum_list$Total_Cmat <- tot.Cmat
      sum_list$Additive_Heatmap <- add.heatmap
      sum_list$Total_Heatmap <- tot.heatmap
    }
    sum_list$Concurrence <- tt.conc
    sum_list$GxE_tab <- ttgxe
    if(add.comp.str == "fa"){
      sum_list$FA_sum <- pct.expl
    }

  }else if(knownG == FALSE){

    if(.print == TRUE){
      addWorksheet(resultswb, sheetName = 'ModelSummary')
      row_pos <- 1

      writeData(resultswb, paste0("GxE model for yield: ", genetic_model),
                sheet = 'ModelSummary',
                startRow = row_pos, startCol = 1, colNames = F)
      addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = row_pos,
               style = createStyle(textDecoration = "bold", fontSize = 16))
      if(gen.comp.str == "fa"){
        row_pos <- row_pos + 1
        writeData(resultswb, paste0("The percentage of genetic variance accounted for by the factorial model was ", vaf.all,"%"),
                  sheet = 'ModelSummary',
                  startRow = row_pos, startCol = 1, colNames = F)

        addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = row_pos,
                 style = createStyle(textDecoration = "bold"))
        row_pos <- row_pos + 2

        writeData(resultswb, "Percentage of genetic variance explained by factors at each Trial:",
                  sheet = 'ModelSummary',
                  startRow = row_pos, startCol = 1, colNames = F)
        addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = row_pos,
                 style = createStyle(textDecoration = "bold", fontSize = 14))
        row_pos <- row_pos + 1

        writeData(resultswb, pct.expl,
                  sheet = 'ModelSummary',
                  startRow = row_pos, startCol = 1, colNames = T,
                  headerStyle = createStyle(textDecoration = "bold",
                                            fontSize = 12,
                                            borderStyle = "medium",
                                            borderColour = "black",
                                            fgFill = "lightgrey"))

        addStyle(resultswb, sheet = 'ModelSummary', cols = 1L,
                 rows = (row_pos):(row_pos + nrow(pct.expl)),
                 style = createStyle(textDecoration = "bold",
                                     fontSize = 12,
                                     borderStyle = "medium",
                                     borderColour = "black",
                                     fgFill = "lightgrey"), gridExpand = TRUE)
        row_pos <- row_pos + nrow(pct.expl) + 2
      }

      writeData(resultswb, "Trial information: Mean, genetic and error variances:",
                sheet = 'ModelSummary',
                startRow = row_pos, startCol = 1, colNames = F)
      addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = row_pos,
               style = createStyle(textDecoration = "bold", fontSize = 14))

      row_pos <- row_pos + 1
      writeData(resultswb, Trial.sum, sheet = 'ModelSummary',
                startRow = row_pos, startCol = 1, colNames = T, rowNames = F,
                headerStyle = createStyle(textDecoration = "bold",
                                          fontSize = 12,
                                          borderStyle = "medium",
                                          borderColour = "black",
                                          fgFill = "lightgrey"))

      addStyle(resultswb, sheet = 'ModelSummary', cols = 1L,
               rows = (row_pos):(row_pos + nrow(Trial.sum)),
               style = createStyle(textDecoration = "bold",
                                   fontSize = 12,
                                   borderStyle = "medium",
                                   borderColour = "black",
                                   fgFill = "lightgrey"), gridExpand = TRUE)
      row_pos <- row_pos + nrow(Trial.sum) + 2

      if(gen.comp.str != "diag"){
        writeData(resultswb, "Genetic correlation matrix:",
                  sheet = 'ModelSummary',
                  startRow = row_pos, startCol = 1, colNames = F)
        addStyle(resultswb, sheet = 'ModelSummary', cols = 1L, rows = row_pos,
                 style = createStyle(textDecoration = "bold", fontSize = 14))
        row_pos <- row_pos + 1

        writeData(resultswb, Cmat, sheet = 'ModelSummary',
                  startRow = row_pos, startCol = 1, colNames = T, rowNames = F,
                  headerStyle = createStyle(textDecoration = "bold",
                                            fontSize = 12,
                                            borderStyle = "medium", borderColour = "black",
                                            fgFill = "lightgrey"))
        addStyle(resultswb, sheet = 'ModelSummary', cols = 1L,
                 rows = (row_pos):(row_pos + nrow(Cmat)),
                 style = createStyle(textDecoration = "bold",
                                     fontSize = 12,
                                     borderStyle = "medium",
                                     borderColour = "black",
                                     fgFill = "lightgrey"), gridExpand = TRUE)

        row_pos <- row_pos + nrow(Cmat) + 2
      }


      if(gen.comp.str != "diag"){
        setColWidths(resultswb, sheet = 'ModelSummary',
                     cols = 1:(ncol(Cmat)+20), widths = 18)

      }else if(gen.comp.str == "diag"){
        setColWidths(resultswb, sheet = 'ModelSummary',
                     cols = 1:8, widths = 18)
      }

      saveWorkbook(resultswb, file = .output, overwrite = TRUE)
    }
    sum_list <- list()
    sum_list$Trials_Summary <- sumtab
    sum_list$Results_Summary <- Trial.sum
    if(gen.comp.str != "diag"){
      sum_list$Cmat <- Cmat
      sum_list$Heatmap <- heatmap
    }
    sum_list$Concurrence <- tt.conc
    sum_list$GxE_tab <- ttgxe
    if(gen.comp.str == "fa"){
      sum_list$FA_sum <- pct.expl
    }
  }
  return(sum_list)
}
