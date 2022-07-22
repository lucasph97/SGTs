
#' Summary of data set
#'
#' @description Aims to summarise trials from a data set. This function considers
#' simple breeding trial designs, namely, rectangular arrays with replicate blocks and genotypes randomised
#' within replicates. Structural terms must be named as "Replicate", "Row" and "Column", and the genotype
#' or variety factor as "Genotype". This function also provides plots of the trial layout
#' coloured by genotype levels.
#'
#' @param .df A dataframe object containing the data to be analysed.
#' @param .trait A vector containing the index or names from the traits that you are interested in
#' summarising. Traits must be set as numeric. The function provides the mean, standard deviation and
#' missing values.
#' @param .trial Set as "Trial" by default. If multiple trials are considered, please enter the name of the trial factor.
#'
#' @return Returns a list comprised of
#' @return - a dataframe with a summary of trial structure and traits selected
#' @return - if multiple trials, a list with a ggplot object per trial. If a single trial, a ggplot object
#'
#' @importFrom magrittr %>%
#' @import ggplot2
#' @import dplyr
#' @importFrom forcats fct_rev
#' @export
#'

sumTrial <- function(.df, .trait = NULL, .trial = "Trial"){

  if(mean(grepl("^-?[0-9.]+$", .trait)) == 1){
    .trait <- colnames(.df)[.trait]
  } else if(mean(grepl("^-?[0-9.]+$", .trait)) == 0){
    .trait <- .trait
  }

  if(is.null(.trial)){

    cols <- c("Replicate" ,"Column", "Row", "Genotype",.trait)

    .df <- .df %>% select(c(all_of(cols)))

    na_sum <- function(x){sum(is.na(x))}

    .df_sum <- .df %>% summarise(across(.cols = c("Replicate" ,"Column", "Row", "Genotype"),
                                        .fns = n_distinct),
                                 across(.cols = c(.trait),
                                        .fns = list(Mn = mean, SD = sd), na.rm = T),
                                 across(.cols = c(.trait),
                                        .fns = list(NAs = na_sum))) %>%
      as.data.frame()
    colnames(.df_sum)[1:4] <- c("Reps","Columns","Rows","Genotypes")

    # Plot of trial layout
    .df$Row <- factor(.df$Row,
                      levels =  sort(as.numeric(levels(.df$Row))))
    .df$Column <- factor(.df$Column,
                         levels =  sort(as.numeric(levels(.df$Column))))

    # To get replicate borders (snippet of code provided by Alec Zwart)
    repBorders <- .df %>%
      group_by(Replicate) %>%
      summarise(rowmin = min(as.integer(Row)) - 0.5, rowmax = max(as.integer(Row)) + 0.5,
                colmin = min(as.integer(Column)) - 0.5, colmax = max(as.integer(Column)) + 0.5)



    layout <- ggplot(data = .df) +
      geom_tile(aes(x = Row,
                    y = fct_rev(as.factor(Column)),
                    fill = Genotype)) +
      geom_text(aes(x = Row,
                    y = fct_rev(as.factor(Column)),
                    label = Replicate)) +
      geom_rect(aes(xmin = rowmin, xmax = rowmax,
                    ymin = colmin, ymax = colmax),
                data = repBorders,
                colour = "black", fill = NA, size = 1.5) +
      ylab("Column") + xlab("Row") +
      theme_bw() + theme(legend.position = "none")

    sum_list <- list()
    sum_list[[1]] <- .df_sum
    sum_list[[2]] <- layout

    names(sum_list)[1] <- "Trial summary"
    names(sum_list)[2] <- "Trial layout"

  } else {
    cols <- c(.trial, "Replicate", "Row", "Column", "Genotype",.trait)

    .df <- .df %>% select(c(all_of(cols)))

    na_sum <- function(x){sum(is.na(x))}

    .df_sum <- .df %>% group_by(Trial = get(.trial)) %>%
      summarise(across(.cols = c("Replicate" ,"Column", "Row", "Genotype"),
                                        .fns = n_distinct),
                                 across(.cols = c(all_of(.trait)),
                                        .fns = list(Mn = mean, SD = sd), na.rm = T),
                                 across(.cols = c(.trait),
                                        .fns = list(NAs = na_sum))) %>%
      as.data.frame()

    colnames(.df_sum)[1:5] <- c("Trial","Reps","Columns","Rows","Genotypes")

    # Trial layouts

    layouts <- list()

    for (i in 1:n_distinct(.df_sum$Trial)){
      mytrial <- as.character(.df_sum$Trial[i])

      mydata <- droplevels(subset(.df, .df$Trial == mytrial))

      mydata$Row <- factor(mydata$Row,
                           levels =  sort(as.numeric(levels(mydata$Row))))
      mydata$Column <- factor(mydata$Column,
                           levels =  sort(as.numeric(levels(mydata$Column))))

      # To get replicate borders (snippet of code provided by Alec Zwart)
      repBorders <- mydata %>%
        group_by(Replicate) %>%
        summarise(rowmin = min(as.integer(Row)) - 0.5, rowmax = max(as.integer(Row)) + 0.5,
                  colmin = min(as.integer(Column)) - 0.5, colmax = max(as.integer(Column)) + 0.5)


      layouts[[mytrial]] <- ggplot(data = mydata) +
        geom_tile(aes(x = Row,
                      y = fct_rev(as.factor(Column)),
                      fill = Genotype)) +
        geom_text(aes(x = Row,
                      y = fct_rev(as.factor(Column)),
                      label = Replicate)) +
        geom_rect(aes(xmin = rowmin, xmax = rowmax,
                      ymin = colmin, ymax = colmax),
                  data = repBorders,
                  colour = "black", fill = NA, size = 1.5) +
        ylab("Column") + xlab("Row") +
        labs(caption = paste0(mytrial,"\n",n_distinct(mydata$Genotype),"genotypes tested")) +
        theme_bw() + theme(legend.position = "none")
      rm(mytrial,mydata)
    }

    sum_list <- list()
    sum_list[[1]] <- .df_sum
    sum_list[[2]] <- layouts

    names(sum_list)[1] <- "Summary of trials"
    names(sum_list)[2] <- "Layout of trials"


  }

  return(sum_list)

}
