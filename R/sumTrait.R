
#' Explore the observed response of a specific trait
#'
#' @description Aims to explore the observed response of a specific trait. This function considers
#' simple breeding trial designs, namely, rectangular arrays with replicate blocks
#' and genotypes randomised within replicates. Structural terms must be named as "Replicate", "Row"
#' and "Column", and the genotype or variety factor as "Genotype". This function
#' provides summary plots of the trait of interest.
#'
#' @param .df A dataframe object containing the data to be analysed.
#' @param .trait The names of the trait that you are interested in. Trait must be set as numeric in the
#' data set.
#' @param .trial Set as "Trial" by default. If multiple environments are considered, please enter the name of
#' the trial factor.
#'
#' @return If a single trial was considered, returns a list comprised of
#' @return - a ggplot object of the observed response in the field
#' @return - a ggplot object with a histogram of the trait
#'
#' @importFrom magrittr %>%
#' @import ggplot2
#' @import dplyr
#' @importFrom forcats fct_rev
#'
#' @export
#'

sumTrait <- function(.df, .trait = NULL, .trial = "Trial"){

  if(is.null(.trial)){

    cols <- c("Replicate" ,"Column", "Row", "Genotype", .trait)

    .df <- .df %>% select(c(cols))

    # Plot of observed response surface in 2D with ggplot2
    .df$Row <- factor(.df$Row,
                      levels =  sort(as.numeric(levels(.df$Row))))
    .df$Column <- factor(.df$Column,
                         levels =  sort(as.numeric(levels(.df$Column))))

    # To get replicate borders (snippet of code provided by Alec Zwart)
    repBorders <- .df %>%
      group_by(Replicate) %>%
      summarise(rowmin = min(as.integer(Row)) - 0.5, rowmax = max(as.integer(Row)) + 0.5,
                colmin = min(as.integer(Column)) - 0.5, colmax = max(as.integer(Column)) + 0.5)



    field <- ggplot(data = .df) +
      geom_tile(aes(x = Row,
                    y = fct_rev(as.factor(Column)),
                    fill = get(.trait))) +
      geom_rect(aes(xmin = rowmin, xmax = rowmax,
                    ymin = colmin, ymax = colmax),
                data = repBorders,
                colour = "black", fill = NA, size = 1.5) +
      ylab("Column") + xlab("Row") +
      scale_fill_viridis_c(name = .trait) +
      theme_bw()

    hist_gg <- ggplot(data = .df) +
      geom_histogram(aes(x = get(.trait)),
                         color = "black", fill = "lightblue") +
      xlab(.trait) + ylab("Count") +
      theme_bw()

    trait_sum <- list()
    trait_sum[[1]] <- field
    trait_sum[[2]] <- hist_gg

    names(trait_sum)[1] <- "Response in the field"
    names(trait_sum)[2] <- "Histogram"

  } else if(!is.null(.trial)){
    cols <- c(.trial, "Replicate", "Row", "Column", "Genotype",.trait)

    .df <- .df %>% select(c(cols))
    .df[,1] <- as.factor(.df[,1])
    # Trial layouts

    trait_sum <- list()

    for (i in 1:n_distinct(.df[,1])){
      mytrial <- as.character(levels(.df[,1])[i])

      mydata <- droplevels(subset(.df, .df[,1] == mytrial))

      mydata$Row <- factor(mydata$Row,
                           levels =  sort(as.numeric(levels(mydata$Row))))
      mydata$Column <- factor(mydata$Column,
                           levels =  sort(as.numeric(levels(mydata$Column))))

      # To get replicate borders (snippet of code provided by Alec Zwart)
      repBorders <- mydata %>%
        group_by(Replicate) %>%
        summarise(rowmin = min(as.integer(Row)) - 0.5, rowmax = max(as.integer(Row)) + 0.5,
                  colmin = min(as.integer(Column)) - 0.5, colmax = max(as.integer(Column)) + 0.5)

      trait_sum[[mytrial]] <- list()

      trait_sum[[mytrial]][["Response in the field"]] <- ggplot(data = mydata) +
        geom_tile(aes(x = Row,
                      y = fct_rev(as.factor(Column)),
                      fill = get(.trait))) +
        geom_rect(aes(xmin = rowmin, xmax = rowmax,
                      ymin = colmin, ymax = colmax),
                  data = repBorders,
                  colour = "black", fill = NA, size = 1.5) +
        ylab("Column") + xlab("Row") +
        scale_fill_viridis_c(name = .trait) +
        labs(caption = mytrial) +
        theme_bw()

      trait_sum[[mytrial]][["Histogram"]] <- ggplot(data = mydata) +
        geom_histogram(aes(x = get(.trait)),
                       color = "black", fill = "lightblue") +
        xlab(.trait) + ylab("Count") + labs(caption = mytrial) +
      theme_bw()
    }

  }
return(trait_sum)
}
