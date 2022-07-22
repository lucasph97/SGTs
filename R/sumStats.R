
#' Summary stats of recorded traits
#'
#' @description Provides a data frame with minimum, median and maximum values from traits of
#' interest and the number of trials that the trait was measured for each genotype, together with pedigree
#' information, if any. The output from this function can be merged with predicted means from other analysed
#' traits. Once the table of summary statistics is complete, with or without additional predicted means,
#' everything should be ready to use printSumStats, the function to print this table in a nice way in an
#' *.xlsx file.
#'
#' @param .df A dataframe object containing the data to be analysed.
#' @param .trait A vector containing the index or names from the traits that you are interested in
#' @param .geno The name of the genotype factor in `.df`. Set as "Genotype" by default
#' @param .trial The name of the trial factor in `.df`. Set as "Trial" by default
#' summarising. Traits must be set as numeric. The function provides the mean, standard deviation and
#' missing values.
#' @param .pedigree A data frame containing the pedigree information from each genotype. Set as NULL by default.
#'
#' @importFrom magrittr %>%
#' @importFrom stringr str_split_fixed
#' @import dplyr
#'
#' @export
#'
sumStats <- function(.df, .trait,.geno = "Genotype", .trial = "Trial", .pedigree = NULL){

  .df_sum <- .df %>% select(c(all_of(.trial), all_of(.geno), all_of(.trait)))

  # Set traits as numeric
  t_fun <- function(x){as.numeric(as.character(x))}

  .df_sum[,colnames(.df_sum) %in% .trait] <- sapply(.df_sum[,colnames(.df_sum) %in% .trait],
                                                    t_fun)

  for(i in 1:length(.trait)){
    c.trait <- .trait[i]
    temp <- .df_sum %>% select(c(.trial, .geno, all_of(c.trait)))

    # Exclude NA values and drop levels - this is to correctly count the nTrials
    temp <- droplevels(subset(temp, !is.na(temp[,3])))

    # Make a frequency table
    check <- table(temp$Trial,temp$Genotype)
    # If the genotype is present, set frequency to 1
    check[check>1] <- 1

    # Melt the table - change from matrix to data.frame
    # Then group by Genotype (Var2) and sum the number of times that appears in
    # each trial. Finally we will get the nTrial by pasting "n" to the number of
    # trials where the trait was recorded for each genotype
    check <- reshape2::melt(check) %>% group_by(Var2) %>%
      summarise(nTrial = sum(value))
    colnames(check)[1] <- "Genotype"
    check$nTrial <- paste0("n", check$nTrial)

    # Now get min, median and max from the trait for each genotype across all
    # trials. "!! rlang::sym() is a way to get the colname without writing it. It
    # is just for practical purposes and to make the loop as general as possible.
    # This way we can cover most scenarios regardless the name of the recorded
    # trait.
    temp <- temp %>% group_by(Genotype) %>%
      summarise(Min = min(!! rlang::sym(names(.)[3]), na.rm = T),
                Median = round(median(!! rlang::sym(names(.)[3]), na.rm = T), 0),
                Max = max(!! rlang::sym(names(.)[3]), na.rm = T)) %>%
      as.data.frame

    # Merge the nTrial info with the stats info and create Summary column by
    # pasting nTrial and stats
    temp <- left_join(check, temp, by = .geno)
    temp$Summary <- apply(temp[ ,2:5], 1, paste, collapse = "|")

    # The following lines work to change column names to identify the trait in the
    # final summary table
    colnames(temp)[2] <- paste0(c.trait, "_", colnames(temp)[2])
    colnames(temp)[3] <- paste0(c.trait, "_", colnames(temp)[3])
    colnames(temp)[4] <- paste0(c.trait, "_", colnames(temp)[4])
    colnames(temp)[5] <- paste0(c.trait, "_", colnames(temp)[5])
    colnames(temp)[6] <- paste0(c.trait, "_", colnames(temp)[6])

    # This works to bind the tables from each trait. The first table will be the
    # summmary of the first trait (i). Then all tables will be merged until all
    # traits are summarised (loop finished)

    if(i == 1){sumtrait <- temp}
    else if(i >1){sumtrait <- left_join(sumtrait, temp, by = .geno)}
    rm(check, temp)

  }

  if(!is.null(.pedigree)){
    sumtrait <- inner_join(sumtrait,.pedigree, by = .geno) %>%
      as.data.frame
    sumtrait <- relocate(sumtrait, c("Genotype","Pedigree"))
  }
  return(sumtrait)
}
