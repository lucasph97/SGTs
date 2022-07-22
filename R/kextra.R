
#' Add extra genotypes without marker to kinship matrix
#'
#' @description Returns an expanded version of the kinship
#' matrix including genotypes without molecular marker data. Diagonal elements for these genotypes
#' are equal to the maximum value from the diagonal elemnts of the initial kinship matrix. The function
#' also prints the number of genotypes that did not have molecular marker information and lists them.
#'
#' @param .df A dataframe object containing the data to be analysed
#' @param .kmat The initial scaled kinship matrix in matrix format from molecular marker data.
#' @param .geno The name of the genotype factor in `.df`. Set as "Genotype" by default.
#'
#' @return A matrix
#'
#' @importFrom magrittr %>%
#' @importFrom dplyr select
#' @export
#'

# Returns expanded kinship matrix with

kextra <- function(.df, .kmat, .geno = "Genotype"){
  genos <- .df %>% select(c(.geno))
  # Genotypes without markers
  womarker <- droplevels(genos[!(genos %in% rownames(kmat))])
  nextras <- dplyr::n_distinct(womarker)

  # Change k matrix to sparse format
  kmat_sparse <- pedicure::mat2sparse(kmat)
  # Create data frame for genotypes without markers with max(diag(kmat) on the
  # diagonal (in sparse form)
  extra <- matrix(c((ncol(kmat) + 1):(ncol(kmat) + nextras),
                    (ncol(kmat) + 1):(ncol(kmat) + nextras),
                    rep(max(diag(kmat)), nextras)), ncol = 3, byrow = FALSE)

  # Bind them together
  colnames(extra) <- colnames((kmat_sparse))
  kmat_sparse2 <- rbind(kmat_sparse,
                        extra)
  # Set the rowNames attribute. Use from the original kmat and the levels from
  # the genotypes without markers
  attr(kmat_sparse2, "rowNames") <- c(rownames(kmat),levels(nomark))

  # Change to matrix format
  kmat_f <- pedicure::sparse2mat(kmat_sparse2)


  # Print genotypes without markers
  message(paste("There were",nextras,"genotypes without marker information. These genotypes were:"))
  print(cbind(levels(womarker)))
  # Return the expanded kinship matrix
  return(kmat_f)

}

