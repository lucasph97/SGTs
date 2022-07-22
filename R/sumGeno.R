
#' Summary of genotype by trial
#'
#' @description Provides information about the genotype concurrences by trial
#' and the genotype by trial.
#'
#' @param .df A dataframe object containing the data to be analysed
#' @param .trial The name of the trial factor in `.df`. Set as "Trial" by default
#' @param .geno The name of the genotype factor in `.df`. Set as "Genotype" by default
#' @param .cbf Set as FALSE by default. If TRUE, the function also returns a heatmap of genotype
#' concurrence by trial with colorblind friendly palette.
#'
#' @return A list containing
#' @return - A list with the genotype concurrence by trial table, a heatmap of this table and, if
#' specified, a heatmap with colorblind friendly palette;
#' @return - A data frame with the table of genotypes by trial
#'
#' @importFrom magrittr %>%
#' @import ggplot2
#' @export
#'

sumGeno <- function(.df, .geno = "Genotype", .trial = "Trial", .cbf = FALSE){

  # 1. Concurrence matrix ----
  tt <- table(.df[,.geno], .df[,.trial])
  tt[tt>1] <- 1
  tt.conc <- t(tt)%*%tt

  myCmat <- tt.conc
  temprow <- rownames(myCmat)
  tempcol <- colnames(myCmat)

  myCmat <- myCmat %>% as.data.frame
  myCmat <- cbind(temprow,myCmat)
  colnames(myCmat) <- c("",tempcol)
  rownames(myCmat) <- NULL
  rm(temprow, tempcol)

  # 1.1. Produce heatmap using ggplot2 (MM's)----

  # Convert genetic concurrence matrix to long format
  Cmat_long <- reshape2::melt(myCmat,)
  head(Cmat_long)
  nrow(Cmat_long)

  colnames(Cmat_long) <- c("Var1","Var2","value")

  # Refacotrise environment levels
  Cmat_long$Var1 <- factor(Cmat_long$Var1)

  Cmat_long$Var2 <- factor(Cmat_long$Var2)

  # Reverse order of factor levels on the y-axis
  Cmat_long$Var2 <- fct_rev(Cmat_long$Var2)

  #Produce heatmap using ggplot
  myheatmap <- ggplot(Cmat_long , aes(Var1 , Var2 )) +
    geom_tile(aes(fill = value)) +
    xlab("Trial") +
    ylab("Trial") +
    scale_fill_gradientn(colours = c("white","grey","yellow","brown"),
                         name = "Genotype\nConcurrence\nby Trial") +
    geom_text(aes(label = value)) +
    guides(fill = guide_colourbar(barwidth = 1, barheight=25)) +
    theme_bw() +
    theme(panel.grid.minor = element_blank(),
          axis.line.x = element_line(colour = "black"),
          legend.title = element_text(size = 17),
          legend.text = element_text(size = 15),
          axis.line.y = element_line(colour = "black"),
          axis.text.x = element_text(size = 12, angle = 45, vjust = 1, hjust = 1),
          axis.title.y = element_text(size = 20),
          axis.title.x = element_text(size = 20),
          axis.text.y = element_text(size = 12),
          strip.text.x = element_text(size = 15),
          strip.text.y = element_text(size = 15))

  if(.cbf == TRUE){
  #Colourblind friendly heatmap using ggplot
    myheatmap_cb <- ggplot(Cmat_long , aes(Var1 , Var2 )) +
      geom_tile(aes(fill = value)) +
      xlab("Trial") +
      ylab("Trial") +
      geom_text(aes(label = value)) +
      guides(fill = guide_colourbar(barwidth = 1, barheight=25)) +
      theme_bw() +
      theme(panel.grid.minor = element_blank(),
            axis.line.x = element_line(colour = "black"),
            legend.title = element_text(size = 17),
            legend.text = element_text(size = 15),
            axis.line.y = element_line(colour = "black"),
            axis.text.x = element_text(size = 12, angle = 45, vjust = 1, hjust = 1),
            axis.title.y = element_text(size = 20),
            axis.title.x = element_text(size = 20),
            axis.text.y = element_text(size = 12),
            strip.text.x = element_text(size = 15),
            strip.text.y = element_text(size = 15)) +
      scale_fill_viridis_c(name="Genotype\nConcurrence\nby Trial")
  }

  tt.conc <- as.data.frame(tt.conc) %>%
    tibble::rownames_to_column("Trial")

  ttgxe <- table(as.character(.df[,.geno]), .df[,.trial])
  ttgxe[ttgxe>1] <- 1
  ttgxe <- as.data.frame.matrix(ttgxe) %>%
    tibble::rownames_to_column("Genotype")

  # Store the outputs in a list
  geno_sum_ls <- list()
  geno_sum_ls[[1]] <- list()
  geno_sum_ls[[2]] <- ttgxe
  geno_sum_ls[[1]][[1]] <- tt.conc
  geno_sum_ls[[1]][[2]] <- myheatmap
  if(.cbf == TRUE){geno_sum_ls[[1]][[3]] <- myheatmap_cb}

  # Name elements from the list
  names(geno_sum_ls)[1] <- "Genotype concurrence by trial"
  names(geno_sum_ls)[2] <- "Genotype by trial table"
  names(geno_sum_ls[[1]])[1] <- "Table"
  names(geno_sum_ls[[1]])[2] <- "Heatmap"
  if(.cbf == TRUE){names(geno_sum_ls[[1]])[3] <- "Heatmap Colorblind firendly"}


  return(geno_sum_ls)
}
