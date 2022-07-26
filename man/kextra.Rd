% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/kextra.R
\name{kextra}
\alias{kextra}
\title{Add extra genotypes without marker to kinship matrix}
\usage{
kextra(.df, .kmat, .geno = "Genotype")
}
\arguments{
\item{.df}{A dataframe object containing the data to be analysed}

\item{.kmat}{The initial scaled kinship matrix in matrix format from molecular marker data.}

\item{.geno}{The name of the genotype factor in `.df`. Set as "Genotype" by default.}
}
\value{
A matrix
}
\description{
Returns an expanded version of the kinship
matrix including genotypes without molecular marker data. Diagonal elements for these genotypes
are equal to the maximum value from the diagonal elemnts of the initial kinship matrix. The function
also prints the number of genotypes that did not have molecular marker information and lists them.
}
