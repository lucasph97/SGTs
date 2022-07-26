% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/useiClass.R
\name{useiClass}
\alias{useiClass}
\title{Use Interaction Classes}
\usage{
useiClass(.fa.asreml.obj, .factors = NULL, .output = "iClass_output.xlsx")
}
\arguments{
\item{.factors}{The factors that will be used to create the interaction classes. Set by default as the
order of the FA model, namely, all factors are considered. It is not necessary to specify contiguous factors.
For example, it is possible to specify factors 1, 3 and 7, without considering factors 2, 4, 5 and 6.}

\item{.output}{An *.xlsx file to write the outputs. The default is "iClass_output.xlsx".}

\item{.fa.asreml.object}{The fa.asreml object from ASExtras4 package}
}
\value{
A list containing

- $iClasses the A data frame containing a table of the composition of the iClasses

- $iClassOP_wide A data frame containing the iClass Overall Performance measures in wide format

- $iClassOP_long A data frame containing the iClass Overall Performance measures in long format

A spreadsheet containing the information from the list. The spreadsheet can be printed in a new or
existing *.xlsx file.
}
\description{
Prints the results from the use of interaction classes, iClass method proposed by Smith et al. 2021.
in a *.xlsx file using openxlsx. The function also returns a list containing the results.
}
