\name{print-methods}
\docType{methods}
\alias{print}
\alias{print-methods}
\alias{print,workbook-method}
\title{ Print a workbook's filename }
\description{
Prints the \code{\linkS4class{workbook}}'s underlying filename.
}
\usage{
\S4method{print}{workbook}(x,...)
}

\arguments{
  \item{x}{The \code{\linkS4class{workbook}} to print}
  \item{...}{Arguments passed on to standard \code{print}}
}

\details{
Prints the specified \code{\linkS4class{workbook}}'s filename (see also the S4 \code{filename}
slot of the \code{\linkS4class{workbook}} class).
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}
}
\examples{
\dontrun{
# Load existing demo Excel file 'mtcars.xlsx' from the XLConnect package
wb.mtcars <- loadWorkbook(system.file("demoFiles/mtcars.xlsx", 
                          package = "XLConnect"))

# Print the workbook's underlying filename
print(wb.mtcars)
}
}
\keyword{methods}
\keyword{print}
