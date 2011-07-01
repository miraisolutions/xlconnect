\name{setAutoFilter-methods}
\docType{methods}
\alias{setAutoFilter}
\alias{setAutoFilter-methods}
\alias{setAutoFilter,workbook,character,character-method}
\alias{setAutoFilter,workbook,numeric,character-method}
\title{Setting auto-filters on worksheets}
\description{
Sets an auto-filter on a specified worksheet. 
}
\usage{
\S4method{setAutoFilter}{workbook,character,character}(object,sheet,reference)
\S4method{setAutoFilter}{workbook,numeric,character}(object,sheet,reference)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet on which to set the auto-filter}
  \item{reference}{A cell range specification (\code{character}) in the form 'A1:B8'}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("autofilter.xlsx", create = TRUE)

# Create a worksheet named 'mtcars'
createSheet(wb, name = "mtcars")

# Create a named region called 'mtcars' on the sheet called 'mtcars'
createName(wb, name = "mtcars", formula = "mtcars!$A$1")

# Write built-in data set 'mtcars' to the above defined named region
# (using header = TRUE)
writeNamedRegion(wb, mtcars, name = "mtcars")

# Set an auto-filter for the named region written above
setAutoFilter(wb, sheet = "mtcars", reference = aref("A1", dim(mtcars)))

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
}
\keyword{methods}
\keyword{utilities}
