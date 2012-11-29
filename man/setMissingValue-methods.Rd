\name{setMissingValue-methods}
\docType{methods}
\alias{setMissingValue}
\alias{setMissingValue-methods}
\alias{setMissingValue,workbook-method}
\alias{setMissingValue,workbook,ANY-method}
\title{Setting missing value identifiers}
\description{
Defines the set of missing values (character or numeric) used when reading and writing data.
}
\usage{
\S4method{setMissingValue}{workbook,ANY}(object,value)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{value}{\code{vector} or \code{list} of missing value identifiers (either
  \code{character} or \code{numeric}) that are recognized as missing (\code{NA}) 
  when reading data. The first element of this vector will be used as missing value 
  identifier when writing data. If \code{value = NULL} (default), missing values are 
  represented by blank cells and only blank cells are recognized as missing.}
}

\details{
If there are no specific missing value identifiers defined the default behavior is
to map missing values to blank (empty) cells. Otherwise, each string or numeric cell is 
checked if it matches one of the defined missing value identifiers. In addition, the first 
missing value identifier (i.e. the first element of the \code{value} argument) is used to represent missing values when writing data.\cr
Note that the missing value identifiers have to be either \code{character} or \code{numeric}.
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=writeNamedRegion-methods]{writeNamedRegion}},
\code{\link[=writeWorksheet-methods]{writeWorksheet}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("missingValue.xlsx", create = TRUE)

# Create a worksheet named 'airquality'
createSheet(wb, name = "airquality")

# Create a named region called 'airquality' on the sheet called
# 'airquality'
createName(wb, name = "airquality", formula = "airquality!$A$1")

# Set the missing value string to 'missing'
setMissingValue(wb, value = "missing")

# Write built-in data set 'airquality' to the above defined named region
writeNamedRegion(wb, airquality, name = "airquality")

# Save workbook
saveWorkbook(wb)
}
\keyword{methods}
\keyword{utilities}
