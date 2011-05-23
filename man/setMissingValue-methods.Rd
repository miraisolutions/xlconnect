\name{setMissingValue-methods}
\docType{methods}
\alias{setMissingValue}
\alias{setMissingValue-methods}
\alias{setMissingValue,workbook-method}
\alias{setMissingValue,workbook,ANY-method}
\title{Setting a missing value string}
\description{
Sets the missing value string used when writing missing values to Excel.
}
\usage{
\S4method{setMissingValue}{workbook,ANY}(object,value)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{value}{The missing value string to use when writing data to Excel.
  If \code{NULL}, missing values result in blank cells (which is the default behavior). Any
  other values are coerced to \code{character} and used as missing value string.}
}

\details{
If no specific missing value string is set using \code{setMissingValue} then the default behavior is
to map missing values to blank (empty) cells. The reason for this default mechanism is that the reverse
operation of reading the data back into R results in the same original data types. However, if a missing 
value string is used (such as "NA"), then a numeric column for example with corresponding missing values 
would be read back in as a string (character) column.\cr
If you want to use e.g. numeric missing value identifiers such as e.g. \code{-9999} then you should replace
missing values first before you write them to Excel using \pkg{XLConnect}.
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
