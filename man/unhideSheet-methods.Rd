\name{unhideSheet-methods}
\docType{methods}
\alias{unhideSheet}
\alias{unhideSheet-methods}
\alias{unhideSheet,workbook,character-method}
\alias{unhideSheet,workbook,numeric-method}
\title{Unhiding worksheets in a workbook}
\description{
Unhides the specified worksheets in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{unhideSheet}{workbook,character}(object,sheet)
\S4method{unhideSheet}{workbook,numeric}(object,sheet)
}
\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet to unhide}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=hideSheet-methods]{hideSheet}},
\code{\link[=isSheetHidden-methods]{isSheetHidden}}, \code{\link[=isSheetVeryHidden-methods]{isSheetVeryHidden}},
\code{\link[=isSheetVisible-methods]{isSheetVisible}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("unhideWorksheet.xlsx", create = TRUE)

# Create sheet 'airquality'
createSheet(wb, name = "airquality")

# Write the built-in data set airquality to worksheet
# 'airquality'
writeWorksheet(wb, airquality, sheet = "airquality")

# Create sheet 'CO2'
createSheet(wb, name = "CO2")

# Write the built-in data set CO2 to worksheet 'C02'
writeWorksheet(wb, CO2, sheet = "CO2")

# Hide sheet 'airquality'
hideSheet(wb, sheet = "airquality")

# Unhide sheet 'airquality'
unhideSheet(wb, sheet = "airquality")
}
\keyword{methods}
\keyword{utilities}
