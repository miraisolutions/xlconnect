\name{unhideSheet-methods}
\docType{methods}
\alias{unhideSheet-methods}
\title{Unhiding worksheets in a workbook}
\description{
Unhides the specified worksheets in a \code{\linkS4class{workbook}} object.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook", sheet = "character")}}{
Unhides the specified (named) \code{sheet}(s).
}

\item{\code{signature(object = "workbook", sheet = "numeric")}}{
Unhides the specified (indexed) \code{sheet}(s).
}
}}
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
