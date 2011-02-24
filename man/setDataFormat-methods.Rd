\name{setDataFormat-methods}
\docType{methods}
\alias{setDataFormat}
\alias{setDataFormat-methods}
\alias{setDataFormat,cellstyle,character-method}
\title{Specifying custom data formats for cell styles}
\description{
Specifies a custom data format for a \code{\linkS4class{cellstyle}}.
}
\usage{
\S4method{setDataFormat}{cellstyle,character}(object,format)
}

\arguments{
  \item{object}{The \code{\linkS4class{cellstyle}} to use}
  \item{format}{A data format string}
}

\details{
Specifies the data \code{format} to be used by the corresponding
\code{\linkS4class{cellstyle}}. Data formats are specified the standard
Excel way. Refer to the Excel help or to the link below for more
information. 
}
\references{
Excel custom data formats:\cr
\url{http://www.ozgrid.com/Excel/CustomFormats.htm}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\linkS4class{cellstyle}}, \code{\link[=setCellStyle-methods]{setCellStyle}}, 
\code{\link[=setStyleAction-methods]{setStyleAction}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("setDataFormat.xlsx", create = TRUE)

# Create a worksheet
createSheet(wb, name = "cellstyles")

# Create a dummy data set with the current date/time (as POSIXct)
now <- data.frame(Now = Sys.time())

# Write the value to the 'cellstyles' worksheet in the top left 
# corner (cell A1)
writeWorksheet(wb, now, sheet = "cellstyles", startRow = 1, 
               startCol = 1, header = FALSE)

# Create a custom anonymous cell style
cs <- createCellStyle(wb)

# Specify a custom data format
setDataFormat(cs, format = "dddd d-m-yyyy h:mm AM/PM")

# Set the cell style created above for the top left cell (A1) in 
# the 'cellstyles' worksheet
setCellStyle(wb, sheet = "cellstyles", row = 1, col = 1, cellstyle = cs)

# Set column width to display whole time/date string
setColumnWidth(wb, sheet = "cellstyles", column = 1, width = 6000)

# Save the workbook
saveWorkbook(wb)
}
\keyword{methods}
\keyword{utilities}
