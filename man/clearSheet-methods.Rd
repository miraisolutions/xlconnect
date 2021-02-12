\name{clearSheet-methods}
\docType{methods}
\alias{clearSheet}
\alias{clearSheet-methods}
\alias{clearSheet,workbook,numeric-method}
\alias{clearSheet,workbook,character-method}
\title{Clearing worksheets in a workbook}
\description{
Clears worksheets with specified names or indices in a \code{\linkS4class{workbook}}.
}
\usage{
  \S4method{clearSheet}{workbook,numeric}(object, sheet)
  \S4method{clearSheet}{workbook,character}(object, sheet)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or the index of the worksheet to clear}
}

\details{
Clearing a worksheet means to clear all the cells in that worksheet. Consequently,
the saved workbook should be smaller in size. Clearing a worksheet can be useful
if data sets in a worksheet need to be replaced, i.e. data are first read, modified
in R and finally written back to the worksheet. Without clearing the worksheet first, 
(parts of) the original data may still be visible if they occupied a larger range of 
the worksheet.
}

\author{
Nicola Lambiase\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=clearNamedRegion-methods]{clearNamedRegion}}, 
\code{\link[=clearRange-methods]{clearRange}}, 
\code{\link[=clearRangeFromReference-methods]{clearRangeFromReference}}
}
\examples{
\dontrun{
# mtcars xlsx file from demoFiles subfolder of 
# package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", 
                             package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Clear worksheets named 'mtcars' and 'mtcars2'
clearSheet(wb, sheet = c("mtcars", "mtcars2"))

# Clear 3rd worksheet
clearSheet(wb, sheet = 3) 
}
}
\keyword{methods}
\keyword{utilities}
