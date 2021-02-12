\name{setRowHeight-methods}
\docType{methods}
\alias{setRowHeight}
\alias{setRowHeight-methods}
\alias{setRowHeight,workbook,character-method}
\alias{setRowHeight,workbook,numeric-method}
\title{Setting the height of a row in a worksheet}
\description{
Sets the height of a row in a worksheet.
}
\usage{
\S4method{setRowHeight}{workbook,character}(object,sheet,row,height)
\S4method{setRowHeight}{workbook,numeric}(object,sheet,row,height)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet to edit}
  \item{row}{The index of the row to resize}
  \item{height}{The height in points. If \code{height < 0} (default: -1),
  the row will be sized to the sheet's default row height.}
}

\details{
Note that the arguments \code{sheet}, \code{row} and \code{height} are
vectorized. As such the row height of multiple rows (potentially on
different worksheets) can be set with one method call. 
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=setColumnWidth-methods]{setColumnWidth}}
}
\examples{
\dontrun{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
mtcarsFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(mtcarsFile)

# Sets the row height of the 1st row on sheet 'mtcars'
# to 20 points
setRowHeight(wb, sheet = "mtcars", row = 1, height = 20)
}
}
\keyword{methods}
\keyword{utilities}
