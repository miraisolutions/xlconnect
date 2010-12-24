\name{setRowHeight-methods}
\docType{methods}
\alias{setRowHeight-methods}
\title{Setting the height of a row in a worksheet}
\description{
Sets the height of a row in a worksheet.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook", sheet = "character", row = "numeric", height = "numeric")}}{
Sets the \code{height} (in points) of the specified \code{row} on (named) \code{sheet}. Note that the arguments
\code{sheet}, \code{row} and \code{height} are vectorized. As such the row height of multiple rows (potentially on
different worksheets) can be set with one method call.
}

\item{\code{signature(object = "workbook", sheet = "numeric", row = "numeric", height = "numeric")}}{
As above but with the \code{sheet} specified by index.
}
}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=setColumnWidth-methods]{setColumnWidth}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
mtcarsFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(mtcarsFile)

# Sets the row height of the 1st row on sheet 'mtcars'
# to 20 points
setRowHeight(wb, sheet = "mtcars", row = 1, height = 20)
}
\keyword{methods}
\keyword{utilities}
