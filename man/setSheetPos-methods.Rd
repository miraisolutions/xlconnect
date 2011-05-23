\name{setSheetPos-methods}
\docType{methods}
\alias{setSheetPos}
\alias{setSheetPos-methods}
\alias{setSheetPos,workbook,character,missing-method}
\alias{setSheetPos,workbook,character,numeric-method}
\title{Setting worksheet position}
\description{
Sets the position of a worksheets in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{setSheetPos}{workbook,character,numeric}(object,sheet,pos)
}
\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name of the worksheet (\code{character}) whose position to set. 
  This argument is vectorized such that the positions of multiple worksheets can be 
  set with one method call.}
  \item{pos}{The position index to set for the corresponding \code{sheet}. If missing, sheets
  will be positioned in the order they are specified in the argument \code{sheet}.}
}
\details{
It is important to note that the worksheet positions will be applied one after the other in the order
they have been specified.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=getSheetPos-methods]{getSheetPos}},
\code{\link[=getSheets-methods]{getSheets}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Move the 'mtcars3' worksheet to the front
setSheetPos(wb, sheet = "mtcars3", pos = 1)
}
\keyword{methods}
\keyword{utilities}
