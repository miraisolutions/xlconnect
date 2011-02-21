\name{setColumnWidth-methods}
\docType{methods}
\alias{setColumnWidth}
\alias{setColumnWidth-methods}
\alias{setColumnWidth,workbook,character,numeric,numeric-method}
\alias{setColumnWidth,workbook,numeric,numeric,numeric-method}
\title{Setting the width of a column in a worksheet}
\description{
Sets the width of a column in a worksheet.
}
\usage{
\S4method{setColumnWidth}{workbook,character,numeric,numeric}(object,sheet,column,width)
\S4method{setColumnWidth}{workbook,numeric,numeric,numeric}(object,sheet,column,width)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet}
  \item{column}{The index of the column to resize}
  \item{width}{The width of the specified column in  units of 1/256th of a character width}
}

\details{
Note that the arguments \code{sheet}, \code{column} and \code{width} are
vectorized. As such the column width of multiple columns (potentially on
different sheets) can be set with one method call. 
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=setRowHeight-methods]{setRowHeight}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
mtcarsFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(mtcarsFile)

# Sets the column width of the 3rd column on sheet 'mtcars' 
# to 4000/256th (= 15.625) character width
setColumnWidth(wb, sheet = "mtcars", column = 3, width = 4000)
}
\keyword{methods}
\keyword{utilities}
