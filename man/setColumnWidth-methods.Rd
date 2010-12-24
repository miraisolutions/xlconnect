\name{setColumnWidth-methods}
\docType{methods}
\alias{setColumnWidth-methods}
\title{Setting the width of a column in a worksheet}
\description{
Sets the width of a column in a worksheet.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook", sheet = "character", column = "numeric", width = "numeric")}}{
Sets the \code{width} of the specified \code{column} on (named) \code{sheet} in units of 1/256th of a character width.
Note that the arguments \code{sheet}, \code{column} and \code{width} are vectorized. As such the column width of
multiple columns (potentially on different sheets) can be set with one method call.
}

\item{\code{signature(object = "workbook", sheet = "numeric", column = "numeric", width = "numeric")}}{
As above but with the \code{sheet} specified by index.
}
}}
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
