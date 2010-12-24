\name{existsSheet-methods}
\docType{methods}
\alias{existsSheet-methods}
\title{Checking for existence of worksheets in a workbook}
\description{
Checks if worksheets with specified names exist in a \code{\linkS4class{workbook}} object.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook", name = "character")}}{
Checks if the specified worksheet \code{name} exists. Returns \code{TRUE} if it exists, otherwise \code{FALSE}. The \code{name}
argument is vectorized which allows to check for existence of multiple worksheets with one call.
}
}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createSheet-methods]{createSheet}},
\code{\link[=removeSheet-methods]{removeSheet}}, \code{\link[=getSheets-methods]{getSheets}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")
# Load workbook
wb <- loadWorkbook(demoExcelFile)
# Check for existence of a worksheet called 'mtcars'
existsSheet(wb, "mtcars")
}
\keyword{methods}
\keyword{utilities}
