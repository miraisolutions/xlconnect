\name{getSheets-methods}
\docType{methods}
\alias{getSheets}
\alias{getSheets-methods}
\alias{getSheets,workbook-method}
\title{Querying available worksheets in a workbook}
\description{
Queries the available worksheets (names) in a \code{\linkS4class{workbook}} object.
}
\usage{
\S4method{getSheets}{workbook}(object)
}
\arguments{
  \item{object}{The workbook to use}
}
\description{
Returns all the worksheet names of the specified \code{\linkS4class{workbook}} object.
}
\author{
Martin Studer\cr
\href{http://www.mirai-solutions.com}{Mirai Solutions GmbH}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createSheet-methods]{createSheet}},
\code{\link[=removeSheet-methods]{removeSheet}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")
# Load workbook
wb <- loadWorkbook(demoExcelFile)
# Query available worksheets
sheets <- getSheets(wb)
}
\keyword{methods}
\keyword{utilities}
