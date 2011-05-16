\name{getSheets-methods}
\docType{methods}
\alias{getSheets}
\alias{getSheets-methods}
\alias{getSheets,workbook-method}
\title{Querying available worksheets in a workbook}
\description{
Returns all worksheet names in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{getSheets}{workbook}(object)
}
\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createSheet-methods]{createSheet}},
\code{\link[=removeSheet-methods]{removeSheet}}, \code{\link[=renameSheet-methods]{renameSheet}},
\code{\link[=getSheetPos-methods]{getSheetPos}}, \code{\link[=setSheetPos-methods]{setSheetPos}}
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
