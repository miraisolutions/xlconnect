\name{getTables-methods}
\docType{methods}
\alias{getTables}
\alias{getTables-methods}
\alias{getTables,workbook,numeric-method}
\alias{getTables,workbook,character-method}
\title{Querying available Excel tables in a workbook}
\description{
Queries the available Excel tables on the specified worksheet.
}
\usage{
\S4method{getTables}{workbook,numeric}(object,sheet)
\S4method{getTables}{workbook,character}(object,sheet)
}
\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{Index (\code{integer}) or name (\code{character}) of worksheet to query}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=getSheets-methods]{getSheets}},
\code{\link[=readTable-methods]{readTable}}
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
