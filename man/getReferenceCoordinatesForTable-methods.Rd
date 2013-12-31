\name{getReferenceCoordinatesForTable-methods}
\docType{methods}
\alias{getReferenceCoordinatesForTable}
\alias{getReferenceCoordinatesForTable-methods}
\alias{getReferenceCoordinatesForTable,workbook,numeric-method}
\alias{getReferenceCoordinatesForTable,workbook,character-method}
\title{Querying the coordinates of the range of an Excel table}
\description{
Queries the coordinates of an Excel table (Office 2007+) in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{getReferenceCoordinatesForTable}{workbook,numeric}(object,sheet,table)
\S4method{getReferenceCoordinatesForTable}{workbook,character}(object,sheet,table)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The index or name of the worksheet on which to look for the specified \code{table}}
  \item{table}{The name of the table to query. This argument is vectorized such that
  multiple tables can be queried with one method call.}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createName-methods]{createName}}, \code{\link[=existsName-methods]{existsName}},
\code{\link[=removeName-methods]{removeName}}, \code{\link[=getReferenceFormula-methods]{getReferenceFormula}},
\code{\link[=getReferenceCoordinatesForName-methods]{getReferenceCoordinatesForName}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Query reference coordinates for table 'MtcarsTable' on sheet
# 'mtcars_table'
print(getReferenceCoordinatesForTable(wb, sheet = "mtcars_table", 
                                      table = "MtcarsTable"))
}
\keyword{methods}
\keyword{utilities}
