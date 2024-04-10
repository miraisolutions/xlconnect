\name{getReferenceCoordinates-methods}
\docType{methods}
\alias{getReferenceCoordinates}
\alias{getReferenceCoordinates-methods}
\alias{getReferenceCoordinates,workbook-method}
\title{Querying the coordinates of the range reference by an Excel name}
\description{
(DEPRECATED) Queries the coordinates of an Excel named range in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{getReferenceCoordinates}{workbook}(object,name)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The name to query. This argument is vectorized such that
  multiple names can be queried with one method call.}
}
\note{
This function is deprecated. Use \code{\link[=getReferenceCoordinatesForName-methods]{getReferenceCoordinatesForName}}
instead.
}
\author{
Thomas Themel\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createName-methods]{createName}}, \code{\link[=existsName-methods]{existsName}},
\code{\link[=removeName-methods]{removeName}}, \code{\link[=getReferenceFormula-methods]{getReferenceFormula}}
}
\examples{
\dontrun{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Query reference coordinate for name 'mtcars'
print(getReferenceCoordinatesForName(wb, name = "mtcars"))
}
}
\keyword{methods}
\keyword{utilities}
