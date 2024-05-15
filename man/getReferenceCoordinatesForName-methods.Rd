\name{getReferenceCoordinatesForName-methods}
\docType{methods}
\alias{getReferenceCoordinatesForName}
\alias{getReferenceCoordinatesForName-methods}
\alias{getReferenceCoordinatesForName,workbook-method}
\title{Querying the coordinates of the range reference by an Excel name}
\description{
Queries the coordinates of an Excel named range in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{getReferenceCoordinatesForName}{workbook}(object,name, worksheetScope)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The name to query. This argument is vectorized such that
  multiple names can be queried with one method call.}
  \item{worksheetScope}{Optional, the name of the worksheet to use for resolving the named region}
}
\details{
  If \code{worksheetScope} is defined, only coordinates for a range scoped strictly to the specified worksheet are 
  returned. To explicitly only query for named ranges in the global scope, pass \code{""} as the value.
}

\author{
Thomas Themel\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createName-methods]{createName}}, \code{\link[=existsName-methods]{existsName}},
\code{\link[=removeName-methods]{removeName}}, \code{\link[=getReferenceFormula-methods]{getReferenceFormula}}, 
\code{\link[=getReferenceCoordinatesForTable-methods]{getReferenceCoordinatesForTable}}
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
