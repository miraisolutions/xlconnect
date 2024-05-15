\name{getReferenceFormula-methods}
\docType{methods}
\alias{getReferenceFormula}
\alias{getReferenceFormula-methods}
\alias{getReferenceFormula,workbook-method}
\title{Querying reference formulas of Excel names}
\description{
Queries the reference formula of an Excel named range in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{getReferenceFormula}{workbook}(object,name, worksheetScope)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The named range to query. This argument is vectorized such that
  multiple names can be queried with one method call.}
  \item{worksheetScope}{Optional - the name of the worksheet in which the name is scoped;
  if undefined a matching name in any scope may be returned. To specify global scope only, use the value \code{""}}.
}

\details{
  If option \code{XLConnect.setCustomAttributes} is \code{TRUE} (default \code{FALSE}),
  the worksheet scope in which the queried name is defined is set as attribute \code{worksheetScope} on the result.
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createName-methods]{createName}}, \code{\link[=existsName-methods]{existsName}},
\code{\link[=removeName-methods]{removeName}}
}
\examples{
\dontrun{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Query reference formula for name 'mtcars'
print(getReferenceFormula(wb, name = "mtcars"))
}
}
\keyword{methods}
\keyword{utilities}
