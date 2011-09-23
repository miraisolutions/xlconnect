\name{getReferenceFormula-methods}
\docType{methods}
\alias{getReferenceFormula}
\alias{getReferenceFormula-methods}
\alias{getReferenceFormula,workbook-method}
\title{Querying reference formulas of Excel names}
\description{
Queries the reference formula of an Excel name in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{getReferenceFormula}{workbook}(object,name)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The name to query. This argument is vectorized such that
  multiple names can be queried with one method call.}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createName-methods]{createName}}, \code{\link[=existsName-methods]{existsName}},
\code{\link[=removeName-methods]{removeName}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Query reference formula for name 'mtcars'
print(getReferenceFormula(wb, name = "mtcars"))
}
\keyword{methods}
\keyword{utilities}
