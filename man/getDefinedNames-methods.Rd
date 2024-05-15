\name{getDefinedNames-methods}
\docType{methods}
\alias{getDefinedNames}
\alias{getDefinedNames-methods}
\alias{getDefinedNames,workbook-method}
\title{Retrieving defined names in a workbook}
\description{
Retrieves the defined names in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{getDefinedNames}{workbook}(object, validOnly, worksheetScope)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{validOnly}{If \code{validOnly = TRUE} only names with valid references
are returned. Valid references are ones not starting with #REF! or
  #NULL! - which could result e.g. due to a missing sheet reference. The default
  value for \code{validOnly} is \code{TRUE}.}
  \item{worksheetScope}{Optional - the name of the worksheet in which the names are scoped;
  to only query names in the global scope, use the value \code{""}}
}

\details{
  If option \code{XLConnect.setCustomAttributes} is \code{TRUE} (default \code{FALSE}),
  a list of the worksheet scopes in which the names were found is set as attribute \code{worksheetScope} on the result.
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createName-methods]{createName}}, \code{\link[=removeName-methods]{removeName}}, 
\code{\link[=existsName-methods]{existsName}}, \code{\link[=readNamedRegion-methods]{readNamedRegion}},\cr
\code{\link[=writeNamedRegion-methods]{writeNamedRegion}}
}
\examples{
\dontrun{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
mtcarsFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(mtcarsFile)

# Retrieve defined names with valid references
getDefinedNames(wb)
}
}
\keyword{methods}
\keyword{utilities}
