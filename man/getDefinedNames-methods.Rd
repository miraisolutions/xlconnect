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
\S4method{getDefinedNames}{workbook}(object, validOnly)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{validOnly}{If \code{validOnly = TRUE} only names with valid references
are returned. Valid references are ones not starting with #REF! or
  #NULL! - which could result e.g. due to a missing sheet reference. The default
  value for \code{validOnly} is \code{TRUE}.}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createName-methods]{createName}}, \code{\link[=removeName-methods]{removeName}}, 
\code{\link[=existsName-methods]{existsName}}, \code{\link[=readNamedRegion-methods]{readNamedRegion}},\cr
\code{\link[=writeNamedRegion-methods]{writeNamedRegion}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
mtcarsFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(mtcarsFile)

# Retrieve defined names with valid references
getDefinedNames(wb)
}
\keyword{methods}
\keyword{utilities}
