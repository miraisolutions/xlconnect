\name{existsName-methods}
\docType{methods}
\alias{existsName}
\alias{existsName-methods}
\alias{existsName,workbook-method}
\title{Checking existence of named ranges in a workbook}
\description{
Checks the existence of a named range in a \code{\linkS4class{workbook}}.
}

\usage{
\S4method{existsName}{workbook}(object, name, worksheetScope)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The name to check for}
  \item{worksheetScope}{Optional - the specific worksheet to check}
}

\details{
Returns \code{TRUE} if the specified \code{name} exists and \code{FALSE}
otherwise. Note that the \code{name} argument is vectorized and
therefore multiple names can be checked for existence in one method call.

If \code{worksheetScope} is provided, TRUE will be returned only if a matching
named range exists in the local scope of the specified sheet. To explicitly match only
 in the global scope, pass \code{""} as the value.

If option \code{XLConnect.setCustomAttributes} is \code{TRUE} (default \code{FALSE}),
the worksheet scope in which the name is defined is set as attribute \code{worksheetScope} on the result.
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createName-methods]{createName}}, \code{\link[=removeName-methods]{removeName}},
\code{\link[=getDefinedNames-methods]{getDefinedNames}}, \code{\link[=readNamedRegion-methods]{readNamedRegion}},\cr 
\code{\link[=writeNamedRegion-methods]{writeNamedRegion}}
}
\examples{
\dontrun{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
mtcarsFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(mtcarsFile)

# Check if the name 'mtcars' exists
# (should return TRUE since the name is defined as 'mtcars!$A$1:$K$33')
existsName(wb, name = "mtcars")

# check if the worksheet-scoped name 'iris' exists
options(XLConnect.setCustomAttributes = TRUE)
wb <- loadWorkbook("demoFiles/iris.xlsx")

# should return TRUE with worksheet scope "iris"
res <- existsName(wb, name = "iris")
res
attributes(res)
}
}
\keyword{methods}
\keyword{utilities}
