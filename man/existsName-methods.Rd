\name{existsName-methods}
\docType{methods}
\alias{existsName}
\alias{existsName-methods}
\alias{existsName,workbook-method}
\title{Checking existence of names in a workbook}
\description{
Checks the existence of a name in a \code{\linkS4class{workbook}}.
}

\usage{
\S4method{existsName}{workbook}(object, name)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The name to check for}
}

\details{
Returns \code{TRUE} if the specified \code{name} exists and \code{FALSE}
otherwise. Note that the \code{name} argument is vectorized and
therefore multiple names can be checked for existence in one method call.
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
}
}
\keyword{methods}
\keyword{utilities}
