\name{existsName-methods}
\docType{methods}
\alias{existsName-methods}
\title{Checking existence of names in a workbook}
\description{
Checks the existence of a name in a \code{\linkS4class{workbook}} object.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook", name = "character")}}{
Returns \code{TRUE} if the name with the specified \code{name} exists and \code{FALSE} otherwise. Note that the \code{name}
argument is vectorized and therefore multiple names can be checked for existence in one method call.
}
}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createName-methods]{createName}}, \code{\link[=removeName-methods]{removeName}},
\code{\link[=getDefinedNames-methods]{getDefinedNames}},
\code{\link[=readNamedRegion-methods]{readNamedRegion}}, \code{\link[=writeNamedRegion-methods]{writeNamedRegion}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
mtcarsFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(mtcarsFile)

# Check if the name 'mtcars' exists
# (should return TRUE since the name is defined as 'mtcars!$A$1:$K$33')
existsName(wb, name = "mtcars")
}
\keyword{methods}
\keyword{utilities}
