\name{getDefinedNames-methods}
\docType{methods}
\alias{getDefinedNames-methods}
\title{Retrieving defined names in a workbook}
\description{
Retrieves the defined names in a \code{\linkS4class{workbook}} object.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook", validOnly = "logical")}}{
Queries the defined names in a \code{\linkS4class{workbook}} object.  If \code{validOnly = TRUE} only names with valid references
are returned. Valid references are ones not starting with #REF! or #NULL! - which could result e.g. due to a missing sheet reference.
}

\item{\code{signature(object = "workbook", validOnly = "missing")}}{
As above with \code{validOnly = TRUE}.
}
}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createName-methods]{createName}}, \code{\link[=removeName-methods]{removeName}}, 
\code{\link[=existsName-methods]{existsName}},
\code{\link[=readNamedRegion-methods]{readNamedRegion}}, \code{\link[=writeNamedRegion-methods]{writeNamedRegion}}
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
