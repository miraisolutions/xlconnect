\name{removeName-methods}
\docType{methods}
\alias{removeName}
\alias{removeName-methods}
\alias{removeName,workbook-method}
\title{Removing names from workbooks}
\description{
Removes a named range reference from a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{removeName}{workbook}(object,name,worksheetScope)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The name to delete}
  \item{worksheetScope}{Optional - the name of the worksheet in which the name is scoped;
  useful if different sheets have locally-scoped named ranges with the same name.}
}

\details{
Removes the named range reference \code{name} from the specified workbook \code{object} if it does exist.
Data in the referenced cells remains unchanged. Multiple names can be specified to be removed.
Use \code{worksheetScope = ""} to only target names defined in the global scope.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
  \code{\linkS4class{workbook}},
  \code{\link[=createName-methods]{createName}},
  \code{\link[=existsName-methods]{existsName}},\cr
  \code{\link[=getDefinedNames-methods]{getDefinedNames}},
  \code{\link[=readNamedRegion-methods]{readNamedRegion}},
  \code{\link[=writeNamedRegion-methods]{writeNamedRegion}}
}
\examples{
\dontrun{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
mtcarsFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(mtcarsFile)

# Remove the named region called 'mtcars' from the above file
# (this named region is defined as 'mtcars!$A$1:$K$33')
removeName(wb, name = "mtcars")
}
}
\keyword{methods}
\keyword{utilities}
