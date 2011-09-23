\name{existsSheet-methods}
\docType{methods}
\alias{existsSheet}
\alias{existsSheet-methods}
\alias{existsSheet,workbook-method}
\title{Checking for existence of worksheets in a workbook}
\description{
Checks the existence of a worksheet in a \code{\linkS4class{workbook}}.
}

\usage{
\S4method{existsSheet}{workbook}(object,name)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The sheet name to check for}
}

\details{
Checks if the specified worksheet exists. Returns \code{TRUE} if it
exists, otherwise \code{FALSE}. The \code{name} argument is vectorized
which allows to check for existence of multiple worksheets with one call.
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createSheet-methods]{createSheet}},
\code{\link[=removeSheet-methods]{removeSheet}}, \code{\link[=renameSheet-methods]{renameSheet}},
\code{\link[=getSheets-methods]{getSheets}}, \code{\link[=cloneSheet-methods]{cloneSheet}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Check for existence of a worksheet called 'mtcars'
existsSheet(wb, "mtcars")
}
\keyword{methods}
\keyword{utilities}
