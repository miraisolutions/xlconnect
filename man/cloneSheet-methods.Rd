\name{cloneSheet-methods}
\docType{methods}
\alias{cloneSheet}
\alias{cloneSheet-methods}
\alias{cloneSheet,workbook,character-method}
\alias{cloneSheet,workbook,numeric-method}
\title{Cloning/copying worksheets}
\description{
Clones (copies) a worksheet in a \code{\linkS4class{workbook}}.
}

\usage{
\S4method{cloneSheet}{workbook,numeric}(object,sheet,name)
\S4method{cloneSheet}{workbook,character}(object,sheet,name)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the worksheet to clone}
  \item{name}{The name to assign to the cloned worksheet. Throws an exception if the name to assign
  is the name of an already existing worksheet.}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createSheet-methods]{createSheet}},
\code{\link[=removeSheet-methods]{removeSheet}}, \code{\link[=renameSheet-methods]{renameSheet}},
\code{\link[=getSheets-methods]{getSheets}}, \code{\link[=existsSheet-methods]{existsSheet}}
}
\examples{
\dontrun{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Clone the 'mtcars' worksheet and assign it the name 'mtcars cloned'
cloneSheet(wb, sheet = "mtcars", name = "mtcars cloned")
}
}
\keyword{methods}
\keyword{utilities}
