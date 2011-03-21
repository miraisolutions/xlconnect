\name{setActiveSheet-methods}
\docType{methods}
\alias{setActiveSheet}
\alias{setActiveSheet-methods}
\alias{setActiveSheet,workbook,character-method}
\alias{setActiveSheet,workbook,numeric-method}
\title{Setting the active worksheet in a workbook}
\description{
Sets the active worksheet of a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{setActiveSheet}{workbook,character}(object,sheet)
\S4method{setActiveSheet}{workbook,numeric}(object,sheet)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet to activate}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
The active worksheet of a \code{\linkS4class{workbook}} is the worksheet that is displayed
when the corresponding Excel file is opened.
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createSheet-methods]{createSheet}},
\code{\link[=removeSheet-methods]{removeSheet}}, \code{\link[=renameSheet-methods]{renameSheet}},
\code{\link[=existsSheet-methods]{existsSheet}}, \code{\link[=getSheets-methods]{getSheets}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
mtcarsFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(mtcarsFile)

# Sets the active sheet to the sheet 'mtcars3'
setActiveSheet(wb, sheet = "mtcars3")
}
\keyword{methods}
\keyword{utilities}
