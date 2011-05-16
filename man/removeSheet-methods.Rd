\name{removeSheet-methods}
\docType{methods}
\alias{removeSheet}
\alias{removeSheet-methods}
\alias{removeSheet,workbook,character-method}
\alias{removeSheet,workbook,numeric-method}
\title{Removing worksheets from workbooks}
\description{
Removes a worksheet from a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{removeSheet}{workbook,character}(object,sheet)
\S4method{removeSheet}{workbook,numeric}(object,sheet)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet to remove}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
When removing a worksheet that is the currently active sheet then \pkg{XLConnect} resets
the active sheet to the first possible worksheet in the \code{\linkS4class{workbook}}.\cr
Also note that deleting worksheets may result in invalid name references.
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createSheet-methods]{createSheet}},
\code{\link[=existsSheet-methods]{existsSheet}}, \code{\link[=getSheets-methods]{getSheets}},
\code{\link[=renameSheet-methods]{renameSheet}}, \code{\link[=cloneSheet-methods]{cloneSheet}},
\code{\link[=setActiveSheet-methods]{setActiveSheet}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
mtcarsFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(mtcarsFile)

# Remove the worksheet called 'mtcars' from the above file
removeSheet(wb, sheet = "mtcars")
}
\keyword{methods}
\keyword{utilities}
