\name{renameSheet-methods}
\docType{methods}
\alias{renameSheet}
\alias{renameSheet-methods}
\alias{renameSheet,workbook,character,character-method}
\alias{renameSheet,workbook,numeric,character-method}
\title{Renaming worksheets from workbooks}
\description{
Renames a worksheet from a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{renameSheet}{workbook,character,character}(object,sheet,newName)
\S4method{renameSheet}{workbook,numeric,character}(object,sheet,newName)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet to rename}
  \item{newName}{The new name of the sheet}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
Note that renaming worksheets may result in invalid name references.
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createSheet-methods]{createSheet}},
\code{\link[=existsSheet-methods]{existsSheet}}, \code{\link[=getSheets-methods]{getSheets}},
\code{\link[=removeSheet-methods]{removeSheet}}, \code{\link[=cloneSheet-methods]{cloneSheet}},
\code{\link[=setActiveSheet-methods]{setActiveSheet}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
mtcarsFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(mtcarsFile)

# Rename the worksheet called 'mtcars' from the above file to 'MyCars'
renameSheet(wb, sheet = "mtcars", newName = "MyCars")
}
\keyword{methods}
\keyword{utilities}
