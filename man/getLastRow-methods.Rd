\name{getLastRow-methods}
\docType{methods}
\alias{getLastRow}
\alias{getLastRow-methods}
\alias{getLastRow,workbook,character-method}
\alias{getLastRow,workbook,numeric-method}
\title{Querying the last (non-empty) row on a worksheet}
\description{
Queries the last (non-empty) row on a worksheet.
}
\usage{
\S4method{getLastRow}{workbook,character}(object,sheet)
\S4method{getLastRow}{workbook,numeric}(object,sheet)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet of which to query the last row}
}

\details{
Returns the numeric index of the last non-empty row in the specified worksheet.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Query the last row of the 'mtcars' worksheet
getLastRow(wb, "mtcars")

# Query the last row of the 'mtcars2' worksheet
getLastRow(wb, "mtcars2")

# Query the last row of the 'mtcars3' worksheet
getLastRow(wb, "mtcars3")
}
\keyword{methods}
\keyword{utilities}
