\name{getLastColumn-methods}
\docType{methods}
\alias{getLastColumn}
\alias{getLastColumn-methods}
\alias{getLastColumn,workbook,character-method}
\alias{getLastColumn,workbook,numeric-method}
\title{Querying the last (non-empty) column on a worksheet}
\description{
Queries the last (non-empty) column on a worksheet.
}
\usage{
\S4method{getLastColumn}{workbook,character}(object,sheet)
\S4method{getLastColumn}{workbook,numeric}(object,sheet)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet of which to query the last column}
}

\details{
Returns the (1-based) numeric index of the last non-empty column in the specified worksheet.
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
getLastColumn(wb, "mtcars")

# Query the last row of the 'mtcars2' worksheet
getLastColumn(wb, "mtcars2")

# Query the last row of the 'mtcars3' worksheet
getLastColumn(wb, "mtcars3")
}
\keyword{methods}
\keyword{utilities}
