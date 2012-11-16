\name{getBoundingBox-methods}
\docType{methods}
\alias{getBoundingBox}
\alias{getBoundingBox-methods}
\alias{getBoundingBox,workbook,character-method}
\alias{getBoundingBox,workbook,numeric-method}
\title{Querying the coordinates of a worksheet bounding box}
\description{
This function queries the coordinates of a bounding box in an Excel worksheet. A bounding box is 
the rectangular region of minimum size containing all the non-empty cells in a sheet.
The result is a matrix containing the following coordinates:
[1,]    top left row
[2,]    top left column
[3,]	bottom right row
[4,]    bottom right column

In case more than one sheet is selected, the result matrix will contain a column for each sheet.
}
\usage{
\S4method{getBoundingBox}{workbook,character}(object,sheet,startRow,startCol,endRow,endCol)
\S4method{getBoundingBox}{workbook,numeric}(object,sheet,startRow,startCol,endRow,endCol)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet from which to get the bounding box}
  \item{startRow}{Start reference row for the bounding box. If not specified, it defaults to 0.}
  \item{startCol}{Start reference column for the bounding box. If not specified, it defaults to 0.}
  \item{endRow}{End reference row for the bounding box. If not specified, it defaults to 0.}
  \item{endCol}{End reference column for the bounding box. If not specified, it defaults to 0.}
}

\author{
Nicola Lambiase\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=getReferenceCoordinates-methods]{getReferenceCoordinates}}
}
\examples{
# multiregion xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/multiregion.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Query bounding box for the second sheet
print(getBoundingBox(wb, sheet="SecondSheet"))

# Query bounding box for the first sheet, selecting the columns from 5 to 8
print(getBoundingBox(wb, sheet="FirstSheet", startCol=5, endCol=8))
}
\keyword{methods}
\keyword{utilities}