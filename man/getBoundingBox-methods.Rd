\name{getBoundingBox-methods}
\docType{methods}
\alias{getBoundingBox}
\alias{getBoundingBox-methods}
\alias{getBoundingBox,workbook,character-method}
\alias{getBoundingBox,workbook,numeric-method}
\title{Querying the coordinates of a worksheet bounding box}
\description{
This function queries the coordinates of a bounding box in an Excel worksheet. A bounding box is the rectangular region of minimum size containing all the non-empty cells in a sheet.
}
\usage{
\S4method{getBoundingBox}{workbook,character}(object,sheet,startRow,startCol,endRow,endCol,autofitRow,autofitCol)
\S4method{getBoundingBox}{workbook,numeric}(object,sheet,startRow,startCol,endRow,endCol,autofitRow,autofitCol)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet from which to get the bounding box}
  \item{startRow}{Start reference row for the bounding box. Defaults to \code{0} meaning that the start row is determined automatically.}
  \item{startCol}{Start reference column for the bounding box. Defaults to \code{0} meaning that the start column is determined automatically.}
  \item{endRow}{End reference row for the bounding box. Defaults to \code{0} meaning that the end row is determined automatically.}
  \item{endCol}{End reference column for the bounding box. Defaults to \code{0} meaning that the end column is determined automatically.}
  \item{autofitRow}{\code{logical} specifying if leading and trailing empty rows should be
  skipped. Defaults to \code{TRUE}.}
  \item{autofitCol}{\code{logical} specifying if leading and trailing empty columns should be skipped. Defaults to \code{TRUE}.}
}
\details{
The result is a \code{matrix} containing the following coordinates:\cr
\code{[1,]}  top left row\cr
\code{[2,]}  top left column\cr
\code{[3,]}  bottom right row\cr
\code{[4,]}  bottom right column\cr
\cr
In case more than one \code{sheet} is selected, the result \code{matrix} will contain a column for each sheet.\cr\cr

The bounding box resolution algorithm works as follows:\cr
If \code{startRow <= 0} then the first available row in the sheet is assumed. 
If \code{endRow <= 0} then the last available row in the sheet is assumed.
If \code{startCol <= 0} then the minimum column between \code{startRow} and
\code{endRow} is assumed. If \code{endCol <= 0} then the maximum column between
\code{startRow} and \code{endRow} is assumed.
The arguments \code{autofitRow} and \code{autofitCol} (both defaulting to \code{TRUE}) can be used to skip leading and trailing empty rows even in case \code{startRow}, \code{endRow}, \code{startCol} and \code{endCol} are specified to values \code{> 0}. This can be useful if data is expected within certain given boundaries but the exact location is not available.
}
\author{
Nicola Lambiase\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}
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