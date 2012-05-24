\name{setSheetColor-methods}
\docType{methods}
\alias{setSheetColor}
\alias{setSheetColor-methods}
\alias{setSheetColor,workbook,character-method,numeric-method}
\alias{setSheetColor,workbook,numeric-method,numeric-method}
\title{Setting colors on worksheet tabs}
\description{
Sets a color on a specified worksheet tab. This only works for xlsx files. 
}
\usage{
\S4method{setSheetColor}{workbook,character,numeric}(object,sheet,color)
\S4method{setSheetColor}{workbook,numeric,numeric}(object,sheet,color)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet on which to set the tab color}
  \item{color}{The color chosen for the sheet tab, as integer or in the form \code{XLC$COLOR.VIOLET}}
}

\author{
Nicola Lambiase\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=XLC]{XLC}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("sheetcolor.xlsx", create = TRUE)

# Create a worksheet named 'Sheet1'
createSheet(wb, name = "Sheet1")

# Set the "Sheet1" tab color as red
setSheetColor(wb, "Sheet1", XLC$COLOR.RED)

# Create a worksheet named 'Sheet2'
createSheet(wb, name = "Sheet2")

# Set the tab color of the second workbook sheet as green
setSheetColor(wb, 2, XLC$COLOR.GREEN)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
}
\keyword{methods}
\keyword{utilities}