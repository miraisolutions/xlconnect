\name{setWrapText-methods}
\docType{methods}
\alias{setWrapText}
\alias{setWrapText-methods}
\alias{setWrapText,cellstyle-method}
\title{Specifying text wrapping behaviour}
\description{
Specifies if text should be wrapped in a cell.
}
\usage{
\S4method{setWrapText}{cellstyle}(object,wrap)
}

\arguments{
  \item{object}{The \code{\linkS4class{cellstyle}} to manipulate}
  \item{wrap}{If \code{wrap = TRUE}, the text is wrapped if it exceeds the width of the cell - otherwise not.}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\linkS4class{cellstyle}}, \code{\link[=setCellStyle-methods]{setCellStyle}}, 
\code{\link[=setStyleAction-methods]{setStyleAction}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("setWrapText.xlsx", create = TRUE)

# Create a worksheet
createSheet(wb, name = "cellstyles")

# Create a dummy data set with some long text
text <- data.frame(
     Text = "Some very very very very very very very long text")

# Write the value to the 'cellstyles' worksheet in the 
# top left corner (cell A1)
writeWorksheet(wb, text, sheet = "cellstyles", startRow = 1, 
               startCol = 1, header = FALSE)

# Create a custom anonymous cell style
cs <- createCellStyle(wb)

# Specify to wrap the text
setWrapText(cs, wrap = TRUE)

# Set the cell style created above for the top left cell (A1) 
# in the 'cellstyles' worksheet
setCellStyle(wb, sheet = "cellstyles", row = 1, col = 1, 
             cellstyle = cs)

# Save the workbook
saveWorkbook(wb)
}
\keyword{methods}
\keyword{utilities}
