\name{setFillBackgroundColor-methods}
\docType{methods}
\alias{setFillBackgroundColor}
\alias{setFillBackgroundColor-methods}
\alias{setFillBackgroundColor,cellstyle,numeric-method}
\title{Specifying the fill background color for cell styles}
\description{
Specifies the fill background color for a \code{\linkS4class{cellstyle}}.
}
\usage{
\S4method{setFillBackgroundColor}{cellstyle,numeric}(object,color)
}

\arguments{
  \item{object}{The \code{\linkS4class{cellstyle}} to manipulate}
  \item{color}{The fill background color to use for the \code{\linkS4class{cellstyle}}. 
  The \code{color} is normally specified via a corresponding color constant from the 
  \code{\link{XLC}} object. }
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\linkS4class{cellstyle}}, \code{\link[=setCellStyle-methods]{setCellStyle}}, 
\code{\link[=setStyleAction-methods]{setStyleAction}}, \code{\link{XLC}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("setFillBackgroundColor.xlsx", create = TRUE)

# Create a worksheet
createSheet(wb, name = "cellstyles")

# Create a custom anonymous cell style
cs <- createCellStyle(wb)

# Specify the fill background color for the cell style created above
setFillBackgroundColor(cs, color = XLC$"COLOR.CORNFLOWER_BLUE")

# Specify the fill foreground color
setFillForegroundColor(cs, color = XLC$"COLOR.YELLOW")

# Specify the fill pattern
setFillPattern(cs, fill = XLC$"FILL.BIG_SPOTS")

# Set the cell style created above for the top left cell (A1) in the 
# 'cellstyles' worksheet
setCellStyle(wb, sheet = "cellstyles", row = 1, col = 1, cellstyle = cs)

# Save the workbook
saveWorkbook(wb)
}
\keyword{methods}
\keyword{utilities}
