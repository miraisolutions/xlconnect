\name{setFillPattern-methods}
\docType{methods}
\alias{setFillPattern}
\alias{setFillPattern-methods}
\alias{setFillPattern,cellstyle,numeric-method}
\title{Specifying the fill pattern for cell styles}
\description{
Specifies the fill pattern for a \code{\linkS4class{cellstyle}}.
}
\usage{
\S4method{setFillPattern}{cellstyle,numeric}(object,fill)
}

\arguments{
  \item{object}{The \code{\linkS4class{cellstyle}} to manipulate}
  \item{fill}{The fill pattern to use for the \code{\linkS4class{cellstyle}}. 
  \code{fill} is normally specified via a corresponding fill constant from the 
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
wb <- loadWorkbook("setFillPattern.xlsx", create = TRUE)

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
