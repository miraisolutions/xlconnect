\name{setFontColor-methods}
\docType{methods}
\alias{setFontColor}
\alias{setFontColor-methods}
\alias{setFontColor,cellstyle,numeric-method}
\title{Specifying the font color for cell styles}
\description{
Specifies the font color for a \code{\linkS4class{cellstyle}}.
}
\usage{
\S4method{setFontColor}{cellstyle,numeric}(object,color)
}

\arguments{
  \item{object}{The \code{\linkS4class{cellstyle}} to manipulate}
  \item{color}{The font color to use for the \code{\linkS4class{cellstyle}}.
  The \code{color} is normally specified via a corresponding color constant from the
  \code{\link{XLC}} object. It can also receive hex values. }
}

\author{
Matt Deppe\cr
}
\seealso{
\code{\linkS4class{workbook}}, \code{\linkS4class{cellstyle}}, \code{\link[=setCellStyle-methods]{setCellStyle}},
\code{\link[=setStyleAction-methods]{setStyleAction}}, \code{\link{XLC}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("setFontColor.xlsx", create = TRUE)

# Create a worksheet
createSheet(wb, name = "cellstyles")

# Create a custom anonymous cell style
cs <- createCellStyle(wb)

# Specify the font color for the cell style created above
setFontColor(cs, color = XLC$"COLOR.CORNFLOWER_BLUE")

# Specify the font color for the cell style create above
setFontColor(cs, color = "#FF0000")


# Set the cell style created above for the top left cell (A1) in the
# 'cellstyles' worksheet
setCellStyle(wb, sheet = "cellstyles", row = 1, col = 1, cellstyle = cs)

# Save the workbook
saveWorkbook(wb)

# clean up
file.remove("setFontColor.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
