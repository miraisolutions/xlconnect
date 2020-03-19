\name{setBorder-methods}
\docType{methods}
\alias{setBorder}
\alias{setBorder-methods}
\alias{setBorder,cellstyle-method}
\title{Specifying borders for cell styles}
\description{
Specifies borders for a \code{\linkS4class{cellstyle}}.
}
\usage{
\S4method{setBorder}{cellstyle}(object,side,type,color)
}

\arguments{
  \item{object}{The \code{\linkS4class{cellstyle}} to edit}
  \item{side}{A vector with any combination of \{"bottom", "left",
    "right", "top", "all"\}}
  \item{type}{Specifies the border type to be used - it is normally
  specified by a corresponding \code{\link{XLC}} constant (see the
  \code{\link{XLC}} border constant,
  e.g. \code{XLC$"BORDER.MEDIUM_DASHED"})}
  \item{color}{Defines the border color and is normally also specified
    via an \code{\link{XLC}} constant.}
}

\details{
Specifies the border for a \code{\linkS4class{cellstyle}}. Note that the arguments
\code{type} and  \code{color} should be of the same length as side. In
other words, for each specified \code{side} there should be a
corresponding specification of \code{type} and \code{color}. If this is
not the case the arguments will be automatically replicated to the
length of \code{side}.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\linkS4class{cellstyle}}, \code{\link[=setCellStyle-methods]{setCellStyle}}, 
\code{\link[=setStyleAction-methods]{setStyleAction}}, \code{\link{XLC}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("setBorder.xlsx", create = TRUE)

# Create a worksheet
createSheet(wb, name = "cellstyles")

# Create a custom anonymous cell style
cs <- createCellStyle(wb)

# Specify the border for the cell style created above
setBorder(cs, side = c("bottom", "right"), type = XLC$"BORDER.THICK", 
          color = c(XLC$"COLOR.BLACK", XLC$"COLOR.RED"))

# Set the cell style created above for the top left cell (A1) in the 
# 'cellstyles' worksheet
setCellStyle(wb, sheet = "cellstyles", row = 1, col = 1, cellstyle = cs)

# Save the workbook
saveWorkbook(wb)

# clean up 
file.remove("setBorder.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
