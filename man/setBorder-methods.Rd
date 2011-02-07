\name{setBorder-methods}
\docType{methods}
\alias{setBorder}
\alias{setBorder-methods}
\alias{setBorder,cellstyle,character,numeric,numeric-method}
\title{Specifying borders for cell styles}
\description{
Specifies borders for a \code{\linkS4class{cellstyle}}.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "cellstyle", side = "character", type = "numeric", color = "numeric")}}{
Specifies the border for a cell style. \code{side} must be specified as a vector with any combination of \{"bottom", "left",
"right", "top", "all"\}.
The argument \code{type} specifies the border type to be used - it is normally specified by a corresponding \code{\link{XLC}}
constant (see the \code{\link{XLC}} border constant, e.g. \code{XLC$"BORDER.MEDIUM_DASHED"}). Further, the argument \code{color} 
defines the border color and is normally also specified via an \code{\link{XLC}} constant. Note that the arguments \code{type} and 
\code{color} should be of the same length as side. In other words, for each specified \code{side} there should be a corresponding
specification of \code{type} and \code{color}. If this is not the case the arguments will be automatically replicated to the length
of \code{side}.
}
}}
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
wb <- loadWorkbook("setBorder.xlsx", create = TRUE)
# Create a worksheet
createSheet(wb, name = "cellstyles")
# Create a custom (unnamed) cell style
cs <- createCellStyle(wb)
# Specify the border for the cell style created above
setBorder(cs, side = c("bottom", "right"), type = XLC$"BORDER.THICK", color = c(XLC$"COLOR.BLACK", XLC$"COLOR.RED"))
# Set the cell style created above for the top left cell (A1) in the 'cellstyles' worksheet
setCellStyle(wb, sheet = "cellstyles", row = 1, col = 1, cellstyle = cs)
# Save the workbook
saveWorkbook(wb)
}
\keyword{methods}
\keyword{utilities}