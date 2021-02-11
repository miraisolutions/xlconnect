\name{getCellStyleForType-methods}
\docType{methods}
\alias{getCellStyleForType}
\alias{getCellStyleForType-methods}
\alias{getCellStyleForType,workbook-method}
\title{Querying the cell style per data type for the DATATYPE style action}
\description{
Queries the cell style for a specific data type as used by the DATATYPE style action.
}
\usage{
\S4method{getCellStyleForType}{workbook}(object,type)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{type}{The data type for which to get the \code{\linkS4class{cellstyle}}.}
}
\details{
Based on the (cell) data type the DATATYPE style action (see \code{\link[=setStyleAction-methods]{setStyleAction}})
sets the \code{\linkS4class{cellstyle}} for the corresponding cells. The data \code{type} is normally specified via a
corresponding data type constant from the \code{\link{XLC}} object. 
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=setCellStyleForType-methods]{setCellStyleForType}},
\code{\link[=setStyleAction-methods]{setStyleAction}}
}
\examples{\dontrun{
file.copy(system.file("demoFiles/template2.xlsx", 
                      package = "XLConnect"),
          "datatype.xlsx", overwrite = TRUE)

# Load workbook
wb <- loadWorkbook("datatype.xlsx")

# Get current (existing) cell style for numerics
cs <- getCellStyleForType(wb, XLC$"DATA_TYPE.NUMERIC")
# Could also say cs <- getCellStyleForType(wb, "numeric")

# Change style
setBorder(cs, side = c("bottom", "right"), type = XLC$"BORDER.THICK", 
          color = c(XLC$"COLOR.BLACK", XLC$"COLOR.RED"))
          
# Set style action to 'datatype'
setStyleAction(wb, XLC$"STYLE_ACTION.DATATYPE")

# Write built-in data set 'mtcars' to the named region 
# 'mtcars' as defined by the Excel template.
writeNamedRegion(wb, mtcars, name = "mtcars")

# Save workbook
saveWorkbook(wb)

# clean up
file.remove("datatype.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
