\name{setCellStyleForType-methods}
\docType{methods}
\alias{setCellStyleForType}
\alias{setCellStyleForType-methods}
\alias{setCellStyleForType,workbook-method}
\title{Setting the cell style per data type for the DATATYPE style action}
\description{
Sets the cell style for a specific data type as used by the DATATYPE style action.
}
\usage{
\S4method{setCellStyleForType}{workbook}(object,type,style)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{type}{The data type for which to set the \code{style}}
  \item{style}{The \code{\linkS4class{cellstyle}} to set}
}
\details{
Based on the (cell) data type the DATATYPE style action (see \code{\link[=setStyleAction-methods]{setStyleAction}})
sets the \code{\linkS4class{cellstyle}} for the corresponding cells. The data \code{type} is normally specified via a
corresponding data type constant from the \code{\link{XLC}} object. 
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=getCellStyleForType-methods]{getCellStyleForType}},
\code{\link[=setStyleAction-methods]{setStyleAction}}
}
\examples{
file.copy(system.file("demoFiles/template2.xlsx", 
                      package = "XLConnect"),
          "datatype.xlsx", overwrite = TRUE)

# Load workbook
wb <- loadWorkbook("datatype.xlsx")

# Create a new cell style to be used
cs <- createCellStyle(wb, name = "mystyle")

# Set data format (number format) as numbers with aligned fractions
setDataFormat(cs, format = "# ???/???")

# Define the above created cell style as style to be used for
# numerics
setCellStyleForType(wb, type = XLC$"DATA_TYPE.NUMERIC", style = cs)
# Could also say cs <- setCellStyleForType(wb, "numeric")
          
# Set style action to 'datatype'
setStyleAction(wb, XLC$"STYLE_ACTION.DATATYPE")

# Write built-in data set 'mtcars' to the named region 
# 'mtcars' as defined by the Excel template.
writeNamedRegion(wb, mtcars, name = "mtcars")

# Save workbook
saveWorkbook(wb)
}
\keyword{methods}
\keyword{utilities}
