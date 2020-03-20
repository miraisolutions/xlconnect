\name{setDataFormatForType-methods}
\docType{methods}
\alias{setDataFormatForType}
\alias{setDataFormatForType-methods}
\alias{setDataFormatForType,workbook-method}
\title{Setting the data format for the DATA_FORMAT_ONLY style action}
\description{
Sets the data format for a specific data type as used by the DATA_FORMAT_ONLY style action.
}
\usage{
\S4method{setDataFormatForType}{workbook}(object,type,format)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{type}{The data type for which to set the \code{format}.}
  \item{format}{A data format string}
}
\details{
Based on the (cell) data type the DATA_FORMAT_ONLY style action (see \code{\link[=setStyleAction-methods]{setStyleAction}})
sets the data format for the corresponding cells. The data \code{type} is normally specified via a
corresponding data type constant from the \code{\link{XLC}} object. Data formats are specified the standard
Excel way. Refer to the Excel help or to the link below for more information. 
}
\references{
Excel custom data formats:\cr
\url{http://www.ozgrid.com/Excel/CustomFormats.htm}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=setStyleAction-methods]{setStyleAction}}
}
\examples{\dontrun{
# Copy existing Excel template to working directory
file.copy(system.file("demoFiles/template2.xlsx", 
                      package = "XLConnect"),
          "dataformat.xlsx", overwrite = TRUE)

# Load workbook
wb <- loadWorkbook("dataformat.xlsx")

# Set the data format for numeric columns (cells)
# (keeping the defaults for all other data types)
setDataFormatForType(wb, type = XLC$"DATA_TYPE.NUMERIC", 
                     format = "0.00")

# Set style action to 'data format only'
setStyleAction(wb, XLC$"STYLE_ACTION.DATA_FORMAT_ONLY")

# Write built-in data set 'mtcars' to the named region 
# 'mtcars' as defined by the Excel template.
writeNamedRegion(wb, mtcars, name = "mtcars")

# Save workbook
saveWorkbook(wb)

# clean up
file.remove("dataformat.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
