\name{getCellStyle-methods}
\docType{methods}
\alias{getCellStyle}
\alias{getCellStyle-methods}
\alias{getCellStyle,workbook-method}
\title{Retrieving named cell styles}
\description{
Retrieves a named cell style from a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{getCellStyle}{workbook}(object,name)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The name of the \code{\linkS4class{cellstyle}} to retrieve}
}

\details{
Retrieves the \code{\linkS4class{cellstyle}} with the specified \code{name}. 
}

\author{
Thomas Themel\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}

\seealso{
\code{\linkS4class{workbook}}, \code{\linkS4class{cellstyle}}, \code{\link[=setStyleAction-methods]{setStyleAction}},
\code{\link[=createCellStyle-methods]{createCellStyle}}, \code{\link[=getOrCreateCellStyle-methods]{getOrCreateCellStyle}},
\code{\link[=existsCellStyle-methods]{existsCellStyle}},
\code{\link[=setStyleNamePrefix-methods]{setStyleNamePrefix}}, \code{\link[=setCellStyle-methods]{setCellStyle}},
\code{\link[=setDataFormat-methods]{setDataFormat}}, \code{\link[=setBorder-methods]{setBorder}},
\code{\link[=setFillBackgroundColor-methods]{setFillBackgroundColor}}, \code{\link[=setFillForegroundColor-methods]{setFillForegroundColor}},
\code{\link[=setFillPattern-methods]{setFillPattern}}, \code{\link[=setWrapText-methods]{setWrapText}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("getCellstyles.xlsx", create = TRUE)

# You wouldn't usually ignore the return value here...
createCellStyle(wb, 'Header')

# ... but if you did it doesn't hurt.
cs <- getCellStyle(wb, 'Header')

# Specify the cell style to use a solid foreground color
setFillPattern(cs, fill = XLC$"FILL.SOLID_FOREGROUND")

# Specify the foreground color to be used
setFillForegroundColor(cs, color = XLC$"COLOR.RED")


# clean up 
file.remove("getCellstyles.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
