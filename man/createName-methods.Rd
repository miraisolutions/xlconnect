\name{createName-methods}
\docType{methods}
\alias{createName-methods}
\title{Creating names in a workbook}
\description{
Creates a name for a specified formula in a \code{\linkS4class{workbook}} object.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook", name = "character", formula = "character", overwrite = "logical")}}{
Creates a name named \code{name} for the specified \code{formula}. If a name with the same \code{name} already exists
and \code{overwrite = TRUE}, then this name is removed first before the new one is created. If a name already exists and
\code{overwrite = FALSE}, then an exception is thrown.\cr
The \code{formula} should be specified as you would type it in Excel. Make sure that the referenced worksheets, functions, ... 
exist that you are referring to in the \code{formula}.\cr
The \code{name}, \code{formula} and \code{overwrite} arguments are vectorized such that multiple names can be created in one
method call.
}

\item{\code{signature(object = "workbook", name = "character", formula = "character", overwrite = "missing")}}{
As above with \code{overwrite = FALSE}.
}
}}
\references{
What are named regions/ranges? \url{http://www.officearticles.com/excel/named_ranges_in_microsoft_excel.htm}\cr
How to create named regions/ranges? \url{http://www.youtube.com/watch?v=iAE9a0uRtpM}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=removeName-methods]{removeName}}, \code{\link[=existsName-methods]{existsName}},
\code{\link[=getDefinedNames-methods]{getDefinedNames}},
\code{\link[=readNamedRegion-methods]{readNamedRegion}}, \code{\link[=writeNamedRegion-methods]{writeNamedRegion}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("createName.xlsx", create = TRUE)

# Create a worksheet named 'mtcars'
createSheet(wb, name = "mtcars")

# Create a named region called 'mtcars' on the sheet called 'mtcars'
createName(wb, name = "mtcars", formula = "mtcars!$A$1")

# Write built-in data set 'mtcars' to the above defined named region
writeNamedRegion(wb, mtcars, name = "mtcars")

# Save workbook
saveWorkbook(wb)
}
\keyword{methods}
\keyword{utilities}
