\name{createName-methods}
\docType{methods}
\alias{createName}
\alias{createName-methods}
\alias{createName,workbook,character,character,logical-method}
\alias{createName,workbook,character,character,missing-method}
\title{Creating names in a workbook}
\description{
Creates a name for a specified formula in a \code{\linkS4class{workbook}} object.
}
\usage{
\S4method{createName}{workbook,character,character,logical}(object,name,formula,overwrite)
}

\arguments{
  \item{object}{The workbook to use}
  \item{name}{The new name to create}
  \item{formula}{Excel formula specifying the name}
  \item{overwrite}{If a name with the same \code{name} already exists
    and \code{overwrite = TRUE}, then this name is removed first before
    the new one is created. If a name already exists and
    \code{overwrite = FALSE}, then an exception is thrown.}
}

\details{
Creates a name named \code{name} for the specified \code{formula}. 

The \code{formula} should be specified as you would type it
in Excel. Make sure that the referenced worksheets, functions, ...
exist that you are referring to in the \code{formula}.

The \code{name}, \code{formula} and \code{overwrite} arguments are
vectorized such that multiple names can be created in one method call.

If \code{overwrite} is missing, it is treated as \code{overwrite = FALSE}.
}


\references{
\href{http://www.officearticles.com/excel/named_ranges_in_microsoft_excel.htm}{What
  are named regions/ranges?} 

\href{http://www.youtube.com/watch?v=iAE9a0uRtpM}{How to create named regions/ranges?}
}
\author{
Martin Studer\cr
\href{Mirai Solutions GmbH}{http://www.mirai-solutions.com}
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
