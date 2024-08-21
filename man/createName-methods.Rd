\name{createName-methods}
\docType{methods}
\alias{createName}
\alias{createName-methods}
\alias{createName,workbook-method}
\title{Creating names in a workbook}
\description{
Creates a named range for a specified formula in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{createName}{workbook}(object, name, formula, overwrite, worksheetScope)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The name of the range to be created}
  \item{formula}{Excel formula specifying the value / data the name refers to}
  \item{overwrite}{If a name with the same \code{name} already exists
    and \code{overwrite = TRUE}, then this name is removed first before
    the new one is created. If a name already exists and
    \code{overwrite = FALSE}, then an exception is thrown. The default
    value for \code{overwrite} is \code{FALSE}.}
  \item{worksheetScope}{Optional - specific worksheet the name should be scoped to. 
    If unspecified the name will be scoped to the whole workbook.}
}

\details{
Creates a named range called \code{name} for the specified \code{formula}. 

The \code{formula} should be specified as you would type it
in Excel. Make sure that the worksheets, functions, ...
exist that you are referring to in the \code{formula}.

The \code{name}, \code{formula} and \code{overwrite} arguments are
vectorized such that multiple names can be created in one method call.
}

\references{
What are named regions/ranges?\cr
\url{https://web.archive.org/web/20240821110221/https://www.officearticles.com/excel/named_ranges_in_microsoft_excel.htm}\cr
How to create named regions/ranges?\cr
\url{https://www.youtube.com/watch?v=iAE9a0uRtpM}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}},
\code{\link[=removeName-methods]{removeName}},
\code{\link[=existsName-methods]{existsName}},
\code{\link[=getDefinedNames-methods]{getDefinedNames}},\cr
\code{\link[=readNamedRegion-methods]{readNamedRegion}},
\code{\link[=writeNamedRegion-methods]{writeNamedRegion}}
}
\examples{\dontrun{
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

# clean up 
file.remove("createName.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
