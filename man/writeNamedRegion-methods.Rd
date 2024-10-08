\name{writeNamedRegion-methods}
\docType{methods}
\alias{writeNamedRegion}
\alias{writeNamedRegion-methods}
\alias{writeNamedRegion,workbook-method}
\alias{writeNamedRegion,workbook,ANY-method}
\title{Writing named regions to a workbook}
\description{
Writes data to the named regions defined in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{writeNamedRegion}{workbook,ANY}(object, data, name, header, 
  overwriteFormulaCells, rownames, worksheetScope)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{data}{Data to write}
  \item{name}{Name of the named region to write to}
  \item{header}{Specifies if the column names should be written. The default is \code{TRUE}.}
  \item{overwriteFormulaCells}{Specifies if existing formula cells in the \code{\linkS4class{workbook}}
  should be overwritten. The default is \code{TRUE}.}
  \item{rownames}{Name (\code{character}) of column to use for the row names of the provided
  \code{data} object. If specified, the row names of the data object (\code{data.frame}) will be
  included as an additional column with the specified name. If \code{rownames = NULL} (default),
  no row names will be included. May also be a \code{list} in case multiple data objects are
  written in one call (see below).}
  \item{worksheetScope}{Optional character vector with worksheet name(s) to limit the scope in 
  which the \code{name}(s) to write to is/are expected to be found}.
}

\details{
Writes \code{data} to the named region specified by \code{name}. Note
that \code{data} is assumed to be a \code{data.frame} and is coerced to
one if this is not already the case. The argument \code{header}
specifies if the column names should be written. Note also that the
arguments are vectorized and as such multiple named regions can be
written with one call. In this case \code{data} is assumed to be a list
of data objects (\code{data.frame}'s).
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
\note{
Named regions are automatically redefined to the area occupied by the
written cells. This guarantees that the complete set of data can be
re-read using
\code{\link[=readNamedRegion-methods]{readNamedRegion}}. Also, this
allows the named region just to be defined as the top left cell to be
written to. There is no need to know the exact size of the data in advance.\cr\cr
When writing data to Excel, \code{writeNamedRegion} further applies cell
styles to the cells as defined by the \code{\linkS4class{workbook}}'s
"style action" (see
\code{\link[=setStyleAction-methods]{setStyleAction}}). 
}
\seealso{
  \code{\linkS4class{workbook}},
  \code{\link[=writeWorksheet-methods]{writeWorksheet}},
  \code{\link[=appendNamedRegion-methods]{appendNamedRegion}},
  \code{\link[=appendWorksheet-methods]{appendWorksheet}},
  \code{\link[=readNamedRegion-methods]{readNamedRegion}},
  \code{\link[=readWorksheet-methods]{readWorksheet}},
  \code{\link{writeNamedRegionToFile}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("writeNamedRegion.xlsx", create = TRUE)

# Create a worksheet named 'mtcars'
createSheet(wb, name = "mtcars")

# Create a named region called 'mtcars' on the sheet called 'mtcars'
createName(wb, name = "mtcars", formula = "mtcars!$A$1")

# Write built-in data set 'mtcars' to the above defined named region
# (using header = TRUE)
writeNamedRegion(wb, mtcars, name = "mtcars")

createSheet(wb, name="iris")
setActiveSheet(wb, "iris")

# Do the same with the iris data set, with a worksheet-scoped name
createName(wb, name = "iris", formula = "iris!$A$1", worksheetScope = "iris")
writeNamedRegion(wb, iris, name = "iris", worksheetScope="iris")

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)

# clean up 
file.remove("writeNamedRegion.xlsx")
}
}
\keyword{methods}
\keyword{IO}
