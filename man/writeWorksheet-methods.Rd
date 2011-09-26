\name{writeWorksheet-methods}
\docType{methods}
\alias{writeWorksheet}
\alias{writeWorksheet-methods}
\alias{writeWorksheet,workbook,ANY,character-method}
\alias{writeWorksheet,workbook,ANY,numeric-method}
\title{Writing data to worksheets}
\description{
Writes data to worksheets of a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{writeWorksheet}{workbook,ANY,character}(object,data,sheet,startRow,startCol,header,rownames)
\S4method{writeWorksheet}{workbook,ANY,numeric}(object,data,sheet,startRow,startCol,header,rownames)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to write to}
  \item{data}{Data to write}
  \item{sheet}{The name or index of the sheet to write to}
  \item{startRow}{Index of the first row to write to. The default is \code{startRow = 1}.}
  \item{startCol}{Index of the first column to write to. The default is \code{startCol = 1}.}
  \item{header}{Specifies if the column names should be written. The default is \code{TRUE}.}
  \item{rownames}{Name (\code{character}) of column to use for the row names of the provided
  \code{data} object. If specified, the row names of the data object (\code{data.frame}) will be
  included as an additional column with the specified name. If \code{rownames = NULL} (default),
  no row names will be included.}
}

\details{
Writes \code{data} to the worksheet specified by
\code{sheet}. \code{data} is assumed to be a  \code{data.frame} and is
coerced to one if this is not already the case. \code{startRow} and
\code{startCol} define the top left corner of the data region to be
written.  Note that the arguments are vectorized and as such multiple
data objects (\code{data.frame}'s) can be written to different
worksheets in one call. In this case \code{data} is assumed to be a \code{list}
of \code{data.frame}s. 
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
When writing data to Excel, \code{writeWorksheet} further applies cell
styles to the cells as defined by the \code{\linkS4class{workbook}}'s
"style action" (see
\code{\link[=setStyleAction-methods]{setStyleAction}}). 
}
\seealso{
  \code{\linkS4class{workbook}},
  \code{\link[=writeNamedRegion-methods]{writeNamedRegion}},
  \code{\link[=appendWorksheet-methods]{appendWorksheet}},
  \code{\link[=appendNamedRegion-methods]{appendNamedRegion}},
  \code{\link[=readWorksheet-methods]{readWorksheet}},
  \code{\link[=readNamedRegion-methods]{readNamedRegion}},
  \code{\link{writeWorksheetToFile}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("writeWorksheet.xlsx", create = TRUE)

# Create a worksheet called 'CO2'
createSheet(wb, name = "CO2")

# Write built-in data set 'CO2' to the worksheet created above;
# offset from the top left corner and with default header = TRUE
writeWorksheet(wb, CO2, sheet = "CO2", startRow = 4, startCol = 2)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
}
\keyword{methods}
\keyword{IO}
