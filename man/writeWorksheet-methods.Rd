\name{writeWorksheet-methods}
\docType{methods}
\alias{writeWorksheet}
\alias{writeWorksheet-methods}
\alias{writeWorksheet,workbook,ANY,character,missing,missing,logical-method}
\alias{writeWorksheet,workbook,ANY,character,missing,missing,missing-method}
\alias{writeWorksheet,workbook,ANY,character,numeric,numeric,logical-method}
\alias{writeWorksheet,workbook,ANY,character,numeric,numeric,missing-method}
\alias{writeWorksheet,workbook,ANY,numeric,missing,missing,logical-method}
\alias{writeWorksheet,workbook,ANY,numeric,missing,missing,missing-method}
\alias{writeWorksheet,workbook,ANY,numeric,numeric,numeric,logical-method}
\title{Writing data to worksheets}
\description{
Writes data to worksheets of a \code{\linkS4class{workbook}} object.
}
\section{Methods}{
\describe{

\item{[]\code{signature(object = "workbook", data = "ANY", sheet = "character", 
startRow = "numeric", startCol = "numeric", header = "logical")}}{
Writes \code{data} to the (named) worksheet specified by \code{sheet}. \code{data} is assumed to be a 
\code{data.frame} and is coerced to one if this is not already the case.\code{startRow} and \code{startCol} 
define the top left corner of the data region to be written. The argument \code{header} specifies if 
the column names should be written. Note that the arguments are vectorized and as such multiple data objects
(\code{data.frame}'s) can be written to different worksheets in one call. In this case \code{data} is 
assumed to be a list of data objects (\code{data.frame}'s).
}

\item{[]\code{signature(object = "workbook", data = "ANY", sheet = "numeric", 
startRow = "numeric", startCol = "numeric", header = "logical")}}{
Same as above except that the worksheet is specified by an index rather than by name.
}

\item{[]\code{signature(object = "workbook", data = "ANY", 
sheet = "character", startRow = "missing", startCol = "missing", header = "logical")}}{
Same as above with \code{startRow = 1} and \code{startCol = 1} meaning that the top left corner (cell 'A1')
of the worksheet is used as a starting point for writing the data region. \code{sheet} is specified by name.
}

\item{[]\code{signature(object = "workbook", data = "ANY", 
sheet = "numeric", startRow = "missing", startCol = "missing", header = "logical")}}{
Same as above with \code{startRow = 1} and \code{startCol = 1}. \code{sheet} specified by index.
}

\item{[]\code{signature(object = "workbook", data = "ANY", 
sheet = "character", startRow = "numeric", startCol = "numeric", header = "missing")}}{
Same as above with \code{header = TRUE}. \code{sheet} specified by name.
}

\item{[]\code{signature(object = "workbook", data = "ANY", 
sheet = "character", startRow = "missing", startCol = "missing", header = "missing")}}{
Same as above with \code{startRow = 1}, \code{startCol = 1} and \code{header = TRUE}. \code{sheet}
specified by name.
}

\item{[]\code{signature(object = "workbook", data = "ANY", 
sheet = "numeric", startRow = "missing", startCol = "missing", header = "missing")}}{
Same as above with \code{startRow = 1}, \code{startCol = 1} and \code{header = TRUE}. \code{sheet}
specified by index.
}

}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
When writing data to Excel, \code{writeWorksheet} further applies cell styles to the cells as defined by the
\code{\linkS4class{workbook}}'s "style action" (see \code{\link[=setStyleAction-methods]{setStyleAction}}).
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=writeNamedRegion-methods]{writeNamedRegion}},
\code{\link[=readWorksheet-methods]{readWorksheet}}, \code{\link[=readNamedRegion-methods]{readNamedRegion}} 
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
