\name{readWorksheet-methods}
\docType{methods}
\alias{readWorksheet-methods}
\title{Reading data from worksheets}
\description{
Reads data from worksheets of a \code{\linkS4class{workbook}} object.
}
\section{Methods}{
\describe{

\item{[]\code{signature(object = "workbook", sheet = "character", startRow = "numeric", 
startCol = "numeric", endRow = "numeric", endCol = "numeric", header = "logical")}}{
Reads data from the worksheet specified by \code{sheet}. Data is read starting at the top left corner specified 
by \code{startRow} and \code{startCol} down to the bottom right corner specified by \code{endRow} and \code{endCol}.
If \code{header = TRUE}, the first row is interpreted as column names of the resulting \code{data.frame}.\cr
If \code{startRow} <= 0 then the first available (logical) row is assumed. If \code{startCol} <= 0 then the column 
of the first (logical) cell of the start row (\code{startRow}) is assumed. If \code{endRow} <= 0 then the last available 
(logical) row is assumed. If \code{endCol} <= 0 then the column of the last (logical) cell of the end column (\code{endCol})
is assumed. In other words, if no boundaries are specified \code{readWorksheet} assumes the "bounding box" of the data as
the corresponding boundaries.
}

\item{[]\code{signature(object = "workbook", sheet = "numeric", startRow = "numeric", startCol = "numeric", endRow = "numeric", endCol = "numeric", header = "logical")}}{
Same as above except that the worksheet (argument \code{sheet}) is specified by index rather than by name.
}

\item{[]\code{signature(object = "workbook", sheet = "character", startRow = "missing", startCol = "missing", endRow = "missing", endCol = "missing", header = "logical")}}{
Same as above with \code{startRow = 0}, \code{startCol = 0}, \code{endRow = 0} and \code{endCol = 0}. In this case 
\code{readWorksheet} assumes the "bounding box" of the data as the corresponding boundaries.
}

\item{[]\code{signature(object = "workbook", sheet = "numeric", startRow = "missing", startCol = "missing", endRow = "missing", endCol = "missing", header = "logical")}}{
Same as above but with the worksheet (argument \code{sheet}) specified by index.
}

\item{[]\code{signature(object = "workbook", sheet = "character", startRow = "missing", startCol = "missing", endRow = "missing", endCol = "missing", header = "missing")}}{
Same as above with \code{header = TRUE}. The worksheet is specified by name.
}

\item{[]\code{signature(object = "workbook", sheet = "numeric", startRow = "missing", startCol = "missing", endRow = "missing", endCol = "missing", header = "missing")}}{
Same as above but with the worksheet specified by index.
}

\item{[]\code{signature(object = "workbook", sheet = "character", startRow = "numeric", startCol = "numeric", endRow = "numeric", endCol = "numeric", header = "missing")}}{
Same as top method with \code{header = TRUE}.
}

\item{[]\code{signature(object = "workbook", sheet = "numeric", startRow = "numeric", startCol = "numeric", endRow = "numeric", endCol = "numeric", header = "missing")}}{
Same as top method with \code{header = TRUE} and the worksheet specified by index.
}
}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
All arguments (except \code{object}) are vectorized. As such, multiple worksheets (and also multiple data regions
from the same worksheet) can be read with one method call. If only one single data region is read, the return value 
is a \code{data.frame}. If multiple data regions are specified, the return value is a \code{list} of \code{data.frame}'s 
returned in the order they have been specified.\cr\cr
\code{readWorksheet} tries to determine the resulting column types based on the read cell types. If different cell
types are found in a specific column, the most general of those is used and mapped to the corresponding R data type.
The order of data types from least to most general is Boolean (\code{logical}) < DateTime (\code{POSIXct}) < 
Numeric (\code{numeric}) < String (\code{character}). E.g. if a column is read that contains cells of type Boolean,
Numeric and String then the resulting column in R would be \code{character} since \code{character} is the most general
type.
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=writeWorksheet-methods]{writeWorksheet}}, 
\code{\link[=readNamedRegion-methods]{readNamedRegion}}, \code{\link[=writeNamedRegion-methods]{writeNamedRegion}} 
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")
# Load workbook
wb <- loadWorkbook(demoExcelFile)
# Read worksheet 'mtcars' (providing no specific area bounds;
# with default header = TRUE)
data <- readWorksheet(wb, sheet = "mtcars")
}
\keyword{methods}
\keyword{IO}
