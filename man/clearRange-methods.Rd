\name{clearRange-methods}
\docType{methods}
\alias{clearRange}
\alias{clearRange-methods}
\alias{clearRange,workbook,numeric-method}
\alias{clearRange,workbook,character-method}
\title{Clearing cell ranges in a workbook}
\description{
Clears cell ranges in a \code{\linkS4class{workbook}}.
}
\usage{
  \S4method{clearRange}{workbook,numeric}(object, sheet, coords)
  \S4method{clearRange}{workbook,character}(object, sheet, coords)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the worksheet in which to clear
  cell ranges}
  \item{coords}{Numeric vector of length 4 or numeric matrix with 4 columns
  where the elements of the vector or rows in the matrix refer to
  the coordinates of the top-left and bottom-right corners of the ranges
  to clear. I.e. a vector or each row specifies the coordinates
  \{top row, left column, bottom row, right column\}. You may use
  \code{\link{aref2idx}} to generate such a matrix. }
}

\details{
Clearing a cell range means to remove any cells associated with that range.
}

\author{
Nicola Lambiase\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=clearSheet-methods]{clearSheet}}, 
\code{\link[=clearNamedRegion-methods]{clearNamedRegion}}, 
\code{\link[=clearRangeFromReference-methods]{clearRangeFromReference}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of 
# package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", 
                             package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Clear range from top left corner (4,2) ^= B4 to
# bottom right corner (6,4) ^= D6
clearRange(wb, sheet = "mtcars", coords = c(4, 2, 6, 4))

# Clear two ranges in one go ...
mat = matrix(c(5, 1, 6, 4, 5, 7, 7, 9), ncol = 4,
             byrow = TRUE)
clearRange(wb, sheet = "mtcars", coords = mat)

# The above is equivalent to ...
clearRange(wb, sheet = "mtcars",
           coords = aref2idx(c("A5:D6", "G5:I7")))
           
# This in turn is the same as ...
clearRangeFromReference(wb, reference = c("mtcars!A5:D6",
                        "mtcars!G5:I7"))
}
\keyword{methods}
\keyword{utilities}
