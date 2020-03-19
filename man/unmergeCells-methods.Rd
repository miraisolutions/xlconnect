\name{unmergeCells-methods}
\docType{methods}
\alias{unmergeCells}
\alias{unmergeCells-methods}
\alias{unmergeCells,workbook,character-method}
\alias{unmergeCells,workbook,numeric-method}
\title{Unmerging cells}
\description{
Unmerges cells in a worksheet.
}
\usage{
\S4method{unmergeCells}{workbook,character}(object,sheet,reference)
\S4method{unmergeCells}{workbook,numeric}(object,sheet,reference)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet on which to unmerge cells}
  \item{reference}{A cell range specification (\code{character}) in the form 'A1:B8'. Note that the specification
  must exactly correspond to the range of the merged cells.}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=mergeCells-methods]{mergeCells}},
\code{\link{idx2cref}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("unmergeCells.xlsx", create = TRUE)

# Create a worksheet called 'merge'
createSheet(wb, name = "merge")

# Merge the cells A1:B8 on the worksheet created above
mergeCells(wb, sheet = "merge", reference = "A1:B8")

# Unmerge the cells A1:B8
unmergeCells(wb, sheet = "merge", reference = "A1:B8")

# clean up 
file.remove("unmergeCells.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
