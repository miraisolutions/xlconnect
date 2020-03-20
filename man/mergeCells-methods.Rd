\name{mergeCells-methods}
\docType{methods}
\alias{mergeCells}
\alias{mergeCells-methods}
\alias{mergeCells,workbook,character-method}
\alias{mergeCells,workbook,numeric-method}
\title{Merging cells}
\description{
Merges cells in a worksheet.
}
\usage{
\S4method{mergeCells}{workbook,character}(object,sheet,reference)
\S4method{mergeCells}{workbook,numeric}(object,sheet,reference)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet on which to merge cells}
  \item{reference}{A cell range specification (\code{character}) in the form 'A1:B8'}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=unmergeCells-methods]{unmergeCells}},
\code{\link{idx2cref}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("mergeCells.xlsx", create = TRUE)

# Create a worksheet called 'merge'
createSheet(wb, name = "merge")

# Merge the cells A1:B8 on the worksheet created above
mergeCells(wb, sheet = "merge", reference = "A1:B8")

# Save workbook
saveWorkbook(wb)

# clean up 
file.remove("mergeCells.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
