\name{getCellFormula-methods}
\docType{methods}
\alias{getCellFormula}
\alias{getCellFormula-methods}
\alias{getCellFormula,workbook,character-method}
\alias{getCellFormula,workbook,numeric-method}
\title{Retrieving formula definitions from cells}
\description{
Retrieves a cell formula from a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{getCellFormula}{workbook,character}(object,sheet,row,col)
\S4method{getCellFormula}{workbook,numeric}(object,sheet,row,col)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the worksheet containing the cell}
  \item{row}{The one-based row index of the cell to query}
  \item{col}{The one-based column index of the cell to query}
}

\details{
Retrieves the formula of the specified cell as a character, without the initial
\code{=} character displayed in Excel. Raises an error if the specified cell 
is not a formula cell.
}

\author{
Thomas Themel\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}

\seealso{
\code{\linkS4class{workbook}}, \code{\link[=setCellFormula-methods]{setCellFormula}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("cellFormula.xlsx", create = TRUE)

createSheet(wb, "Formula")

# Assign a formula to A1
setCellFormula(wb, "Formula", 1, 1, "SUM($B$1:$B$29)")

# Returns the formula for Sheet1!A1
getCellFormula(wb, "Formula", 1, 1)
# The same with a numeric sheet index
getCellFormula(wb, 1, 1, 1)

# clean up 
file.remove("cellFormula.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
