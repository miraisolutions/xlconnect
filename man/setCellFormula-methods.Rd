\name{setCellFormula-methods}
\docType{methods}
\alias{setCellFormula}
\alias{setCellFormula-methods}
\alias{setCellFormula,workbook,character-method}
\alias{setCellFormula,workbook,numeric-method}
\title{Setting cell formulas}
\description{
Sets cell formulas for specific cells in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{setCellFormula}{workbook,character}(object,sheet,row,col,formula)
\S4method{setCellFormula}{workbook,numeric}(object,sheet,row,col,formula)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{Name or index of the sheet the cell is on}
  \item{row}{Row index of the cell to edit}
  \item{col}{Column index of the cell to edit}
  \item{formula}{The formula to apply to the cell, without the initial 
                \code{=} character used in Excel}
}

\details{
Note that the arguments are vectorized such that multiple cells can be set with
one method call. 
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
  \code{\linkS4class{workbook}},
  \code{\link[=getCellFormula-methods]{getCellFormula}},
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("setCellFormula.xls", create = TRUE)

# Create a sheet named 'mtcars'
createSheet(wb, name = "mtcars")

# Create a named region called 'mtcars' referring to the sheet
# called 'mtcars'
createName(wb, name = "mtcars", formula = "mtcars!$A$1")

# Write built-in data set 'mtcars' to the above defined named region.
writeNamedRegion(wb, mtcars, name = "mtcars")

# Now, let us get Excel to calculate average weights.
# Where did we write the dataset?
corners <- getReferenceCoordinatesForName(wb, "mtcars")
# Put the average under the wt column
colIndex <- which(names(mtcars) == "wt") 
rowIndex <- corners[2,1] + 1

# Construct the input range & formula
input <- paste(idx2cref(c(corners[1,1], colIndex, 
                          corners[2,1], colIndex)), collapse=":")
formula <- paste("AVERAGE(", input, ")", sep="") 
            
setCellFormula(wb, "mtcars", rowIndex, colIndex, formula)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)

# clean up 
file.remove("setCellFormula.xls")
}
}
\keyword{methods}
\keyword{utilities}
