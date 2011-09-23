\name{setCellStyle-methods}
\docType{methods}
\alias{setCellStyle}
\alias{setCellStyle-methods}
\alias{setCellStyle,workbook,character-method}
\alias{setCellStyle,workbook,numeric-method}
\title{Setting cell styles}
\description{
Sets cell styles for specific cells in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{setCellStyle}{workbook,character}(object,sheet,row,col,cellstyle)
\S4method{setCellStyle}{workbook,numeric}(object,sheet,row,col,cellstyle)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{Name or index of the sheet the cell is on}
  \item{row}{Row index of the cell to edit}
  \item{col}{Column index of the cell to edit}
  \item{cellstyle}{\code{\linkS4class{cellstyle}} to apply to the cell}
}

\details{
Sets the specified \code{\linkS4class{cellstyle}} for the specified cell
(\code{row}, \code{col}) on the specified \code{sheet}. Note that the
arguments are vectorized such that multiple cells can be styled with one
method call. 
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
  \code{\linkS4class{workbook}},
  \code{\linkS4class{cellstyle}},
  \code{\link[=createCellStyle-methods]{createCellStyle}},
  \code{\link[=setDataFormat-methods]{setDataFormat}},
  \code{\link[=setBorder-methods]{setBorder}},\cr
  \code{\link[=setFillBackgroundColor-methods]{setFillBackgroundColor}},
  \code{\link[=setFillForegroundColor-methods]{setFillForegroundColor}},
  \code{\link[=setFillPattern-methods]{setFillPattern}},\cr
  \code{\link[=setWrapText-methods]{setWrapText}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("setCellStyle.xlsx", create = TRUE)

# We don't set a specific style action in this demo, so the default 
# 'XLConnect' will be used (XLC$"STYLE_ACTION.XLCONNECT")

# Create a sheet named 'mtcars'
createSheet(wb, name = "mtcars")

# Create a named region called 'mtcars' referring to the sheet
# called 'mtcars'
createName(wb, name = "mtcars", formula = "mtcars!$C$4")

# Write built-in data set 'mtcars' to the above defined named region.
# This will use the default style action 'XLConnect'.
writeNamedRegion(wb, mtcars, name = "mtcars")

# Now let's color all weight cells of cars with a weight > 3.5 in red
# (mtcars$wt > 3.5)

# First, create a corresponding (named) cell style
heavyCar <- createCellStyle(wb, name = "HeavyCar")

# Specify the cell style to use a solid foreground color
setFillPattern(heavyCar, fill = XLC$"FILL.SOLID_FOREGROUND")

# Specify the foreground color to be used
setFillForegroundColor(heavyCar, color = XLC$"COLOR.RED")

# Which cars have a weight > 3.5 ?
rowIndex <- which(mtcars$wt > 3.5)

# NOTE: The mtcars data.frame has been written offset with
# top left cell C4 - and we have also written a header row!
# So, let's take that into account appropriately. Obviously,
# the two steps could be combined directly into one ...
rowIndex <- rowIndex + 4

# The same holds for the column index
colIndex <- which(names(mtcars) == "wt") + 2

# Set the 'HeavyCar' cell style for the corresponding cells.
# Note: the row and col arguments are vectorized!
setCellStyle(wb, sheet = "mtcars", row = rowIndex, col = colIndex, 
             cellstyle = heavyCar)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
}
\keyword{methods}
\keyword{utilities}
