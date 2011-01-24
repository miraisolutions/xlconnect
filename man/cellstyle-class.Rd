\name{cellstyle-class}
\Rdversion{1.1}
\docType{class}
\alias{cellstyle-class}
\alias{$,cellstyle-method}
\alias{setBorder,cellstyle,character,numeric,numeric-method}
\alias{setCellStyle,workbook,character,numeric,numeric,cellstyle-method}
\alias{setCellStyle,workbook,numeric,numeric,numeric,cellstyle-method}
\alias{setDataFormat,cellstyle,character-method}
\alias{setFillBackgroundColor,cellstyle,numeric-method}
\alias{setFillForegroundColor,cellstyle,numeric-method}
\alias{setFillPattern,cellstyle,numeric-method}
\alias{setWrapText,cellstyle,logical-method}
\title{Class "cellstyle"}
\description{
This class represents a cell style in a Microsoft Excel workbook. S4 objects of this class
and corresponding methods are used to manipulate cell styles. This includes setting data formats,
borders, background- and foreground-colors, etc.
}
\section{Objects from the Class}{
Cell styles are created by calling the \code{\link[=createCellStyle-methods]{createCellStyle}} method on a
\code{\linkS4class{workbook}} object.
}
\section{Slots}{
  \describe{
    \item{\code{jobj}:}{Object of class \code{jobjRef} (see package \pkg{rJava}) which represents a Java object reference that is used
    in the back-end to manipulate the underlying Excel cellstyle instance.}
  }
}
\references{
Cell styles in Excel 2003: \url{http://office.microsoft.com/en-us/excel-help/using-styles-in-excel-HA001110199.aspx}\cr
Cell styles in Excel 2007: \url{http://office.microsoft.com/en-us/excel-help/apply-create-or-remove-a-cell-style-HP001216732.aspx}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
\pkg{XLConnect} generally makes use of custom (named) cell styles. This allows Excel users to more easily manage cell styles
in a workbook. For example, assuming you were using a specific custom cell style for your data table headers, you can change the header
styling with a few clicks across all tables.
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createCellStyle-methods]{createCellStyle}}, \code{\link[=setStyleAction-methods]{setStyleAction}},
\code{\link[=setCellStyle-methods]{setCellStyle}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("cellstyles.xlsx", create = TRUE)

# We don't set a specific style action in this demo, so the default 'XLConnect'
# will be used (XLC$"STYLE_ACTION.XLCONNECT")

# Create a sheet named 'mtcars'
createSheet(wb, name = "mtcars")

# Create a named region called 'mtcars' referring to the sheet called 'mtcars'
createName(wb, name = "mtcars", formula = "mtcars!$C$4")

# Write built-in data set 'mtcars' to the above defined named region.
# This will use the default style action 'XLConnect'.
writeNamedRegion(wb, mtcars, name = "mtcars")

# Now let's color all weight cells of cars with a weight > 3.5 in red
# (mtcars$wt > 3.5)

# First, create a corresponding (named) cellstyle
heavyCar <- createCellStyle(wb, name = "HeavyCar")
# Specify the cellstyle to use a solid foreground color
setFillPattern(heavyCar, fill = XLC$"FILL.SOLID_FOREGROUND")
# Specify the foreground color to be used
setFillForegroundColor(heavyCar, color = XLC$"COLOR.RED")

# Which cars have a weight > 3.5 ?
rowIndex <- which(mtcars$wt > 3.5)

# NOTE: The mtcars data.frame has been written offset with top left cell
# C4 - and we have also written a header row!
# So, let's take that into account appropriately. Obviously, the two steps
# could be combined directly into one ...
rowIndex <- rowIndex + 4

# The same holds for the column index
colIndex <- which(names(mtcars) == "wt") + 2

# Set the 'HeavyCar' cellstyle for the corresponding cells.
# Note: the row and col arguments are vectorized!
setCellStyle(wb, sheet = "mtcars", row = rowIndex, col = colIndex, 
	cellstyle = heavyCar)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
}
\keyword{classes}
\keyword{utilities}
