\name{cellstyle-class}
\Rdversion{1.1}
\docType{class}
\alias{cellstyle-class}
\title{Class "cellstyle"}
\description{
This class represents a cell style in a Microsoft Excel \code{\linkS4class{workbook}}. S4 objects of this class
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
    in the back-end to manipulate the underlying Excel cell style instance.}
  }
}
\references{
Apply, create, or remove a cell style:\cr
\url{https://support.microsoft.com/en-us/office/apply-create-or-remove-a-cell-style-472213bf-66bd-40c8-815c-594f0f90cd22?ocmsassetid=hp001216732&correlationid=5691ac73-b7a2-40c3-99aa-a06e806bb566&ui=en-us&rs=en-us&ad=us} 
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\note{
\pkg{XLConnect} generally makes use of custom (named) cell styles. This allows users to more easily manage cell styles via Excel's cell style
menu. For example, assuming you were using a specific custom cell style for your data table headers, you can change the header
styling with a few clicks in Excel's cell style menu across all tables.
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createCellStyle-methods]{createCellStyle}}, \code{\link[=setStyleAction-methods]{setStyleAction}},
\code{\link[=setCellStyle-methods]{setCellStyle}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("cellstyles.xlsx", create = TRUE)

# We don't set a specific style action in this demo, so the 
# default 'XLConnect' will be used (XLC$"STYLE_ACTION.XLCONNECT")

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

# NOTE: The mtcars data.frame has been written offset with top 
# left cell C4 - and we have also written a header row!
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

# clean up 
file.remove("cellstyles.xlsx")
}
}
\keyword{classes}
\keyword{utilities}
