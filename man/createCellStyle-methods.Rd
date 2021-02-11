\name{createCellStyle-methods}
\docType{methods}
\alias{createCellStyle}
\alias{createCellStyle-methods}
\alias{createCellStyle,workbook,character-method}
\alias{createCellStyle,workbook,missing-method}
\title{Creating custom named and anonymous cell styles}
\description{
Creates a custom named or anonymous \code{\linkS4class{cellstyle}}.
}
\usage{
\S4method{createCellStyle}{workbook,character}(object,name)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The name of the new \code{\linkS4class{cellstyle}} to create. Omit
  to create an anonymous \code{\linkS4class{cellstyle}}.}
}

\details{
Creates a named \code{\linkS4class{cellstyle}} with the specified \code{name}. Named cell
styles may be used in conjunction with the \emph{name prefix} style
action (see \code{\link[=setStyleAction-methods]{setStyleAction}}) or
may also be used directly with the method
\code{\link[=setCellStyle-methods]{setCellStyle}}. Named cell styles can
easily be changed from within Excel using the cell styles menu.

If \code{name} is missing, an anonymous cell style is created.
Anonymous cell styles can be used in conjunction with the
\code{\link[=setCellStyle-methods]{setCellStyle}} method.
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\linkS4class{cellstyle}}, \code{\link[=getOrCreateCellStyle-methods]{getOrCreateCellStyle}},
\code{\link[=existsCellStyle-methods]{existsCellStyle}}, \code{\link[=setStyleAction-methods]{setStyleAction}},
\code{\link[=setStyleNamePrefix-methods]{setStyleNamePrefix}}, \code{\link[=setCellStyle-methods]{setCellStyle}},
\code{\link[=setDataFormat-methods]{setDataFormat}}, \code{\link[=setBorder-methods]{setBorder}},
\code{\link[=setFillBackgroundColor-methods]{setFillBackgroundColor}}, \code{\link[=setFillForegroundColor-methods]{setFillForegroundColor}},
\code{\link[=setFillPattern-methods]{setFillPattern}}, \code{\link[=setWrapText-methods]{setWrapText}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("createCellstyles.xlsx", create = TRUE)

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

# clean up 
file.remove("createCellstyles.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
