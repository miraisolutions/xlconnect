\name{createFreezePane-methods}
\docType{methods}
\alias{createFreezePane}
\alias{createFreezePane-methods}
\alias{createFreezePane,workbook,character-method}
\alias{createFreezePane,workbook,numeric-method}
\title{Creating a freeze pane on a worksheet}
\description{
Creates a freeze pane on a specified worksheet.
}
\usage{
\S4method{createFreezePane}{workbook,character}(object, sheet, colSplit, rowSplit, leftColumn, topRow)
\S4method{createFreezePane}{workbook,numeric}(object, sheet, colSplit, rowSplit, leftColumn, topRow)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet on which to create a freeze pane}
  \item{colSplit}{Horizontal position of freeze (as column index or name)}
  \item{rowSplit}{Vertical position of freeze (as number of rows)}
  \item{leftColumn}{Left column (as column index or name) visible in right pane. If not specified, the default is \code{leftColumn}=\code{colSplit}}
  \item{topRow}{Top row (as index) visible in bottom pane. If not specified, the default is \code{topRow}=\code{rowSplit}}
}
\references{
How to create a freeze pane/split pane in Office 2007
\url{http://office.microsoft.com/en-us/excel-help/freeze-or-lock-rows-and-columns-HP001217048.aspx}
}
\author{
Nicola Lambiase\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
To keep an area of a worksheet visible while you scroll to another area of the worksheet,
you can lock specific rows or columns in one area by freezing or splitting panes.

When you freeze panes, you keep specific rows or columns visible when you scroll in the worksheet.
For example, you might want to keep row and column labels visible as you scroll.

When you split panes, you create separate worksheet areas that you can scroll within,
while rows or columns in the non-scrolled area remain visible.
}
\seealso{
\code{\linkS4class{workbook}}
\code{\link[=createSplitPane-methods]{createSplitPane}}
\code{\link[=removePane-methods]{removePane}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("freezePaneTest.xlsx", create = TRUE)

# Create a worksheet named 'Sheet1'
createSheet(wb, name = "Sheet1")

# Create a freeze pane on Sheet1, using as reference position the 5th column and the 5th row,
# showing the 10th column as the leftmost visible one in the right pane
# and the 10th row as the top visible one in the bottom pane.
createFreezePane(wb, "Sheet1", 5, 5, 10, 10)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)

# clean up 
file.remove("freezePaneTest.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
