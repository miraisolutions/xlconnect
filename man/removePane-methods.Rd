\name{removePane-methods}
\docType{methods}
\alias{removePane}
\alias{removePane-methods}
\alias{removePane,workbook,character-method}
\alias{removePane,workbook,numeric-method}
\title{Removing panes from worksheet}
\description{
Removes the split pane/freeze pane from the specified worksheet.
}
\usage{
\S4method{removePane}{workbook,character}(object,sheet)
\S4method{removePane}{workbook,numeric}(object,sheet)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet from which to remove the split pane/freeze pane}
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
\code{\link[=createFreezePane-methods]{createFreezePane}}
\code{\link[=createSplitPane-methods]{createSplitPane}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("removePaneTest.xlsx", create = TRUE)

# Create a worksheet named 'Sheet1'
createSheet(wb, name = "Sheet1")

# Create a split pane on Sheet1, with coordinates (10000, 5000) expressed as 1/20th of a point,
# 10 (-> J) as left column visible in right pane and 10 as top row visible in bottom pane 
createSplitPane(wb, "Sheet1", 10000, 5000, 10, 10)

# Remove the split pane from Sheet1
removePane(wb, "Sheet1")

# Save workbook (this actually writes the file to disk). Now the workbook has no split pane.
saveWorkbook(wb)

# clean up 
file.remove("removePaneTest.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
