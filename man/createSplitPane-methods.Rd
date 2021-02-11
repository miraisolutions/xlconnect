\name{createSplitPane-methods}
\docType{methods}
\alias{createSplitPane}
\alias{createSplitPane-methods}
\alias{createSplitPane,workbook,character-method}
\alias{createSplitPane,workbook,numeric-method}
\title{Creating a split pane on a worksheet}
\description{
Creates a split pane on a specified worksheet.
}
\usage{
\S4method{createSplitPane}{workbook,character}(object,sheet,xSplitPos,ySplitPos,leftColumn,topRow)
\S4method{createSplitPane}{workbook,numeric}(object,sheet,xSplitPos,ySplitPos,leftColumn,topRow)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet on which to create a split pane}
  \item{xSplitPos}{Horizontal position of split (in 1/20th of a point)}
  \item{ySplitPos}{Vertical position of split (in 1/20th of a point)}
  \item{leftColumn}{Left column (as index or column name) visible in right pane}
  \item{topRow}{Top row visible in bottom pane}
}
\references{
How to create a freeze pane/split pane in Office 2007
\url{https://support.microsoft.com/en-us/office/freeze-panes-to-lock-rows-and-columns-dab2ffc9-020d-4026-8121-67dd25f2508f?ocmsassetid=hp001217048&correlationid=b4f5baeb-b622-4487-a96f-514d2f00208a&ui=en-us&rs=en-us&ad=us}
}
\author{
Nicola Lambiase\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
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
\code{\link[=removePane-methods]{removePane}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("splitPaneTest.xlsx", create = TRUE)

# Create a worksheet named 'Sheet1'
createSheet(wb, name = "Sheet1")

# Create a split pane on Sheet1, with coordinates (10000, 5000) expressed as 1/20th of a point,
# 10 (-> J) as left column visible in right pane and 10 as top row visible in bottom pane 
createSplitPane(wb, "Sheet1", 10000, 5000, 10, 10)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)

# clean up 
file.remove("splitPaneTest.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
