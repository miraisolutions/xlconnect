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
  \item{leftColumn}{Left column (as index) visible in right pane}
  \item{topRow}{Top row visible in bottom pane}
}
\references{
How to create a freeze pane/split pane in Office 2007
\url{http://office.microsoft.com/en-us/excel-help/freeze-or-lock-rows-and-columns-HP001217048.aspx}
}
\author{
Nicola Lambiase\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}
\code{\link[=createFreezePane-methods]{createFreezePane}}
\code{\link[=removePane-methods]{removePane}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("splitPaneTest.xlsx", create = TRUE)

# Create a worksheet named 'Sheet1'
createSheet(wb, name = "Sheet1")

# Create a split pane on Sheet1, with coordinates (10000, 5000) expressed as 1/20th of a point, 10 (-> J) as left column visible in right pane and 10 as top row visible in bottom pane 
createSplitPane(wb, "Sheet1", 10000, 5000, 10, 10)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
}
\keyword{methods}
\keyword{utilities}