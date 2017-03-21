\name{setHyperlink-methods}
\docType{methods}
\alias{setHyperlink}
\alias{setHyperlink-methods}
\alias{setHyperlink,workbook,missing,character-method}
\alias{setHyperlink,workbook,missing,numeric-method}
\alias{setHyperlink,workbook,character,missing-method}
\title{Setting hyperlinks}
\description{
Sets hyperlinks for specific cells in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{setHyperlink}{workbook,missing,character}(object,formula,sheet,row,col,type,address)
\S4method{setHyperlink}{workbook,missing,numeric}(object,formula,sheet,row,col,type,address)
\S4method{setHyperlink}{workbook,character,missing}(object,formula,sheet,row,col,type,address)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{formula}{A formula specification in the form Sheet!B8:C17. Use either the argument \code{formula} or the combination of \code{sheet}, \code{row} and \code{col}.}
  \item{sheet}{Name or index of the sheet the cell is on. Use either the argument \code{formula} or the combination of \code{sheet}, \code{row} and \code{col}.}
  \item{row}{Row index of the cell to apply the cellstyle to.}
  \item{col}{Column index of the cell to apply the cellstyle to.}
  \item{type}{Hyperlink type. See the corresponding "HYPERLINK.*" constants from the \code{\link{XLC}} object.}
  \item{address}{Hyperlink address. This needs to be a valid URI including scheme. E.g. for email \code{mailto:myself@me.org}, for a URL \code{http://www.somewhere.net} or for a file \code{file:///a/b/c.dat}}
}

\details{
Sets a hyperlink for the specified cells. Note that \code{\linkS4class{cellstyle}}s for hyperlinks can be defined independently using
\code{\link[=setCellStyle-methods]{setCellStyle}}. The arguments are vectorized such that multiple hyperlinks can be set in one
method call. Use either the argument \code{formula} or the combination of \code{sheet}, \code{row} and \code{col}.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
  \code{\linkS4class{workbook}}, \code{\link[=setCellStyle-methods]{setCellStyle}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("setHyperlink.xlsx", create = TRUE)

# Create a sheet named 'mtcars'
createSheet(wb, name = "mtcars")


# Write built-in data set 'mtcars' to the above defined worksheet
writeWorksheet(wb, mtcars, sheet = "mtcars", rownames = "Car")

# Set hyperlinks
links <- paste0("http://www.google.com?q=", gsub(" ", "+", rownames(mtcars)))
setHyperlink(wb, sheet = "mtcars", row = seq_len(nrow(mtcars)) + 1, col = 1,
  type = XLC$HYPERLINK.URL, address = links)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
}
\keyword{methods}
\keyword{utilities}
