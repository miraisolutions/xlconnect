\name{getSheetPos-methods}
\docType{methods}
\alias{getSheetPos}
\alias{getSheetPos-methods}
\alias{getSheetPos,workbook,character-method}
\title{Querying worksheet position}
\description{
Queries the position of a worksheet in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{getSheetPos}{workbook,character}(object,sheet)
}
\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name of the worksheet (\code{character}) to query. This argument is vectorized such that
  multiple worksheets can be queried with one method call.}
}
\value{
Returns the position index of the corresponding worksheet. Note that querying a non-existing worksheet results
in a 0 index and does not throw an exception!
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=setSheetPos-methods]{setSheetPos}},
\code{\link[=getSheets-methods]{getSheets}}
}
\examples{
\dontrun{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Query worksheet positions for the worksheets 'mtcars2', 'mtcars3',
# 'mtcars' and 'NotThere' (which actually does not exist)
print(getSheetPos(wb, sheet = c("mtcars2", "mtcars3", "mtcars", "NotThere")))
}
}
\keyword{methods}
\keyword{utilities}
