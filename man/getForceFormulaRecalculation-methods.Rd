\name{getForceFormulaRecalculation-methods}
\docType{methods}
\alias{getForceFormulaRecalculation}
\alias{getForceFormulaRecalculation-methods}
\alias{getForceFormulaRecalculation,workbook,character-method}
\alias{getForceFormulaRecalculation,workbook,numeric-method}
\title{Querying the coordinates of the range reference by an Excel name}
\description{
Queries the "force formula recalculation" flag on an Excel worksheet.
}
\usage{
\S4method{getForceFormulaRecalculation}{workbook,character}(object,sheet)
\S4method{getForceFormulaRecalculation}{workbook,numeric}(object,sheet)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet to query. This argument is vectorized such that
  multiple sheets can be queried with one method call. If \code{sheet = "*"}, the flag is
  queried for all sheets in the \code{\linkS4class{workbook}} (in the order as returned
  by \code{\link[=getSheets-methods]{getSheets}}).}
}

\author{
Thomas Themel\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=getSheets-methods]{getSheets}},
\code{\link[=setForceFormulaRecalculation-methods]{setForceFormulaRecalculation}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Ask whether Excel will automatically recalculate formulas on sheet mtcars
print(getForceFormulaRecalculation(wb, sheet = "mtcars"))
}
\keyword{methods}
\keyword{utilities}
