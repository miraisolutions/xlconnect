\name{setForceFormulaRecalculation-methods}
\docType{methods}
\alias{setForceFormulaRecalculation}
\alias{setForceFormulaRecalculation-methods}
\alias{setForceFormulaRecalculation,workbook,character-method}
\alias{setForceFormulaRecalculation,workbook,numeric-method}
\title{Forcing Excel to recalculate formula values when opening a workbook}
\description{
 This function controls a flag that forces Excel to recalculate formula values when a workbook
 is opened. 
}
\usage{
\S4method{setForceFormulaRecalculation}{workbook,character}(object,sheet,value)
\S4method{setForceFormulaRecalculation}{workbook,numeric}(object,sheet,value)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet for which to force formula recalculation.
  If \code{sheet = "*"}, the flag is set for all sheets in the \code{\linkS4class{workbook}}.}
  \item{value}{\code{logical} specifying if formula recalculation should be forced or not}
}

\details{
  The arguments \code{sheet} and \code{value} are vectorized such
  that multiple worksheets can be controlled with one method call. 
}

\author{
Thomas Themel\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
  A typical use for this flag is forcing Excel into updating formulas
  that reference cells affected by \code{\link[=writeWorksheet-methods]{writeWorksheet}} 
  or \code{\link[=writeNamedRegion-methods]{writeNamedRegion}}. 
  The exact behavior of Excel when the flag is set depends on version and file format.
}
\seealso{
\code{\linkS4class{workbook}}, 
\code{\link[=getForceFormulaRecalculation-methods]{getForceFormulaRecalculation}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Tell Excel to automatically recalculate formulas on sheet mtcars
setForceFormulaRecalculation(wb, sheet = "mtcars", TRUE)
# The same with a numerical sheet index
setForceFormulaRecalculation(wb, sheet = 1, TRUE)
}
\keyword{methods}
\keyword{utilities}
