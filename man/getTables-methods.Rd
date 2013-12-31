\name{getTables-methods}
\docType{methods}
\alias{getTables}
\alias{getTables-methods}
\alias{getTables,workbook,numeric-method}
\alias{getTables,workbook,character-method}
\title{Querying available Excel tables in a workbook}
\description{
Queries the available Excel tables on the specified worksheet.
}
\usage{
\S4method{getTables}{workbook,numeric}(object,sheet,simplify)
\S4method{getTables}{workbook,character}(object,sheet,simplify)
}
\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{Index (\code{integer}) or name (\code{character}) of worksheet to query}
  \item{simplify}{\code{logical} specifying if the result should be simplified (defaults
  to \code{TRUE}). See details.}
}
\details{
Since this is a vectorized function (multiple sheets can be specified) the result is a
named list (one component per sheet) if no simplification is applied. In cases where only 
one sheet is queried and \code{simplify = TRUE} (default) the result is simplified to a
vector.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=getSheets-methods]{getSheets}},
\code{\link[=readTable-methods]{readTable}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Query available tables (table names) on sheet 'mtcars_table'
tables <- getTables(wb, sheet = "mtcars_table")

# ... or via sheet index
tables <- getTables(wb, sheet = 4)
}
\keyword{methods}
\keyword{utilities}
