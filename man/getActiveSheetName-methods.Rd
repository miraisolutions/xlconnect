\name{getActiveSheetName-methods}
\docType{methods}
\alias{getActiveSheetName}
\alias{getActiveSheetName-methods}
\alias{getActiveSheetName,workbook-method}
\title{Querying the active worksheet name}
\description{
Queries the name of the active worksheet in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{getActiveSheetName}{workbook}(object)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=getActiveSheetIndex-methods]{getActiveSheetIndex}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Query the active sheet name
activeSheet <- getActiveSheetName(wb)
}
\keyword{methods}
\keyword{utilities}
