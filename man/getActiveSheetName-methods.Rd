\name{getActiveSheetName-methods}
\docType{methods}
\alias{getActiveSheetName}
\alias{getActiveSheetName-methods}
\alias{getActiveSheetName,workbook-method}
\title{Querying the active worksheet name}
\description{
Queries the name of the active worksheet in a \code{\linkS4class{workbook}} object.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook")}}{
Queries the name of the active worksheet of the provided \code{\linkS4class{workbook}} object.
}
}}
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
