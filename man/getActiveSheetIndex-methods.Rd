\name{getActiveSheetIndex-methods}
\docType{methods}
\alias{getActiveSheetIndex-methods}
\title{Querying the active worksheet index}
\description{
Queries the index of the active worksheet in a \code{\linkS4class{workbook}} object.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook")}}{
Queries the index of the active worksheet of the provided \code{\linkS4class{workbook}} object.
}
}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=getActiveSheetName-methods]{getActiveSheetName}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")
# Load workbook
wb <- loadWorkbook(demoExcelFile)
# Query the active sheet index
activeSheet <- getActiveSheetIndex(wb)
}
\keyword{methods}
\keyword{utilities}
