\name{getActiveSheetIndex-methods}
\docType{methods}
\alias{getActiveSheetIndex}
\alias{getActiveSheetIndex-methods}
\alias{getActiveSheetIndex,workbook-method}
\title{Querying the active worksheet index}
\description{
Queries the index of the active worksheet in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{getActiveSheetIndex}{workbook}(object)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=getActiveSheetName-methods]{getActiveSheetName}}
}
\examples{
\dontrun{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Query the active sheet index
activeSheet <- getActiveSheetIndex(wb)
}
}
\keyword{methods}
\keyword{utilities}
