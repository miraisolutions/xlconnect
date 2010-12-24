\name{setActiveSheet-methods}
\docType{methods}
\alias{setActiveSheet-methods}
\title{Setting the active worksheet in a workbook}
\description{
Sets the active worksheet of a \code{\linkS4class{workbook}}.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook", sheet = "character")}}{
Sets the (named) \code{sheet} as the active worksheet.
}

\item{\code{signature(object = "workbook", sheet = "numeric")}}{
Sets the (indexed) \code{sheet} as the active worksheet.
}
}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
The active worksheet of a \code{\linkS4class{workbook}} is the worksheet that is displayed
when the corresponding Excel file is opened.
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createSheet-methods]{createSheet}},
\code{\link[=removeSheet-methods]{removeSheet}}, \code{\link[=existsSheet-methods]{existsSheet}}, 
\code{\link[=getSheets-methods]{getSheets}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
mtcarsFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(mtcarsFile)

# Sets the active sheet to the sheet 'mtcars3'
setActiveSheet(wb, sheet = "mtcars3")
}
\keyword{methods}
\keyword{utilities}
