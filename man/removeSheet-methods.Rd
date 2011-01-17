\name{removeSheet-methods}
\docType{methods}
\alias{removeSheet}
\alias{removeSheet-methods}
\alias{removeSheet,workbook,character-method}
\alias{removeSheet,workbook,numeric-method}
\title{Removing worksheets from workbooks}
\description{
Removes a worksheet from a \code{\linkS4class{workbook}} object.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook", sheet = "character")}}{
Removes the specified (named) \code{sheet}(s) from the workbook.
}

\item{\code{signature(object = "workbook", sheet = "numeric")}}{
Removes the specified (indexed) \code{sheet}(s) from the workbook.
}
}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
When removing a worksheet that is the currently active sheet then \pkg{XLConnect} resets
the active sheet to the first possible worksheet in the \code{\linkS4class{workbook}}.\cr
Also note that deleting worksheets may result in invalid name references.
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createSheet-methods]{createSheet}}, 
\code{\link[=existsSheet-methods]{existsSheet}}, \code{\link[=getSheets-methods]{getSheets}},
\code{\link[=setActiveSheet-methods]{setActiveSheet}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
mtcarsFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(mtcarsFile)

# Remove the worksheet called 'mtcars' from the above file
removeSheet(wb, sheet = "mtcars")
}
\keyword{methods}
\keyword{utilities}
