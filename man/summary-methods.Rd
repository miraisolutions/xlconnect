\name{summary-methods}
\docType{methods}
\alias{summary-methods}
\title{ Summarizing workbook objects }
\description{
 Outputs a \code{\linkS4class{workbook}} summary including the underlying Excel filename, contained worksheets,
 hidden sheets, very hidden sheets, defined names and the active sheet name.
}
\section{Methods}{
\describe{
\item{\code{signature(object = "workbook")}}{
Outputs a summary of the specified \code{\linkS4class{workbook}} object.
}
}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=getSheets-methods]{getSheets}}, 
\code{\link[=isSheetHidden-methods]{isSheetHidden}}, \code{\link[=isSheetVeryHidden-methods]{isSheetVeryHidden}},
\code{\link[=getDefinedNames-methods]{getDefinedNames}}, \code{\link[=getActiveSheetName-methods]{getActiveSheetName}}
}
\examples{
# Load existing demo Excel file 'mtcars.xlsx' from the XLConnect package
wb.mtcars <- loadWorkbook(system.file("demoFiles/mtcars.xlsx", 
	package = "XLConnect"))
# Print a workbook summary
summary(wb.mtcars)
}
\keyword{methods}
\keyword{print}
\keyword{utilities}
