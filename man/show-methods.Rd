\name{show-methods}
\docType{methods}
\alias{show-methods}
\alias{show,workbook-method}
\title{Display a workbook object}
\description{
 Displays a \code{\linkS4class{workbook}} object by printing it. This actually calls the \code{\linkS4class{workbook}}'s
 \code{\link[=print-methods]{print}} method.
}
\section{Methods}{
\describe{
\item{\code{signature(object = "workbook")}}{
Displays the specified \code{\linkS4class{workbook}} object.
}
}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=print-methods]{print}}
}
\examples{
# Load existing demo Excel file 'mtcars.xlsx' from the XLConnect package
wb.mtcars <- loadWorkbook(system.file("demoFiles/mtcars.xlsx", 
	package = "XLConnect"))
# Display the wb.mtcars object
wb.mtcars
# Alternatively, show can be called explicitely
show(wb.mtcars)
}
\keyword{methods}
\keyword{print}
