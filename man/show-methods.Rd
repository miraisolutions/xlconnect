\name{show-methods}
\docType{methods}
\alias{show}
\alias{show-methods}
\alias{show,workbook-method}
\title{Display a workbook object}
\description{
 Displays a \code{\linkS4class{workbook}} by printing it. This actually calls the \code{\linkS4class{workbook}}'s
 \code{\link[=print-methods]{print}} method.
}
\usage{
\S4method{show}{workbook}(object)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to display}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=print-methods]{print}}
}
\examples{
\dontrun{
# Load existing demo Excel file 'mtcars.xlsx' from the XLConnect package
wb.mtcars <- loadWorkbook(system.file("demoFiles/mtcars.xlsx", 
                          package = "XLConnect"))

# Display the wb.mtcars object
wb.mtcars

# Alternatively, show can be called explicitely
show(wb.mtcars)
}
}
\keyword{methods}
\keyword{print}
