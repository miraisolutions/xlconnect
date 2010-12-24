\name{isSheetHidden-methods}
\docType{methods}
\alias{isSheetHidden-methods}
\title{Checking if worksheets are hidden in a workbook}
\description{
Checks if the specified worksheets are hidden (but not very hidden) in a \code{\linkS4class{workbook}} object.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook", sheet = "character")}}{
Returns \code{TRUE} if the specified named \code{sheet} is hidden (not visible but also not very hidden),
otherwise \code{FALSE}. \code{sheet} is vectorized such that multiple worksheets can be queried with one
method call. An exception is thrown if the specified \code{sheet} does not exist.
}

\item{\code{signature(object = "workbook", sheet = "numeric")}}{
As above but with \code{sheet} specified by index rather than by name.
}
}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=hideSheet-methods]{hideSheet}}, \code{\link[=unhideSheet-methods]{unhideSheet}},
\code{\link[=isSheetVeryHidden-methods]{isSheetVeryHidden}}, \code{\link[=isSheetVisible-methods]{isSheetVisible}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("isSheetHidden.xlsx", create = TRUE)

# Write a couple of built-in data.frame's into sheets
# with corresponding name
for(obj in c("CO2", "airquality", "swiss")) {
	createSheet(wb, name = obj)
	writeWorksheet(wb, get(obj), sheet = obj)
}

# Hide sheet 'airquality'
hideSheet(wb, sheet = "airquality")

# Check if sheet 'airquality' is hidden;
# this should obviously return TRUE
isSheetHidden(wb, "airquality")

# Check if sheet 'swiss' is hidden;
# this should obviously return FALSE
isSheetHidden(wb, "swiss")

}
\keyword{methods}
\keyword{utilities}
