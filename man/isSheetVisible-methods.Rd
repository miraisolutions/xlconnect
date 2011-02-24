\name{isSheetVisible-methods}
\docType{methods}
\alias{isSheetVisible}
\alias{isSheetVisible-methods}
\alias{isSheetVisible,workbook,character-method}
\alias{isSheetVisible,workbook,numeric-method}
\title{Checking if worksheets are visible in a workbook}
\description{
Checks if the specified worksheets are visible in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{isSheetVisible}{workbook,character}(object,sheet)
\S4method{isSheetVisible}{workbook,numeric}(object,sheet)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet to check}
}

\details{
Returns \code{TRUE} if the specified named \code{sheet} is visible (not
hidden and not very hidden), otherwise \code{FALSE}. \code{sheet} is
vectorized such that multiple worksheets can be queried with one method
call. An exception is thrown if the specified \code{sheet} does not
exist.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=hideSheet-methods]{hideSheet}}, \code{\link[=unhideSheet-methods]{unhideSheet}},
\code{\link[=isSheetHidden-methods]{isSheetHidden}}, \code{\link[=isSheetVeryHidden-methods]{isSheetVeryHidden}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("isSheetVisible.xlsx", create = TRUE)

# Write a couple of built-in data.frame's into sheets
# with corresponding name
for(obj in c("CO2", "airquality", "swiss")) {
  createSheet(wb, name = obj)
  writeWorksheet(wb, get(obj), sheet = obj)
}

# Hide sheet 'CO2'
hideSheet(wb, sheet = "CO2", veryHidden = FALSE)

# Very hide sheet 'airquality'
hideSheet(wb, sheet = "airquality", veryHidden = TRUE)

# Check if sheet 'swiss' is visible;
# this should obviously return TRUE
isSheetVisible(wb, "swiss")

# Check if sheet 'CO2' is visible;
# this should obviously return FALSE
isSheetVisible(wb, "CO2")

# Check if sheet 'airquality' is visible;
# this should obviously return FALSE
isSheetVisible(wb, "airquality")

}
\keyword{methods}
\keyword{utilities}
