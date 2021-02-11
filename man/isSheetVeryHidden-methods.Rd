\name{isSheetVeryHidden-methods}
\docType{methods}
\alias{isSheetVeryHidden}
\alias{isSheetVeryHidden-methods}
\alias{isSheetVeryHidden,workbook,character-method}
\alias{isSheetVeryHidden,workbook,numeric-method}
\title{Checking if worksheets are very hidden in a workbook}
\description{
Checks if the specified worksheets are very hidden (but not just hidden) in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{isSheetVeryHidden}{workbook,character}(object,sheet)
\S4method{isSheetVeryHidden}{workbook,numeric}(object,sheet)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the sheet to check}
}

\details{
Returns \code{TRUE} if the specified named \code{sheet} is very hidden
(not visible but also not just hidden), otherwise
\code{FALSE}. \code{sheet} is vectorized such that multiple worksheets
can be queried with one method call. An exception is thrown if the
specified \code{sheet} does not exist. 
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=hideSheet-methods]{hideSheet}}, \code{\link[=unhideSheet-methods]{unhideSheet}},
\code{\link[=isSheetHidden-methods]{isSheetHidden}}, \code{\link[=isSheetVisible-methods]{isSheetVisible}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("isSheetVeryHidden.xlsx", create = TRUE)

# Write a couple of built-in data.frame's into sheets
# with corresponding name
for(obj in c("CO2", "airquality", "swiss")) {
  createSheet(wb, name = obj)
  writeWorksheet(wb, get(obj), sheet = obj)
}

# Very hide sheet 'airquality'
hideSheet(wb, sheet = "airquality", veryHidden = TRUE)

# Hide sheet 'CO2'
hideSheet(wb, sheet = "CO2", veryHidden = FALSE)

# Check if sheet 'airquality' is very hidden;
# this should obviously return TRUE
isSheetVeryHidden(wb, "airquality")

# Check if sheet 'swiss' is very hidden;
# this should obviously return FALSE
isSheetVeryHidden(wb, "swiss")

# Check if sheet 'CO2' is very hidden;
# this should also return FALSE - the sheet
# is just hidden but not very hidden
isSheetVeryHidden(wb, "CO2")


# clean up 
file.remove("isSheetVeryHidden.xlsx")
}
}
\keyword{methods}
\keyword{utilities}

