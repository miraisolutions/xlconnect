\name{hideSheet-methods}
\docType{methods}
\alias{hideSheet}
\alias{hideSheet-methods}
\alias{hideSheet,workbook,character,logical-method}
\alias{hideSheet,workbook,character,missing-method}
\alias{hideSheet,workbook,numeric,logical-method}
\alias{hideSheet,workbook,numeric,missing-method}
\title{Hiding worksheets in a workbook}
\description{
 (Very) hides the specified worksheets in a \code{\linkS4class{workbook}} object.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook", sheet = "character", veryHidden = "logical")}}{
Hides the specified named \code{sheet}. If \code{veryHidden = TRUE}, the specified \code{sheet} is "very" hidden (see note), otherwise
it is just hidden. The arguments \code{sheet} and \code{veryHidden} are vectorized such that multiple worksheets can be (very)
hidden with one method call. An exception is thrown if the specified \code{sheet} does not exist.
}

\item{\code{signature(object = "workbook", sheet = "character", veryHidden = "missing")}}{
As above with \code{veryHidden = FALSE}.
}

\item{\code{signature(object = "workbook", sheet = "numeric", veryHidden = "logical")}}{
As above but with the \code{sheet} being specified by index rather than by name.
}

\item{\code{signature(object = "workbook", sheet = "numeric", veryHidden = "missing")}}{
As above - \code{sheet} specified by index and \code{veryHidden = FALSE}.
}
}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
Note that hidden worksheets can be unhidden by users directly within Excel via standard functionality. Therefore Excel 
knows the concept of "very hidden" worksheets. These worksheets cannot be unhidden with standard Excel functionality
but need programatic intervention to be made visible.\cr\cr
Also note that in case the specified worksheet to hide is the currently active worksheet, then \code{hideSheet} tries to
set the active worksheet to the first non-hidden (not hidden and not very hidden) worksheet in the workbook. If there is no
such worksheet, \code{hideSheet} will throw a corresponding exception.
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=unhideSheet-methods]{unhideSheet}},
\code{\link[=isSheetHidden-methods]{isSheetHidden}}, \code{\link[=isSheetVeryHidden-methods]{isSheetVeryHidden}},
\code{\link[=isSheetVisible-methods]{isSheetVisible}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("hiddenWorksheet.xlsx", create = TRUE)

# Write a couple of built-in data.frame's into sheets
# with corresponding name
for(obj in c("CO2", "airquality", "swiss")) {
	createSheet(wb, name = obj)
	writeWorksheet(wb, get(obj), sheet = obj)
}

# Hide sheet 'airquality';
# the sheet may be unhidden by a user from within Excel
# since veryHidden defaults to FALSE
hideSheet(wb, sheet = "airquality")
}
\keyword{methods}
\keyword{utilities}
