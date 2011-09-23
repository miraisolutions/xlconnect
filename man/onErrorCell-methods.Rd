\name{onErrorCell-methods}
\docType{methods}
\alias{onErrorCell}
\alias{onErrorCell-methods}
\alias{onErrorCell,workbook-method}
\title{ Behavior when error cells are detected }
\description{
This function defines the behavior when reading data from a worksheet and error cells
are detected.
}
\usage{
\S4method{onErrorCell}{workbook}(object,behavior)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{behavior}{The behavior to follow when an error cell is detected. This is normally
  specified by a corresponding \code{\link{XLC}} error constant, i.e. either
  \code{XLC$"ERROR.WARN"} or \code{XLC$"ERROR.STOP"}. \code{XLC$"ERROR.WARN"} means the
  error cell will be read as missing value (\code{NA}) and a corresponding warning will be
  generated (this is the default behavior). \code{XLC$"ERROR.STOP"} means that an exception
  will be thrown and further execution will be stopped immediately.
}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=readNamedRegion-methods]{readNamedRegion}},
\code{\link{readNamedRegionFromFile}}, \code{\link[=readWorksheet-methods]{readWorksheet}},\cr
\code{\link{readWorksheetFromFile}}
}
\examples{
# errorCell xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/errorCell.xlsx", 
							  package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Set error behavior to XLC$ERROR.WARN when detecting error cells
# Note: this is the default behavior
onErrorCell(wb, XLC$ERROR.WARN)
# Alternatively: wb$onErrorCell(XLC$ERROR.WARN)

# Read named region 'MyData' (with default header = TRUE)
data <- readNamedRegion(wb, name = "MyData")

# Now set error behavior to XLC$ERROR.STOP to immediately
# issue an exception and stop in case an error cell is
# detected
onErrorCell(wb, XLC$ERROR.STOP)
# Alternatively: wb$onErrorCell(XLC$ERROR.STOP)

# Read (again) named region 'MyData' (with default header = TRUE)
res <- try(readNamedRegion(wb, name = "MyData"))
# Did we get an error?
print(is(res, "try-error"))
}
\keyword{utilities}
\keyword{error}
