\name{createSheet-methods}
\docType{methods}
\alias{createSheet}
\alias{createSheet-methods}
\alias{createSheet,workbook-method}
\title{Creating worksheets in a workbook}
\description{
Creates worksheets with specified names in a \code{\linkS4class{workbook}}.
}
\usage{
  \S4method{createSheet}{workbook}(object, name)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The name of the sheet to create}
}

\details{
Creates a worksheet with the specified \code{name} if it does not
already exist. Note that the naming of worksheets needs to be in line
with Excel's convention, otherwise an exception will be thrown. For
example, worksheet names cannot be longer than 31 characters. Also note
that the \code{name} argument is vectorized, so multiple worksheets can
be created in one method call.
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=removeSheet-methods]{removeSheet}}, 
\code{\link[=renameSheet-methods]{renameSheet}},
\code{\link[=existsSheet-methods]{existsSheet}}, \code{\link[=getSheets-methods]{getSheets}},
\code{\link[=cloneSheet-methods]{cloneSheet}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("createSheet.xlsx", create = TRUE)

# Create a worksheet called 'CO2'
createSheet(wb, name = "CO2")

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)

# clean up 
file.remove("createSheet.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
