\name{saveWorkbook-methods}
\docType{methods}
\alias{saveWorkbook}
\alias{saveWorkbook-methods}
\alias{saveWorkbook,workbook-method}
\title{Saving Microsoft Excel workbooks}
\description{
 Saves a \code{\linkS4class{workbook}} to the corresponding Excel file. This method actually writes the \code{\linkS4class{workbook}} object to disk.
}
\usage{
\S4method{saveWorkbook}{workbook}(object)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to save}
 }

\details{
Saves the specified \code{\linkS4class{workbook}} object to disk. The \code{\linkS4class{workbook}}'s underlying filename specifies the
path (see the \code{\linkS4class{workbook}} class for more information).
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
As already mentioned in the documentation of the
\code{\linkS4class{workbook}} class,  a \code{\linkS4class{workbook}}'s
underlying Excel file is not saved (or being created in case the file
did not exist and \code{create = TRUE} has been specified) unless the
\code{saveWorkbook} method has been called on the object. This provides
more flexibility to the user to decide when changes are saved and also
provides better performance in that several changes can be written in
one go (normally at the end, rather than  after every operation causing
the file to be rewritten again completely each time). This is due to the
fact that workbooks are manipulated in-memory and are only written to
disk with specifically calling \code{saveWorkbook}.\cr\cr
Further note that calling \code{saveWorkbook} more than once leads to an
exception. This is due to a current issue in the underlying POI libraries.
However, with \pkg{XLConnect} there should be no need to call \code{saveWorkbook}
more than once so virtually this is no issue.
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link{loadWorkbook}}
}
\examples{
# Create a new workbook 'saveMe.xlsx'
# (assuming the file to not exist already)
wb <- loadWorkbook("saveMe.xlsx", create = TRUE)

# Create a worksheet called 'mtcars'
createSheet(wb, name = "mtcars")

# Write built-in dataset 'mtcars' to sheet 'mtcars' created above
writeWorksheet(wb, mtcars, sheet = "mtcars")

# Save workbook - this actually writes the file 'saveMe.xlsx' to disk
saveWorkbook(wb)
}
\keyword{methods}
\keyword{file}
