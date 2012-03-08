\name{saveWorkbook-methods}
\docType{methods}
\alias{saveWorkbook}
\alias{saveWorkbook-methods}
\alias{saveWorkbook,workbook,missing-method}
\alias{saveWorkbook,workbook,character-method}
\title{Saving Microsoft Excel workbooks}
\description{
 Saves a \code{\linkS4class{workbook}} to the corresponding Excel file. This method actually writes the \code{\linkS4class{workbook}} object to disk.
}
\usage{
\S4method{saveWorkbook}{workbook,missing}(object,file)
\S4method{saveWorkbook}{workbook,character}(object,file)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to save}
  \item{file}{The file to which to save the \code{\linkS4class{workbook}} ("save as").
  If not specified (missing), the workbook will be saved to the \code{\linkS4class{workbook}}'s underlying 
  file which is the file specified in \code{\link{loadWorkbook}} (also see the \code{\linkS4class{workbook}} 
  class for more information). Note that due to currently missing functionality in Apache POI, workbooks can
  only be saved in the same file format - i.e. if the workbooks underlying file format is xls, then the 
  \code{file} argument may only specify another xls file. Also note that when specifying the \code{file} 
  argument the \code{\linkS4class{workbook}}'s underlying filename changes to reflect the "save as" behavior.\cr
  Paths are expanded using \code{path.expand}.
 }
}

\details{
Saves the specified \code{\linkS4class{workbook}} object to disk.
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
