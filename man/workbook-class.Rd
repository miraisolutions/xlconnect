\name{workbook-class}
\Rdversion{1.1}
\docType{class}
\alias{workbook-class}
\alias{$,workbook-method}
\title{Class "workbook"}
\description{
This is \pkg{XLConnect}'s main entity representing a Microsoft Excel workbook. S4 objects of this class
and corresponding methods are used to manipulate the underlying Excel workbook instances.
}
\section{Objects from the Class}{
Objects can be created by calls of the form \code{\link{loadWorkbook}(filename, create)}. This is
a shortcut form of \code{new("workbook", filename, create)} with some additional error checking.
}
\section{Slots}{
  \describe{
    \item{\code{filename}:}{Object of class \code{character} which represents the filename of the
    underlying Microsoft Excel workbook.}
    \item{\code{jobj}:}{Object of class \code{jobjRef} (see package \pkg{rJava}) which represents a Java object reference that is used
    in the back-end to manipulate the underlying Excel workbook instance.}
  }
  These slots should not be accessed directly in user code. \code{workbook} objects should only be manipulated
  via the corresponding methods.
}
\references{
Wikipedia: Office Open XML\cr
\url{http://en.wikipedia.org/wiki/Office_Open_XML}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
\pkg{XLConnect} supports both Excel '97 (*.xls) and OOXML (Excel 2007+, *.xlsx) file formats.\cr\cr

A \code{workbook}'s underlying Excel file is not saved (or being created in case the file did not exist and \code{create = TRUE} has been
specified) unless the \code{\link[=saveWorkbook-methods]{saveWorkbook}} method has been called on the object. This provides more flexibility to the user to decide
when changes are saved and also provides better performance in that several changes can be written in one go (normally at the end, rather than 
after every operation causing the file to be rewritten again completely each time). This is due to the fact that workbooks are manipulated in-memory and
are only written to disk with specifically calling \code{\link[=saveWorkbook-methods]{saveWorkbook}}.
}
\seealso{
\code{\link{loadWorkbook}}, \code{\link[=saveWorkbook-methods]{saveWorkbook}}
}
\examples{
# Create a new workbook 'myWorkbook.xlsx' 
# (assuming the file to not exist already)
wb <- loadWorkbook("myWorkbook.xlsx", create = TRUE)

# Create a worksheet called 'mtcars'
createSheet(wb, name = "mtcars")

# Write built-in dataset 'mtcars' to sheet 'mtcars' created above
writeWorksheet(wb, mtcars, sheet = "mtcars")

# Save workbook - this actually writes the file 'myWorkbook.xlsx' to disk
saveWorkbook(wb)
}
\keyword{classes}
