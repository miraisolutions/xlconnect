\name{loadWorkbook}
\alias{loadWorkbook}
\title{
Loading Microsoft Excel workbooks
}
\description{
Loads or creates a Microsoft Excel \code{\linkS4class{workbook}} for further manipulation.
}
\usage{
loadWorkbook(filename, create = FALSE)
}
\arguments{
  \item{filename}{
  Filename (absolute or relative) of Excel workbook to be loaded. Supported are Excel '97 (*.xls) and
  OOXML (Excel 2007+, *.xlsx) file formats.
}
  \item{create}{
  Specifies if the file should be created if it does not already exist (default is \code{FALSE}). Note that
  \code{create = TRUE} has no effect if the specified file exists, i.e. an existing file is loaded and not
  being recreated if \code{create = TRUE}.
}
}
\value{
Returns a \code{\linkS4class{workbook}} object for further manipulation.
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
\code{loadWorkbook} is basically just a shortcut form of \code{new("workbook", filename, create)} with some additional error checking. As such
it is the preferred way of creating \code{\linkS4class{workbook}} instances.
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=saveWorkbook-methods]{saveWorkbook}}
}
\examples{
# Load existing demo Excel file 'mtcars.xlsx' from the XLConnect package
wb.mtcars <- loadWorkbook(system.file("demoFiles/mtcars.xlsx", 
                          package = "XLConnect"))

# Create new workbook
wb.new <- loadWorkbook("myNewExcelFile.xlsx", create = TRUE)

# NOTE: The above statement does not write the file to disk! 
# saveWorkbook(wb.new) would need to be called in order to write/save 
# the file to disk!
}
\keyword{file}
