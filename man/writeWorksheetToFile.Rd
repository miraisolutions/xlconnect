\name{writeWorksheetToFile}
\alias{writeWorksheetToFile}
\title{Writing data to worksheets in an Excel file (wrapper function)}
\description{
Writes data to worksheets in an Excel file.
}
\usage{
writeWorksheetToFile(file, data, sheet, ..., styleAction = XLC$STYLE_ACTION.XLCONNECT)
}
\arguments{
\item{file}{The path name of the file to write to.}
\item{data}{Data to write}
\item{sheet}{The name or index of the sheet to write to}
\item{...}{Additional arguments passed to \link[=writeWorksheet-methods]{writeWorksheet}}
\item{styleAction}{Style action to be used when writing the data - not vectorized! 
The default is \code{XLC$STYLE_ACTION.XLCONNECT}. 
See \code{\link[=setStyleAction-methods]{setStyleAction}} for more information.}
}

\author{
Thomas Themel\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}

\seealso{
  \code{\link[=writeWorksheet-methods]{writeWorksheet}},
  \code{\link{writeNamedRegionToFile}},
  \code{\link{readWorksheetFromFile}},\cr
  \code{\link{readNamedRegionFromFile}} 
}
\examples{
# multiregion xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/multiregion.xlsx",
                             package = "XLConnect")

# create a copy of the template
template <- "template-ws.xlsx"
file.copy(demoExcelFile, template) 

# Write single data.frame to a specified location in an existing file
writeWorksheetToFile(template, data=iris, sheet="FirstSheet",
                     startRow=9, startCol = 9)

# create a copy of the template
template <- "template-multi-ws.xlsx"

# Write a few copies of the data.frame all over a new file
writeWorksheetToFile(template, data = list(i1 = iris, i2 = iris, i3 = iris),
                     sheet = c("FirstSheet", "SecondSheet", "FirstSheet"),
                     startRow = c(1,101,201), startCol = c(1,11,21))

}
\keyword{IO}