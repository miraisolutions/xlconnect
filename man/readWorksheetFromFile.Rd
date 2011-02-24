\name{readWorksheetFromFile}
\alias{readWorksheetFromFile}
\title{Reading data from worksheets in an Excel file (wrapper function)}
\description{
Reads data from worksheets in an Excel file.
}
\usage{
readWorksheetFromFile(file, ...)
}
\arguments{
\item{file}{The path name of the file to read from.}
\item{...}{Arguments passed to \link[=readWorksheet-methods]{readWorksheet}}
}

\details{
See \link[=readWorksheet-methods]{readWorksheet} for more information.
}

\author{
Thomas Themel\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}

\seealso{
  \code{\link[=readWorksheet-methods]{readWorksheet}},
  \code{\link{readNamedRegionFromFile}},
  \code{\link{writeWorksheetToFile}},\cr
  \code{\link{writeNamedRegionToFile}}
}
\examples{
# multiregion xlsx file from demoFiles subfolder of 
# package XLConnect
demoExcelFile <- system.file("demoFiles/multiregion.xlsx", 
                             package = "XLConnect")

# Read single area from first sheet of existing file,
# "B2:C3" in Excel speak
df.one <- readWorksheetFromFile(demoExcelFile, sheet = 1, 
                                header = FALSE, startCol = 2, 
                                startRow = 2, endCol = 3, 
                                endRow = 3)

# Read three data sets in one from known positions
dflist <- readWorksheetFromFile(demoExcelFile,
                                sheet = c("FirstSheet", 
                                          "FirstSheet", 
                                          "SecondSheet"),
                                header = TRUE, 
                                startRow = c(2,2,3), 
                                startCol = c(2,5,2),
                                endCol = c(5,8,6), 
                                endRow = c(9,15,153))
}
\keyword{IO}
