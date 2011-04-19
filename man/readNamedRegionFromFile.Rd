\name{readNamedRegionFromFile}
\alias{readNamedRegionFromFile}
\title{Reading named regions from an Excel file (wrapper function)}
\description{
Reads named regions from an Excel file.
}
\usage{
readNamedRegionFromFile(file, name, header = "TRUE")
}
\arguments{
  \item{file}{The file name of the workbook to read}
  \item{name}{The name of the named region to read from}
  \item{header}{The argument \code{header} specifies if the first row
    should be interpreted as column names. The default value is \code{TRUE}.}
}

\details{This is a convenience wrapper to read named regions from a file
  without creating an intermediate \code{\linkS4class{workbook}} object.
  See \link[=readNamedRegion-methods]{readNamedRegion} for more details.}

\author{
Thomas Themel\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}

\seealso{
  \code{\link[=readNamedRegion-methods]{readNamedRegion}},
  \code{\link{readWorksheetFromFile}},
  \code{\link{writeNamedRegionToFile}},\cr
  \code{\link{writeWorksheetToFile}},
  \code{\link[=onErrorCell-methods]{onErrorCell}}
}
\examples{
# multiregion xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/multiregion.xlsx", 
                             package = "XLConnect")
                             
# Load a single named region into a single data.frame.
df <- readNamedRegionFromFile(demoExcelFile, name="Iris")

# Load multiple regions at once - returns a (named) list 
# of data.frames.
df <- readNamedRegionFromFile(demoExcelFile, 
                              name=c("Calendar", "Iris", "IQ"))
}
\keyword{IO}
