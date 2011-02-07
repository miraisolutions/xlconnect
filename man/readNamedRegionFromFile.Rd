\name{readNamedRegionFromFile}
\alias{readNamedRegionFromFile}
\title{Reading named regions from a file (wrapper function)}
\description{
  \code{readNamedRegionFromFile} reads named regions from a file
  specified by path.  
}
\usage{
readNamedRegionFromFile(file, name, header = "TRUE")
}
\arguments{}
\references{
What are named regions/ranges? \url{http://www.officearticles.com/excel/named_ranges_in_microsoft_excel.htm}\cr
How to create named regions/ranges? \url{http://www.youtube.com/watch?v=iAE9a0uRtpM}
}
\author{
Thomas Themel\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
}

\seealso{
\code{\link[=readNamedRegion-methods]{readNamedRegion}}, \code{\link{readWorksheetFromFile}}, 
\code{\link{writeNamedRegionToFile}}, \code{\link{writeWorksheetToFile}}
}
\examples{
# multiregion xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/multiregion.xlsx", package = "XLConnect")
# Load a single named region into a single data.frame.
df <- readNamedRegionFromFile(demoExcelFile, name="Iris")
# Load multiple regions at once - returns a list of named data.frames.
df <- readNamedRegionFromFile(demoExcelFile, name=c("Calendar", "Iris", "IQ"))
}
\keyword{IO}
