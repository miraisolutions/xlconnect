\name{writeWorksheetToFile}
\alias{writeWorksheetToFile}
\title{Writing named regions to a file (wrapper function)}
\description{
  \code{writeWorksheetToFile} writes named regions to a file
  specified by path.
}
\usage{
writeWorksheetToFile(file, styleAction = "PREDEFINED", ...)
}
\arguments{
\item{file}{The path name of the file to write to.}
\item{...}{Arguments passed on to \link[=writeWorksheet-methods]{writeWorksheet}}
\item{styleAction}{A string indicating how to style the output - not vectorized!}
}
\author{
Thomas Themel\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}

\seealso{
\code{\link[=writeWorksheet-methods]{writeWorksheet}}, \code{\link{readWorksheetFromFile}}, 
\code{\link{readNamedRegionFromFile}}, \code{\link{writeNamedRegionToFile}}
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
writeWorksheetToFile(template, data=list(i1=iris,i2=iris,i3=iris),
sheet=c("FirstSheet","SecondSheet", "FirstSheet"),
startRow=c(1,101,201), startCol = c(1,11,21))

}
\keyword{IO}