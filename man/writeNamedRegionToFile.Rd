\name{writeNamedRegionToFile}
\alias{writeNamedRegionToFile}
\title{Writing named regions to an Excel file (wrapper function)}
\description{
Writes named regions to an Excel file.
}
\usage{
writeNamedRegionToFile(file, data, name, formula=NA, ...,
styleAction = XLC$STYLE_ACTION.XLCONNECT)
}
\arguments{
\item{file}{The path name of the file to write to}
\item{data}{Data to write}
\item{name}{Name of the named region to write to}
\item{formula}{If formula is specified, each item defines the formula of
the named region identified by the corresponding entry of
\code{name}. Use this if you want to create the document from scratch
instead of writing to a template!}
\item{...}{Additional arguments passed to \code{\link[=writeNamedRegion-methods]{writeNamedRegion}}}
\item{styleAction}{Style action to be used when writing the data.\cr
The default is \code{XLC$STYLE_ACTION.XLCONNECT}. 
See \code{\link[=setStyleAction-methods]{setStyleAction}} for more information.}
}

\author{
Thomas Themel\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}

\seealso{
  \code{\link[=writeNamedRegion-methods]{writeNamedRegion}},
  \code{\link{writeWorksheetToFile}},
  \code{\link{readNamedRegionFromFile}},\cr
  \code{\link{readWorksheetFromFile}}
}
\examples{
# multiregion xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/multiregion.xlsx", 
                             package = "XLConnect")

template <- "template-ws.xlsx"
file.copy(demoExcelFile, template) 

# Write single data.frame to a named region in an existing file
writeNamedRegionToFile(template, name = "Iris", iris)

# Write to a new file, defining the sheet and named region as we write.
# Format according to XLConnect defaults
writeNamedRegionToFile("iris.xlsx", name = "Iris", data = iris,
                       formula = "IrisData!$C$4", 
                       styleAction = "XLCONNECT")
}
\keyword{IO}
