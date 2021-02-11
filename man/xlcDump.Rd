\name{xlcDump}
\alias{xlcDump}
\title{
Dumping data sets to Excel files
}
\description{
Dumps data sets to Excel files by writing each object to a separate worksheet.
}
\usage{
xlcDump(list, ..., file = "dump.xlsx", pos = -1, overwrite = FALSE)
}
\arguments{
  \item{list}{
  \code{character} vector of names of objects inside environment \code{pos} to dump into an Excel file. Objects
  will be written using \code{\link[=writeWorksheet-methods]{writeWorksheet}} - as such any object will be coerced
  to a \code{data.frame}. If \code{missing}, the list of objects will be determined via the function \code{ls} which
  takes any arguments specified via \dots.
}
  \item{\dots}{
  Arguments that will be passed to the \code{ls} function for getting a list of object names in case the \code{list}
  argument is \code{missing}.
}
  \item{file}{
  Excel file to which objects will be dumped. Can be an existing or a new file. Defaults to \code{"dump.xlsx"}.
}
  \item{pos}{
  Environment in which to look for objects. Can be specified either as an \code{integer}
  specifying the position in the \code{search} list, as a \code{character} naming an element in the \code{search}
  list or as an \code{environment}. Defaults to \code{-1} which refers to the current environment.
}
  \item{overwrite}{
  \code{logical} specifying if data should be overwritten if objects with the same name have already been dumped
  to the Excel \code{file}.
}
}
\details{
Each object is written to a separate worksheet named by the name of the object. Objects are written using the
\code{\link[=writeWorksheet-methods]{writeWorksheet}} method - as such any object will be coerced to \code{data.frame}. 
}
\value{
Named \code{logical} vector specifying if objects have been dumped or not. An object may not be dumped because there
was an issue with the coercion to a \code{data.frame} or the object already existed (and \code{overwrite = FALSE}) in
the workbook.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\link{xlcRestore}}, \code{\link[=writeNamedRegion-methods]{writeNamedRegion}}, 
\code{\link[=writeWorksheet-methods]{writeWorksheet}}, \code{\link{writeNamedRegionToFile}},
\code{\link{writeWorksheetToFile}}, \code{\link{xlcEdit}}
}
\examples{\dontrun{
require(datasets)
xlcDump(c("airquality", "CO2", "iris", "PlantGrowth", "swiss"),
        file = "myDump.xlsx", pos = "package:datasets")
xlcRestore(file = "myDump.xlsx", overwrite = TRUE)
# clean up
file.remove("myDump.xlsx")
}
}
\keyword{utilities}
\keyword{file}
