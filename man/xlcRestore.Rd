\name{xlcRestore}
\alias{xlcRestore}
\title{
Restoring objects from Excel files
}
\description{
Restores objects from Excel files that have been dumped using \code{\link{xlcDump}}.
}
\usage{
xlcRestore(file = "dump.xlsx", pos = -1, overwrite = FALSE)
}
\arguments{
  \item{file}{
  Excel file from which to restore objects. This is normally a file that has been produced with
  \code{\link{xlcDump}}. Defaults to \code{"dump.xlsx"}.
}
  \item{pos}{
  Environment into which to restore objects. Can be specified either as an \code{integer}
  specifying the position in the \code{search} list, as a \code{character} naming an element in the \code{search}
  list or as an \code{environment}. Defaults to \code{-1} which refers to the current environment.
}
  \item{overwrite}{
  \code{logical} specifying if data objects should be overwritten if they already exist inside the
  environment \code{pos}.
}
}
\value{
Named \code{logical} vector specifying if objects have been restored or not. An object may not be restored because there
was an issue with reading the data from the worksheet or the object already existed in the environment \code{pos}
(and \code{overwrite = FALSE}).
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\link{xlcDump}}, \code{\link[=readNamedRegion-methods]{readNamedRegion}}, 
\code{\link[=readWorksheet-methods]{readWorksheet}}, \code{\link{readNamedRegionFromFile}},\cr
\code{\link{readWorksheetFromFile}}, \code{\link{xlcEdit}}
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
