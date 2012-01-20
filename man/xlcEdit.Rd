\name{xlcEdit}
\alias{xlcEdit}
\title{
Editing data sets in an Excel file editor
}
\description{
Provides the capability to edit an object/data.frame in an Excel file
editor. After editing, the object is restored in the R session with the
corresponding changes.
}
\usage{
xlcEdit(obj, pos = globalenv(), ext = ".xlsx")
}
\arguments{
  \item{obj}{
  Object (\code{data.frame}) to edit.
}
  \item{pos}{
  Where to look for the object specified by \code{obj}.
  See \code{pos} argument of \code{\link{get}} for more
  information.
}
  \item{ext}{
  Extension to use for the Excel file being created. Defaults
  to ".xlsx".
}
}
\details{
This function uses \code{\link{xlcDump}} and \code{\link{xlcRestore}} to
dump objects to and restore objects from Excel files. An OS command is invoked
to open the temporary Excel file in the default editor. Changes to the file have 
to be saved in order for them to take effect in the restored object.
}
\value{
Invisibly returns the value of the \code{\link{xlcRestore}} operation.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
This function only works under Windows and MacOS with a corresponding Excel
file editor, e.g. MS Excel or LibreOffice. Attempts to use this function under
another OS will result in an error being thrown.
}
\seealso{
\code{\link{xlcDump}}, \code{\link{xlcRestore}}, 
\code{\link[=writeNamedRegion-methods]{writeNamedRegion}}, 
\code{\link[=writeWorksheet-methods]{writeWorksheet}}, \code{\link{writeNamedRegionToFile}},
\code{\link{writeWorksheetToFile}}
}
\examples{
\dontrun{
myObj = mtcars
xlcEdit(myObj)
}
}
\keyword{utilities}
\keyword{misc}
