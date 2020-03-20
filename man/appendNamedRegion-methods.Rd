\name{appendNamedRegion-methods}
\docType{methods}
\alias{appendNamedRegion}
\alias{appendNamedRegion-methods}
\alias{appendNamedRegion,workbook-method}
\alias{appendNamedRegion,workbook,ANY-method}
\title{Appending data to a named region}
\description{
Appends data to an existing named region.
}
\usage{
\S4method{appendNamedRegion}{workbook,ANY}(object,data,name,header,rownames)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{data}{Data to write}
  \item{name}{Name of the (existing) named region to which to append the \code{data}}
  \item{header}{Specifies if the column names should be written. The default is \code{FALSE}.}
  \item{rownames}{Name (\code{character}) of column to use for the row names of the provided
  \code{data} object. If specified, the row names of the data object (\code{data.frame}) will be
  included as an additional column with the specified name. If \code{rownames = NULL} (default),
  no row names will be included.}
}

\details{
Appends \code{data} to the existing named region specified by \code{name}. The \code{data} is
appended at the bottom of the named region. See \code{\link[=writeNamedRegion-methods]{writeNamedRegion}}
for further information on writing named regions.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
Named regions are automatically redefined to the area occupied by the
previous and the newly appended data. This guarantees that the complete set 
of data can be re-read using \code{\link[=readNamedRegion-methods]{readNamedRegion}}.
Note however, that no checks are performed to see whether the appended data has the
same shape/structure as the previous data.
}
\seealso{
  \code{\linkS4class{workbook}},
  \code{\link[=writeNamedRegion-methods]{writeNamedRegion}},
  \code{\link[=readNamedRegion-methods]{readNamedRegion}},
  \code{\link[=writeWorksheet-methods]{writeWorksheet}},
  \code{\link[=appendWorksheet-methods]{appendWorksheet}},
  \code{\link[=readWorksheet-methods]{readWorksheet}}
}
\examples{
\dontrun{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Append mtcars data set to named region named 'mtcars'
appendNamedRegion(wb, mtcars, name = "mtcars")
}
}
\keyword{methods}
\keyword{IO}
