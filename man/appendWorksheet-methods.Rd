\name{appendWorksheet-methods}
\docType{methods}
\alias{appendWorksheet}
\alias{appendWorksheet-methods}
\alias{appendWorksheet,workbook,ANY,character-method}
\alias{appendWorksheet,workbook,ANY,numeric-method}
\title{Appending data to worksheets}
\description{
Appends data to worksheets of a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{appendWorksheet}{workbook,ANY,character}(object,data,sheet,header,rownames)
\S4method{appendWorksheet}{workbook,ANY,numeric}(object,data,sheet,header,rownames)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to write to}
  \item{data}{Data to append}
  \item{sheet}{The name or index of the sheet to append the \code{data} to}
  \item{header}{Specifies if the column names should be written. The default is \code{TRUE}.}
  \item{rownames}{Name (\code{character}) of column to use for the row names of the provided
  \code{data} object. If specified, the row names of the data object (\code{data.frame}) will be
  included as an additional column with the specified name. If \code{rownames = NULL} (default),
  no row names will be included.}
}

\details{
Appends \code{data} to the worksheet specified by \code{sheet}. Data will be appended at the
bottom and left most column containing some data. If more complex "appending schemes" are required
you may make direct use of \code{\link[=writeWorksheet-methods]{writeWorksheet}}.
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
  \code{\linkS4class{workbook}},
  \code{\link[=writeWorksheet-methods]{writeWorksheet}},
  \code{\link[=readWorksheet-methods]{readWorksheet}},
  \code{\link[=writeNamedRegion-methods]{writeNamedRegion}},
  \code{\link[=appendNamedRegion-methods]{appendNamedRegion}},
  \code{\link[=readNamedRegion-methods]{readNamedRegion}}
}
\examples{
\dontrun{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Append mtcars data set to worksheet named 'mtcars'
appendWorksheet(wb, mtcars, sheet = "mtcars")
}
}
\keyword{methods}
\keyword{IO}
