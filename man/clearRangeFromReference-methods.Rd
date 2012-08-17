\name{clearRangeFromReference-methods}
\docType{methods}
\alias{clearRangeFromReference}
\alias{clearRangeFromReference-methods}
\alias{clearRangeFromReference,workbook,character-method}
\title{Clearing cell ranges in a workbook}
\description{
Clears cell ranges specified by area reference in a \code{\linkS4class{workbook}}.
}
\usage{
  \S4method{clearRangeFromReference}{workbook,character}(object, reference)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{reference}{\code{character} specifying an area reference in
  the form 'SheetX!A7:B19'}
}

\details{
Clearing a cell range means to remove any cells associated with that range. This
method is very similar to \code{\link[=clearRange-methods]{clearRange}}.
}

\author{
Nicola Lambiase\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=clearSheet-methods]{clearSheet}}, 
\code{\link[=clearNamedRegion-methods]{clearNamedRegion}}, 
\code{\link[=clearRange-methods]{clearRange}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of 
# package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", 
                             package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Clear ranges A5:D6 and G5:I7 on sheet mtcars
clearRangeFromReference(wb, reference = c("mtcars!A5:D6",
                        "mtcars!G5:I7"))
}
\keyword{methods}
\keyword{utilities}
