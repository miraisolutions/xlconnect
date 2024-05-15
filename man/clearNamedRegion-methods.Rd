\name{clearNamedRegion-methods}
\docType{methods}
\alias{clearNamedRegion}
\alias{clearNamedRegion-methods}
\alias{clearNamedRegion,workbook,character-method}
\title{Clearing named regions in a workbook}
\description{
Clears named regions in a \code{\linkS4class{workbook}}.
}
\usage{
  \S4method{clearNamedRegion}{workbook,character}(object, name, worksheetScope)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The name of the named region to clear}
  \item{worksheetScope}{Optional - the name of the worksheet in which the region is scoped;
  useful if different sheets have scoped regions with the same name.}
}

\details{
Clearing a named region/range means to clear all the cells associated with that named
region. Clearing named regions can be useful if (named) data sets in a worksheet need 
to be replaced, i.e. data is first read, modified in R and finally written back to the 
the same named region. Without clearing the named region first, (parts of) the original 
data may still be visible if they occupied a larger range in the worksheet.

If \code{worksheetScope} is unspecified, the first matching name found anywhere in the workbook
will be cleared. Otherwise, only a name specifically scoped to the worksheet may be cleared.
To only clear a name in global scope, pass \code{""} as the value.
}

\author{
Nicola Lambiase\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=clearSheet-methods]{clearSheet}}, 
\code{\link[=clearRange-methods]{clearRange}}, 
\code{\link[=clearRangeFromReference-methods]{clearRangeFromReference}},
\code{\link[=clearSheet-methods]{clearSheet}}
}
\examples{
  \dontrun{
# mtcars xlsx file from demoFiles subfolder of 
# package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", 
                             package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Read named region 'mtcars'
data <- readNamedRegion(wb, name = "mtcars", header = TRUE)

# Only consider cars with a weight >= 5
data <- data[data$wt >= 5, ]

# Clear original named region
clearNamedRegion(wb, name = "mtcars")

# Write subsetted data back
# Note: this is covering a smaller area now -
# writeNamedRegion automatically redefines the named region
# to the size/area of the data
writeNamedRegion(wb, data = data, name = "mtcars",
                 header = TRUE) 
} 
}
\keyword{methods}
\keyword{utilities}
