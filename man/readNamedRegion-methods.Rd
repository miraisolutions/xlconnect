\name{readNamedRegion-methods}
\docType{methods}
\alias{readNamedRegion}
\alias{readNamedRegion-methods}
\alias{readNamedRegion,workbook,character,logical-method}
\alias{readNamedRegion,workbook,character,missing-method}
\title{Reading named regions from a workbook}
\description{
Reads named regions from a \code{\linkS4class{workbook}} object.
}
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook", name = "character", header = "logical")}}{
Reads the named regions \code{name} from the specified \code{\linkS4class{workbook}} object. 
The argument \code{header} specifies if the first row should be interpreted as column names.
}

\item{\code{signature(object = "workbook", name = "character", header = "missing")}}{
Same as above with \code{header = TRUE}.
}
}}
\references{
What are named regions/ranges? \url{http://www.officearticles.com/excel/named_ranges_in_microsoft_excel.htm}\cr
How to create named regions/ranges? \url{http://www.youtube.com/watch?v=iAE9a0uRtpM}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
The arguments \code{name} and \code{header} are vectorized. As such, multiple named regions
can be read with one method call. If only one single named region is read, the return value is a \code{data.frame}.
If multiple named regions are specified, the return value is a (named) \code{list} of \code{data.frame}'s returned in the
order they have been specified with the argument \code{name}.\cr\cr
\code{readNamedRegion} tries to determine the resulting column types based on the read cell types. If different cell
types are found in a specific column, the most general of those is used and mapped to the corresponding R data type.
The order of data types from least to most general is Boolean (\code{logical}) < DateTime (\code{POSIXct}) < 
Numeric (\code{numeric}) < String (\code{character}). E.g. if a column is read that contains cells of type Boolean,
Numeric and String then the resulting column in R would be \code{character} since \code{character} is the most general
type.
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=readWorksheet-methods]{readWorksheet}}, 
\code{\link[=writeNamedRegion-methods]{writeNamedRegion}}, \code{\link[=writeWorksheet-methods]{writeWorksheet}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")
# Load workbook
wb <- loadWorkbook(demoExcelFile)
# Read named region 'mtcars' (with default header = TRUE)
data <- readNamedRegion(wb, name = "mtcars")
}
\keyword{methods}
\keyword{IO}
