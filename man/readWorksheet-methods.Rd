\name{readWorksheet-methods}
\docType{methods}
\alias{readWorksheet}
\alias{readWorksheet-methods}
\alias{readWorksheet,workbook,character-method}
\alias{readWorksheet,workbook,numeric-method}
\title{Reading data from worksheets}
\description{
Reads data from worksheets of a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{readWorksheet}{workbook,numeric}(object,sheet,startRow,startCol,endRow,endCol,region,header,rownames,colTypes,forceConversion,dateTimeFormat)
\S4method{readWorksheet}{workbook,character}(object,sheet,startRow,startCol,endRow,endCol,region,header,rownames,colTypes,forceConversion,dateTimeFormat)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the worksheet to read from}
  \item{startRow}{The index of the first row to read from}
  \item{startCol}{The index of the first column to read from}
  \item{endRow}{The index of the last row to read from}
  \item{endCol}{The index of the last column to read from}
  \item{region}{A range specifier in the form 'A10:B18'. This provides an alternative way to specify
  \code{startRow}, \code{startCol}, \code{endRow} and \code{endCol}. Range specifications take precedence
  over index specifications.}
  \item{header}{Interpret the first row of the specified area as column
    headers. The default is \code{TRUE}.}
  \item{rownames}{Index (\code{numeric}) or name (\code{character}) of column that
  should be used as row names. The corresponding column will be removed from the data
  set. Defaults to \code{NULL} which means that no row names are applied.}
  \item{colTypes}{Column types to use when reading in the data. Specified as a \code{character}
  vector of the corresponding type names (see \code{\link{XLC}}; \code{XLC$DATA_TYPE.<?>}). You
  may also use R class names such as \code{numeric}, \code{character}, \code{logical} and \code{POSIXt}.
  The types are applied in the given order to the columns - elements are recycled if necessary.
  Defaults to \code{character(0)} meaning that column types are determined automatically (see the
  Note section for more information).\cr
  By default, type conversions are only applied if the specified column type is a more generic type
  (e.g. from Numeric to String) - otherwise \code{NA} is returned. The \code{forceConversion} flag can be
  set to force conversion into less generic types where possible.}
  \item{forceConversion}{\code{logical} specifying if conversions to less generic types should be
  forced. Defaults to \code{FALSE} meaning that if a column is specified
  to be of a certain type via the \code{colTypes} argument and a more generic type is detected in
  the column, then \code{NA} will be returned (example: column is specified to be DateTime but a
  more generic String is found). Specifying \code{forceConversion = TRUE} will try to enforce a
  conversion - if it succeeds the corresponding (converted) value will be returned, otherwise
  \code{NA}. See the Note section for some additional information.}
  \item{dateTimeFormat}{Date/time format used when doing date/time conversions. Defaults to\cr
  \code{getOption("XLConnect.dateTimeFormat")}. This should be a POSIX format specifier according
  to \code{\link{strptime}} although not all specifications have been implemented yet - the most
  important ones however are available.}
}

\details{
Reads data from the worksheet specified by \code{sheet}. Data is read
starting at the top left corner specified by \code{startRow} and
\code{startCol} down to the bottom right corner specified by
\code{endRow} and \code{endCol}. If \code{header = TRUE}, the first row
is interpreted as column names of the resulting \code{data.frame}.\cr If
\code{startRow <= 0} then the first available (logical) row is
assumed. If \code{startCol <= 0} then the column  of the first (logical)
cell of the start row (\code{startRow}) is assumed. If \code{endRow <= 0}
then the last available  (logical) row is assumed. If \code{endCol <= 0}
then the maximum column between \code{startRow} and \code{endRow} is assumed. 
In other words, if no boundaries are specified \code{readWorksheet} assumes 
the "bounding box" of the data as the corresponding boundaries.\cr\cr

If all four coordinate arguments are missing this behaves as above with
\code{startRow = 0}, \code{startCol = 0}, \code{endRow = 0} and
\code{endCol = 0}. In this case  \code{readWorksheet} assumes the
"bounding box" of the data as the corresponding boundaries.\cr\cr

All arguments (except \code{object}) are vectorized. As such, multiple worksheets (and also multiple data regions
from the same worksheet) can be read with one method call. If only one single data region is read, the return value 
is a \code{data.frame}. If multiple data regions are specified, the return value is a \code{list} of \code{data.frame}'s 
returned in the order they have been specified.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
If no specific column types (see argument \code{colTypes}) are specified,
\code{readWorksheet} tries to determine the resulting column types based on the read cell types. If different cell
types are found in a specific column, the most general of those is used and mapped to the corresponding R data type.
The order of data types from least to most general is Boolean (\code{logical}) < DateTime (\code{POSIXct}) < 
Numeric (\code{numeric}) < String (\code{character}). E.g. if a column is read that contains cells of type Boolean,
Numeric and String then the resulting column in R would be \code{character} since \code{character} is the most general
type.\cr\cr
Some additional information with respect to forcing data type conversion
using \code{forceConversion = TRUE}:
\itemize{
\item Forcing conversion from String to Boolean: \code{TRUE} is returned
if and only if the target string is "true" (ignoring any capitalization).
Any other string will return \code{FALSE}.
\item Forcing conversion from Numeric to DateTime: since Excel understands
Dates/Times as Numerics with some additional formatting, a conversion from
a Numeric to a DateTime is actually possible. Numerics in this case represent
the number of days since 1900-01-01. Fractional days represent hours, minutes, 
and seconds.
}
}
\seealso{
  \code{\linkS4class{workbook}},
  \code{\link[=writeWorksheet-methods]{writeWorksheet}},
  \code{\link[=readNamedRegion-methods]{readNamedRegion}},
  \code{\link[=writeNamedRegion-methods]{writeNamedRegion}},\cr
  \code{\link{readWorksheetFromFile}},
  \code{\link[=onErrorCell-methods]{onErrorCell}}
}
\examples{
## Example 1:
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Read worksheet 'mtcars' (providing no specific area bounds;
# with default header = TRUE)
data <- readWorksheet(wb, sheet = "mtcars")


## Example 2:
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Read worksheet 'mtcars' (providing area bounds; with default header = TRUE)
data <- readWorksheet(wb, sheet = "mtcars", startRow = 1, startCol = 3,
                      endRow = 15, endCol = 8)


## Example 3:
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Read worksheet 'mtcars' (providing area bounds using the region argument;
# with default header = TRUE)
data <- readWorksheet(wb, sheet = "mtcars", region = "C1:H15")


## Example 4:
# conversion xlsx file from demoFiles subfolder of package XLConnect
excelFile <- system.file("demoFiles/conversion.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(excelFile)

# Read worksheet 'Conversion' with pre-specified column types
# Note: in the worksheet all data was entered as strings!
# forceConversion = TRUE is used to force conversion from String
# into the less generic data types Numeric, DateTime & Boolean
df <- readWorksheet(wb, sheet = "Conversion", header = TRUE,
                    colTypes = c(XLC$DATA_TYPE.NUMERIC,
                                 XLC$DATA_TYPE.DATETIME,
                                 XLC$DATA_TYPE.BOOLEAN),
                    forceConversion = TRUE,
                    dateTimeFormat = "\%Y-\%m-\%d \%H:\%M:\%S")
}
\keyword{methods}
\keyword{IO}
