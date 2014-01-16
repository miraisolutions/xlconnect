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
\S4method{readWorksheet}{workbook,numeric}(object,sheet,startRow,startCol,endRow,endCol,autofitRow,autofitCol,
region,header,rownames,colTypes,forceConversion,dateTimeFormat,check.names,
useCachedValues,keep,drop, simplify, readStrategy)
\S4method{readWorksheet}{workbook,character}(object,sheet,startRow,startCol,endRow,endCol,autofitRow,autofitCol,
region,header,rownames,colTypes,forceConversion,dateTimeFormat,check.names,
useCachedValues,keep,drop, simplify, readStrategy)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The name or index of the worksheet to read from}
  \item{startRow}{The index of the first row to read from. Defaults to \code{0} meaning that the start row is determined automatically.}
  \item{startCol}{The index of the first column to read from. Defaults to \code{0} meaning that the start column is determined automatically.}
  \item{endRow}{The index of the last row to read from. Defaults to \code{0} meaning that the end row is determined automatically.}
  \item{endCol}{The index of the last column to read from. Defaults to \code{0} meaning that the end column is determined automatically.}
  \item{autofitRow}{\code{logical} specifying if leading and trailing empty rows should be
  skipped. Defaults to \code{TRUE}.}
  \item{autofitCol}{\code{logical} specifying if leading and trailing empty columns should be skipped. Defaults to \code{TRUE}.}
  \item{region}{A range specifier in the form 'A10:B18'. This provides an alternative way to specify
  \code{startRow}, \code{startCol}, \code{endRow} and \code{endCol}. Range specifications take precedence
  over index specifications.}
  \item{header}{Interpret the first row of the specified area as column
    headers. The default is \code{TRUE}.}
  \item{rownames}{Index (\code{numeric}) or name (\code{character}) of column that
  should be used as row names. The corresponding column will be removed from the data
  set. Defaults to \code{NULL} which means that no row names are applied. Row names must be either
  \code{integer} or \code{character}. Non-numeric columns will be coerced to \code{character}.}
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
  \item{check.names}{\code{logical} specifying if column names of the resulting \code{data.frame}
  should be checked to ensure that they are syntactically valid variable names and are not
  duplicated. See the \code{check.names} argument of \code{\link{data.frame}}. Defaults to \code{TRUE}.}
  \item{useCachedValues}{\code{logical} specifying whether to read cached formula results from the 
  workbook instead of re-evaluating them. This is particularly helpful in cases for reading data
  produced by Excel features not supported in XLConnect like references to external workbooks. 
  Defaults to \code{FALSE}, which means that formulas will be evaluated by XLConnect.}
  \item{keep}{Vector of column names or indices to be kept in the output data frame.
  It is possible to specify either \code{keep} or \code{drop}, but not both at the same time.
  Defaults to \code{NULL}. If a vector is passed as argument, it will be wrapped into a list.
  This list gets replicated to match the length of the other arguments.
  Example: if \code{sheet = c("Sheet1", "Sheet2", "Sheet3")} and \code{keep = c(1,2)}, \code{keep} will be
  internally converted into \code{list(c(1,2))} and then replicated to match the number of sheets, i.e.
  \code{keep = list(c(1,2), c(1,2), c(1,2))}. The result is that the first two columns of each sheet
  are kept. If \code{keep = list(1,2)} is specified, it will be replicated as \code{list(1,2,1)}, i.e. respectively
  the first, second and first column of the sheets "Sheet1", "Sheet2", "Sheet3"
  will be kept.}
  \item{drop}{Vector of column names or indices to be dropped in the output data frame.
  It is possible to specify either \code{keep} or \code{drop}, but not both at the same time.
  Defaults to \code{NULL}. If a vector is passed as argument, it will be wrapped into a list.
  This list gets replicated to match the length of the other arguments.
  Example: if \code{sheet = c("Sheet1", "Sheet2", "Sheet3")} and \code{drop = c(1,2)}, \code{drop} will be
  internally converted into \code{list(c(1,2))} and then replicated to match the number of sheets, i.e.
  \code{drop = list(c(1,2), c(1,2), c(1,2))}. The result is that the first two columns of each sheet
  are dropped. If \code{drop = list(1,2)} is specified, it will be replicated as \code{list(1,2,1)}, i.e. respectively
  the first, second and first column of the sheets "Sheet1", "Sheet2", "Sheet3"
  will be dropped.}
  \item{simplify}{\code{logical} specifying if the result should be simplified, e.g. in case the \code{data.frame}
  would only have one row or one column (and data types match). Simplifying here is identical to calling \code{unlist}
  on the otherwise resulting \code{data.frame} (using \code{use.names = FALSE}). The default is \code{FALSE}.}
  \item{readStrategy}{\code{character} specifying the reading strategy to use. Currently supported strategies are:
  \itemize{
  \item \code{"default"} (default): Can handle all supported data types incl. date/time values and can deal 
  directly with missing value identifiers (see \code{\link[=setMissingValue-methods]{setMissingValue}})
  \item \code{"fast"}: Increased read performance. Date/time values are read as numeric (number of days since 1900-01-01; 
  fractional days represent hours, minutes, and seconds) and only blank cells are recognized as missing (missing value 
  identifiers as set in \code{\link[=setMissingValue-methods]{setMissingValue}} are ignored)
  }}
}

\details{
Reads data from the worksheet specified by \code{sheet}. Data is read
starting at the top left corner specified by \code{startRow} and
\code{startCol} down to the bottom right corner specified by
\code{endRow} and \code{endCol}. If \code{header = TRUE}, the first row
is interpreted as column names of the resulting \code{data.frame}.\cr If
\code{startRow <= 0} then the first available row in the sheet is assumed. 
If \code{endRow <= 0} then the last available row in the sheet is assumed.
If \code{startCol <= 0} then the minimum column between \code{startRow} and
\code{endRow} is assumed. If \code{endCol <= 0} then the maximum column between
\code{startRow} and \code{endRow} is assumed. In other words, if no boundaries are specified \code{readWorksheet} assumes the "bounding box" of the data as the corresponding boundaries.\cr
The arguments \code{autofitRow} and \code{autofitCol} (both defaulting to \code{TRUE}) can be used to skip leading and trailing empty rows even in case \code{startRow}, \code{endRow}, \code{startCol} and \code{endCol} are specified to values \code{> 0}. This can be useful if data is expected within certain given boundaries but the exact location is not available.\cr\cr

If all four coordinate arguments are missing this behaves as above with
\code{startRow = 0}, \code{startCol = 0}, \code{endRow = 0} and
\code{endCol = 0}. In this case  \code{readWorksheet} assumes the
"bounding box" of the data as the corresponding boundaries.\cr\cr

All arguments (except \code{object}) are vectorized. As such, multiple worksheets (and also multiple data regions
from the same worksheet) can be read with one method call. If only one single data region is read, the return value 
is a \code{data.frame}. If multiple data regions are specified, the return value is a \code{list} of \code{data.frame}'s 
returned in the order they have been specified. If worksheets have been specified by name, the \code{list} will be a
named \code{list} named by the corresponding worksheets.
}
\author{
Martin Studer\cr
Thomas Themel\cr
Nicola Lambiase\cr
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
the number of days since 1900-01-00 (yes, day 00! - see \url{http://www.cpearson.com/excel/datetime.htm}).
Note that in R 0 is represented as 1899-12-31 since there is no 1900-01-00.
Fractional days represent hours, minutes, and seconds.
}
}
\seealso{
  \code{\linkS4class{workbook}},
  \code{\link[=writeWorksheet-methods]{writeWorksheet}},
  \code{\link[=readNamedRegion-methods]{readNamedRegion}},
  \code{\link[=writeNamedRegion-methods]{writeNamedRegion}},\cr
  \code{\link{readWorksheetFromFile}},
  \code{\link[=readTable-methods]{readTable}},
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
                    
## Example 5:
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Read the columns 1, 3 and 5 from the sheet 'mtcars' (with default header = TRUE)
data <- readWorksheet(wb, sheet = "mtcars", keep=c(1,3,5))
}
\keyword{methods}
\keyword{IO}
