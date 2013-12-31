\name{readTable}
\docType{methods}
\alias{readTable}
\alias{readTable-methods}
\alias{readTable,workbook,numeric-method}
\alias{readTable,workbook,character-method}
\title{Reading Excel tables from a workbook}
\description{
Reads Excel tables (Office 2007+) from a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{readTable}{workbook,numeric}(object, sheet, table, header, rownames, colTypes, forceConversion, 
dateTimeFormat, check.names, useCachedValues, keep, drop, simplify, readStrategy)
\S4method{readTable}{workbook,character}(object, sheet, table, header, rownames, colTypes, forceConversion, 
dateTimeFormat, check.names, useCachedValues, keep, drop, simplify, readStrategy)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{sheet}{The index or name of the worksheet on which to look for the specified \code{table}}
  \item{table}{The name of the table to read}
  \item{header}{The argument \code{header} specifies if the first row
    should be interpreted as column names. The default value is \code{TRUE}.}
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
  \item{check.names}{\code{logical} specifying if column names of the resulting \code{data.frame}
  should be checked to ensure that they are syntactically valid valid variable names and are not
  duplicated. See the \code{check.names} argument of \code{\link{data.frame}}. Defaults to \code{TRUE}.}
  \item{useCachedValues}{\code{logical} specifying whether to read cached formula results from the 
  workbook instead of re-evaluating them. This is particularly helpful in cases for reading data
  produced by Excel features not supported in XLConnect like references to external workbooks. 
  Defaults to \code{FALSE}, which means that formulas will be evaluated by XLConnect.}
  \item{keep}{List of column names or indices to be kept in the output data frame.
  It is possible to specify either \code{keep} or \code{drop}, but not both at the same time. 
  Defaults to \code{NULL}. If a vector is passed as argument, it will be wrapped into a list.
  This list gets replicated to match the length of the other arguments.}
  \item{drop}{List of column names or indices to be dropped in the output data frame.
  It is possible to specify either \code{keep} or \code{drop}, but not both at the same time.
  Defaults to \code{NULL}. If a vector is passed as argument, it will be wrapped into a list.
  This list gets replicated to match the length of the other arguments.}
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
\references{
Overview of Excel tables\cr
\url{http://office.microsoft.com/en-001/excel-help/overview-of-excel-tables-HA010048546.aspx}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
If no specific column types (see argument \code{colTypes}) are specified,
\code{readNamedRegion} tries to determine the resulting column types
based on the read cell types. If different cell types are found in a
specific column, the most general of those is used and mapped to the
corresponding R data type. The order of data types from least to most
general is Boolean (\code{logical}) < DateTime (\code{POSIXct}) <
Numeric (\code{numeric}) < String (\code{character}). E.g. if a column
is read that contains cells of type Boolean, Numeric and String then the
resulting column in R would be \code{character} since \code{character}
is the most general type.\cr\cr
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
  \code{\link[=readNamedRegion-methods]{readNamedRegion}},
  \code{\link[=readWorksheet-methods]{readWorksheet}},
  \code{\link[=writeNamedRegion-methods]{writeNamedRegion}},\cr
  \code{\link[=writeWorksheet-methods]{writeWorksheet}},
  \code{\link{readNamedRegionFromFile}},
  \code{\link[=onErrorCell-methods]{onErrorCell}}
}
\examples{
# mtcars xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/mtcars.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(demoExcelFile)

# Read table 'MtcarsTable' from sheet 'mtcars_table'
data <- readTable(wb, sheet = "mtcars_table", table = "MtcarsTable")
}
\keyword{methods}
\keyword{IO}
