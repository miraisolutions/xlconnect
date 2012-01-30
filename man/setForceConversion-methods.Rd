\name{setForceConversion-methods}
\docType{methods}
\alias{setForceConversion}
\alias{setForceConversion-methods}
\alias{setForceConversion,workbook-method}
\title{Forcing data type conversion when reading in data}
\description{
 This function sets a flag to force data type conversion when reading in data.
}
\usage{
\S4method{setForceConversion}{workbook}(object,value)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{value}{\code{logical} specifying if data type conversion should be forced or not.
  Defaults to \code{FALSE}.}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
Data type conversions are applied only where possible. In case where a conversion is not
possible (e.g. from Boolean to DateTime) a missing value (\code{NA}) is returned. 
}
\seealso{
\code{\linkS4class{workbook}}
}
\examples{
# conversion xlsx file from demoFiles subfolder of package XLConnect
excelFile <- system.file("demoFiles/conversion.xlsx", package = "XLConnect")

# Load workbook
wb <- loadWorkbook(excelFile)

# Force conversion when reading in data
setForceConversion(wb, TRUE)

# Read named region 'conversion' with pre-specified column types
df <- readNamedRegion(wb, name = "conversion", header = TRUE,
                      colTypes = c(XLC$DATA_TYPE.NUMERIC,
                                   XLC$DATA_TYPE.DATETIME,
                                   XLC$DATA_TYPE.BOOLEAN))
}
\keyword{methods}
\keyword{utilities}
