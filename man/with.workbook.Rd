\name{with.workbook}
\alias{with.workbook}
\title{Evaluate an R expression in a workbook environment}
\description{
Evaluate an R expression in an environment constructed from the named regions of an 
Excel workbook.
}
\usage{
\S3method{with}{workbook}(data, expr, ...)
}
\arguments{
\item{data}{A workbook object, as returned by \link{loadWorkbook}.}
\item{expr}{expression to evaluate}
\item{...}{Additional arguments passed to \link[=readNamedRegion-methods]{readNamedRegion}}
}

\details{This method will read all named regions from the workbook when
creating the environment. Names in the workbook will be processed through
\link{make.names} to obtain the variable names.

Changes to the variables representing named regions will not affect the
workbook contents and need to be saved explicitly using
\link[=writeNamedRegion-methods]{writeNamedRegion} and
\link[=saveWorkbook-methods]{saveWorkbook}. If the workbook contains names that
do not map to R identifiers, }

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}

\seealso{
  \code{\link[=getDefinedNames-methods]{getDefinedNames}},
  \code{\link[=readNamedRegion-methods]{readNamedRegion}},
}
\examples{
# multiregion xlsx file from demoFiles subfolder of package XLConnect
demoExcelFile <- system.file("demoFiles/multiregion.xlsx",
                             package = "XLConnect")

# load workbook
wb <- loadWorkbook(demoExcelFile)

# named regions: Calendar, IQ, Iris
print(getDefinedNames(wb))

# named regions as variables
with(wb, { 
  print(Calendar)
  summary(IQ)
  summary(Iris)
})
}
\keyword{IO}
