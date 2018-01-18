\name{existsCellStyle-methods}
\docType{methods}
\alias{existsCellStyle}
\alias{existsCellStyle-methods}
\alias{existsCellStyle,workbook-method}
\title{Retrieving named cell styles}
\description{
Checks whether a named cell style exists in a \code{\linkS4class{workbook}}.
}
\usage{
\S4method{existsCellStyle}{workbook}(object,name)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The name of the \code{\linkS4class{cellstyle}} to check}
}

\details{
Checks whether the \code{\linkS4class{cellstyle}} with the specified \code{name} exists. 
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}

\seealso{
\code{\linkS4class{workbook}}, \code{\linkS4class{cellstyle}}, \code{\link[=setCellStyle-methods]{setCellStyle}},
\code{\link[=createCellStyle-methods]{createCellStyle}}, \code{\link[=getOrCreateCellStyle-methods]{getOrCreateCellStyle}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("existsCellStyle.xlsx", create = TRUE)

# Cell style 'MyStyle' does not exist yet
stopifnot(!existsCellStyle(wb, "MyStyle"))

# Create the style "MyStyle"
createCellStyle(wb, "MyStyle")

# And now it is here
stopifnot(existsCellStyle(wb, "MyStyle"))
}
\keyword{methods}
\keyword{utilities}
