\name{setStyleNamePrefix-methods}
\docType{methods}
\alias{setStyleNamePrefix}
\alias{setStyleNamePrefix-methods}
\alias{setStyleNamePrefix,workbook-method}
\title{Setting the style name prefix for the "name prefix" style action}
\description{
Sets the style name prefix for the "name prefix" style action.
}
\usage{
\S4method{setStyleNamePrefix}{workbook}(object,prefix)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{prefix}{The name prefix}
}

\details{
Sets the \code{prefix} for the "name prefix" style action. See the method \code{\link[=setStyleAction-methods]{setStyleAction}} for more information.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\linkS4class{cellstyle}}, \code{\link[=setStyleAction-methods]{setStyleAction}}, 
\code{\link[=createCellStyle-methods]{createCellStyle}}
}
\keyword{methods}
\keyword{utilities}
