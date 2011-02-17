\name{setStyleNamePrefix-methods}
\docType{methods}
\alias{setStyleNamePrefix}
\alias{setStyleNamePrefix-methods}
\alias{setStyleNamePrefix,workbook,character-method}
\title{Setting the style name prefix for the "name prefix" style action}
\description{
Sets the style name prefix for the "name prefix" style action.
}
\usage{
\S4method{setStyleNamePrefix}{workbook,character}(object,prefix)
}

\arguments{
  \item{object}{The workbook to use}
  \item{prefix}{The name prefix}
}

\description{
Sets the \code{prefix} for the "name prefix" style action. See the method \code{\link[=setStyleAction-methods]{setStyleAction}} for more information.
}
\author{
Martin Studer\cr
\href{http://www.mirai-solutions.com}{Mirai Solutions GmbH}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\linkS4class{cellstyle}}, \code{\link[=setStyleAction-methods]{setStyleAction}}, 
\code{\link[=createCellStyle-methods]{createCellStyle}}
}
\keyword{methods}
\keyword{utilities}
