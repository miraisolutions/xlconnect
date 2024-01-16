\name{xlcCall}
\alias{xlcCall}
\title{
Automatic argument vectorization, default Java exception and warnings handling for XLConnect
}
\description{
Performs automatic argument vectorization, provides default Java
exception and warnings handling for \pkg{XLConnect}.
}
\usage{
xlcCall(obj, fun, ..., .recycle = TRUE, .simplify = TRUE, .checkWarnings = TRUE)
}
\arguments{
  \item{obj}{
  \pkg{XLConnect} S4 object (e.g \code{\linkS4class{workbook}}, \code{\linkS4class{cellstyle}}) which
  has a slot \code{jobj} representing a \pkg{rJava} \code{jobjRef} object whose method represented
  by the argument \code{fun} should be called.
}
  \item{fun}{
  Name of Java method to be called.
}
  \item{\dots}{
  Arguments to the Java method to be called.
}
  \item{.recycle}{
  If \code{TRUE} (default), arguments passed through \dots will be automatically vectorized using \code{\link{mapply}} and
  \code{fun} will be called the appropriate amount of times.
  If \code{FALSE}, no vectorization will take place in R and \code{fun} will be called once only. This may be used to implement
  vectorization on the Java side.
  }
  \item{.simplify}{
  Only relevant if \code{.recycle = TRUE}. Argument to internal \code{\link{mapply}} to set if there should be an attempt to 
  reduce the result to a vector or matrix.
}
  \item{.checkWarnings}{
  Whether to check for warnings registered on the underlying Java object.
  }
}
\details{
This function makes use of \code{\link{mapply}} to perform automatic argument vectorization. Non-atomic arguments are
wrapped in a list using \code{\link{wrapList}} first. The list of the resulting arguments is then being vectorized
using \code{\link{mapply}} before the specified function is being called. The function call is wrapped with
\code{\link{jTryCatch}} to catch Java exceptions. Further, any warnings are retrieved from the Java side and are logged
via R's standard warning mechanism.
}
\value{
Result of the specified function by calling it with the vectorized argument list.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\link{wrapList}}, \code{\link{jTryCatch}}
}
\keyword{internal}
