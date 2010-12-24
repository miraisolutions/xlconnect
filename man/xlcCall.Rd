\name{xlcCall}
\alias{xlcCall}
\title{
Automatic argument vectorization & default Java exception handling for XLConnect
}
\description{
Performs automatic argument vectorization and provides default Java
exception handling for \pkg{XLConnect}.
}
\usage{
xlcCall(fun, ..., SIMPLIFY = TRUE)
}
\arguments{
  \item{fun}{
  Function to be called.
}
  \item{\dots}{
  Arguments to the function to be called. These arguments will be automatically vectorized using \code{\link{mapply}}.
}
  \item{SIMPLIFY}{
  Argument to \code{\link{mapply}}. Attempt to reduce result of function call to a vector or matrix?
}
}
\details{
This function makes use of \code{\link{mapply}} to perform automatic argument vectorization. Non-atomic arguments are
wrapped in a list using \code{\link{wrapList}} first. The list of the resulting arguments is then being vectorized
using \code{\link{mapply}} before the specified function is being called. The function call is wrapped with
\code{\link{jTryCatch}} to catch Java exceptions.
}
\value{
Result of the specified function by calling it with the vectorized argument list.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\link{wrapList}}, \code{\link{jTryCatch}}
}
\keyword{internal}
