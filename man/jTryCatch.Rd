\name{jTryCatch}
\alias{jTryCatch}
\title{
Standard Java exception handling for XLConnect
}
\description{
Catches Java exceptions to create nicer error messages.
}
\usage{
jTryCatch(...)
}
\arguments{
  \item{\dots}{
  Arguments to \code{\link{tryCatch}} - an expression and optionally the argument
  \code{finally} to specify an expression to be evaluated before returning or exiting.
}
}
\details{
This function specifies a Java Throwable error handler that catches any Java exceptions. Caught exceptions
are post-processed and re-thrown using \code{\link{stop}}.

}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\link{tryCatch}}, \code{\link{stop}}
}
\examples{
try(XLConnect:::jTryCatch(J("java.lang.Integer")$valueOf("not an integer")))
}
\keyword{error}
\keyword{internal}
