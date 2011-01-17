\name{runUnitTests}
\alias{runUnitTests}
\title{
XLConnect Unit Testing Framework
}
\description{
Runs XLConnect's unit testing suite.
}
\usage{
runUnitTests()
}
\details{
This function runs XLConnect's unit test suite as defined by the unit tests implemented in the unitTests subfolder of 
the \pkg{XLConnect} library. Text and HTML protocols are produced - the HTML version should be opened automatically once
the test suite completed execution.
}
\references{
\url{http://rwiki.sciviews.org/doku.php?id=developers:runit}, \url{http://cran.r-project.org/web/packages/RUnit/index.html}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
Running XLConnect's unit test suite requires the \pkg{RUnit} package to be installed.
}
\examples{
\donttest{runUnitTests()}
}
\keyword{utilities}
\keyword{misc}
