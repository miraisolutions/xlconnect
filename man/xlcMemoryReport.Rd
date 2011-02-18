\name{xlcMemoryReport}
\alias{xlcMemoryReport}
\title{
Reporting free Java Virtual Machine memory
}
\description{
Reports the amount of free memory in the Java Virtual Machine (JVM).
}
\usage{
xlcMemoryReport()
}
\details{
This function uses Java's Runtime class to query the free JVM memory.
}
\references{
java.lang.Runtime:\cr
\url{http://download.oracle.com/javase/6/docs/api/java/lang/Runtime.html}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\link{xlcFreeMemory}}
}
\examples{
xlcMemoryReport()
}
\keyword{utilities}
