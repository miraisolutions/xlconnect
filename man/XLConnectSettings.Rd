\name{XLConnectSettings}
\alias{XLConnectSettings}
\title{
Performing general settings for XLConnect
}
\description{
Performs general settings as used by \pkg{XLConnect}.
}
\usage{
XLConnectSettings(pdesc)
}
\arguments{
  \item{pdesc}{
  An object of class packageDescription for the \pkg{XLConnect} package.
}
}
\details{
This function is called by \pkg{XLConnect}'s \code{.onLoad} function to perform general settings
used by \pkg{XLConnect}.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\keyword{internal}
