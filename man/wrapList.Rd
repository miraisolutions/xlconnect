\name{wrapList}
\alias{wrapList}
\title{
Wrapping of arguments in a list
}
\description{
Wraps the specified argument in a list.
}
\usage{
wrapList(x)
}
\arguments{
  \item{x}{
  Any object.
}
}
\details{
Wraps an argument in a list if it not already is a list.
Further extracts the 'jobj' slot from S4 classes (such as \code{\linkS4class{workbook}} 
and \code{\linkS4class{cellstyle}}).
}
\value{
\code{\link{list}} which wraps the provided argument if it not already was a \code{\link{list}}.
If the provided argument is a S4 class with a 'jobj' slot, the corresponding slot is extracted and
wrapped in a list.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\seealso{
\code{\link{xlcCall}}
}
\keyword{internal}
