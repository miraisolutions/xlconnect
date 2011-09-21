\name{$-methods}
\docType{methods}
\alias{$}
\alias{$-methods}
\alias{$,workbook-method}
\alias{$,cellstyle-method}
\title{Executing workbook methods in object$method(...) form}
\description{
Allows to execute \code{\linkS4class{workbook}} methods in \code{workbook-object$method(...)} form.
}
\arguments{
  \item{x}{The object (\code{\linkS4class{workbook}}, \code{\linkS4class{cellstyle}}) to use}
}
\details{
\code{x$method(...)} (where \code{x} is a \code{\linkS4class{workbook}}-object) is equivalent to \code{method(x, ...)}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
The \code{\linkS4class{workbook}} \code{$}-operator allows to call \code{\linkS4class{workbook}}-methods in \code{workbook-object$method(...)} form. 
This form might be considered more convenient or readable for programmers coming from other object-oriented languages such as Java, C#, ...
}
\seealso{
\code{\linkS4class{workbook}}
}
\examples{
# Load workbook (create if not existing)
wb <- loadWorkbook("dollar.xlsx", create = TRUE)

# Create a worksheet called 'CO2'
wb$createSheet(name = "CO2")

# Write built-in data set 'CO2' to the worksheet created above
wb$writeWorksheet(CO2, sheet = "CO2", startRow = 4, startCol = 2)

# Save workbook
wb$saveWorkbook()
}
\keyword{methods}
\keyword{utilities}
