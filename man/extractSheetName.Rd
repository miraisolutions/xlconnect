\name{extractSheetName}
\alias{extractSheetName}
\title{
Extracting the sheet name from a formula
}
\description{
Extracts the sheet name from a formula of the form <SHEET_NAME>!<CELL_ADDRESS>
}
\usage{
extractSheetName(formula)
}
\arguments{
  \item{formula}{Formula string of the form <SHEET_NAME>!<CELL_ADDRESS>. Note that the validity
  of the formula won't be checked.}
}
\value{
Returns the name of the sheet referenced in the formula. For quoted sheet names (required if names
contain e.g. whitespaces or exclamation marks (!)) in formulas the function returns the unquoted
name.
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\examples{
\dontrun{
extractSheetName(c("MySheet!$A$1", "'My Sheet'!$A$1", "'My!Sheet'!$A$1"))
}
}
\keyword{utilities}
