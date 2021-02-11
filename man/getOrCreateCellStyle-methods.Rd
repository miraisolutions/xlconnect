\name{getOrCreateCellStyle-methods}
\docType{methods}
\alias{getOrCreateCellStyle}
\alias{getOrCreateCellStyle-methods}
\alias{getOrCreateCellStyle,workbook,character-method}
\title{Retrieving or creating named cell styles}
\description{
Retrieves or creates cell styles in \code{\linkS4class{workbook}}s.
}
\usage{
\S4method{getOrCreateCellStyle}{workbook,character}(object,name)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{name}{The name of the \code{\linkS4class{cellstyle}} to retrieve or to create}
}

\details{
Retrieves an existing \code{\linkS4class{cellstyle}} if it exists or creates a new one if it
does not exist yet.
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}

\seealso{
\code{\linkS4class{workbook}}, \code{\linkS4class{cellstyle}}, \code{\link[=setCellStyle-methods]{setCellStyle}},
\code{\link[=createCellStyle-methods]{createCellStyle}}, \code{\link[=existsCellStyle-methods]{existsCellStyle}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("getOrCreateCellStyle.xlsx", create = TRUE)

# The first time, the style does not exist yet and gets created
myStyle <- getOrCreateCellStyle(wb, name = "MyStyle")

# The second time, we retrieve the already existing style 
myStyle <- getOrCreateCellStyle(wb, name = "MyStyle")

# clean up 
file.remove("getOrCreateCellStyle.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
