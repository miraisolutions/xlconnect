\name{extraction-methods}
\docType{methods}
\alias{[}
\alias{[[}
\alias{[<-}
\alias{[[<-}
\alias{[-methods}
\alias{[[-methods}
\alias{[<--methods}
\alias{[[<--methods}
\alias{[,workbook-method}
\alias{[[,workbook-method}
\alias{[<-,workbook-method}
\alias{[[<-,workbook-method}
\title{Workbook data extraction & replacement operators}
\description{
Operators that allow to extract/replace data from/on a \code{\linkS4class{workbook}}.
}
\arguments{
  \item{x}{The \code{\linkS4class{workbook}} object to use}
  \item{i}{Name of worksheet (\code{[}, \code{[<-}) or name of Excel name (\code{[[}, \code{[[<-})
  to extract or replace}
  \item{j}{Only used with \code{[[<-}: Optional formula to define the Excel name if it does not yet exist
  on the workbook.}
  \item{drop}{Not used}
  \item{value}{Data object used for replacement}
  \item{...}{Arguments passed to the corresponding underlying function to read/write the data}
}
\details{
The workbook extraction operators are basically syntactic sugar for the common methods \code{\link[=readWorksheet-methods]{readWorksheet}} (\code{[}), 
\code{\link[=writeWorksheet-methods]{writeWorksheet}} (\code{[<-}), \code{\link[=readNamedRegion-methods]{readNamedRegion}} (\code{[[}), 
\code{\link[=writeNamedRegion-methods]{writeNamedRegion}} (\code{[[<-}).
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{https://mirai-solutions.ch}
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=readWorksheet-methods]{readWorksheet}}, \code{\link[=writeWorksheet-methods]{writeWorksheet}},
\code{\link[=readNamedRegion-methods]{readNamedRegion}}, \code{\link[=writeNamedRegion-methods]{writeNamedRegion}}
}
\examples{\dontrun{
# Load workbook (create if not existing)
wb <- loadWorkbook("extraction.xlsx", create = TRUE)

# Write mtcars data set on a worksheet named 'mtcars1'.
# Note: The 'mtcars1' sheet will be created automatically if it does
# not exist yet. Also, default values for other writeWorksheet arguments
# hold, i.e. the data set is written starting at the top left corner. 
wb["mtcars1"] = mtcars

# Write mtcars data set on a worksheet named 'mtcars2'.
# Again, the 'mtcars2' worksheet is created automatically.
# Additionally specify arguments passed to the underlying method
# writeWorksheet.
wb["mtcars2", startRow = 6, startCol = 11, header = FALSE] = mtcars

# Read worksheets 'mtcars1' and 'mtcars2'.
# Note: The default arguments hold for the underlying method
# readWorksheet.
wb["mtcars1"]
wb["mtcars2"]

# Write mtcars data set to a named region named 'mtcars3'. Since
# it doesn't exist yet we also need to specify the formula to
# define it. Also note that the sheet 'mtcars3' referenced in the
# formula does not yet exist - it will be created automatically!
# Moreover, default values for other writeNamedRegion arguments hold.
wb[["mtcars3", "mtcars3!$B$7"]] = mtcars

# Redefine named region 'mtcars3'. Note that no formula specification
# is required since named region is already defined (see above example).
wb[["mtcars3"]] = mtcars

# Write mtcars data set to a named region 'mtcars4'. Since the named
# region does not yet exist a formula specification is required. Also,
# additional arguments are specified that are passed to the underlying
# method writeNamedRegion.
wb[["mtcars4", "mtcars4!$D$8", rownames = "Car"]] = mtcars

# Read the named regions 'mtcars3' and 'mtcars4'.
# Note: Default values hold for the underlying method readNamedRegion.
wb[["mtcars3"]]
wb[["mtcars4"]]

# clean up 
file.remove("extraction.xlsx")
}
}
\keyword{methods}
\keyword{utilities}
