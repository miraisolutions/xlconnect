\name{XLConnect-package}
\alias{XLConnect-package}
\alias{XLConnect}
\docType{package}
\title{
Excel Connector for R
}
\description{
Provides comprehensive functionality to read, write and format Excel data.
}
\details{
For an overview over the package please refer to the available demos:\cr
\code{demo(package = "XLConnect")}
}
\author{
Mirai Solutions GmbH, \email{xlconnect@mirai-solutions.com}
}
\references{
Mirai Solutions GmbH: \url{https://mirai-solutions.ch}\cr
XLConnect on GitHub: \url{https://github.com/miraisolutions/xlconnect}
Mirai Solutions on GitHub: \url{https://github.com/miraisolutions}\cr
Apache POI: \url{https://poi.apache.org}
}
\keyword{ package }
\examples{\dontrun{
# Load workbook; create if not existing
wb <- loadWorkbook("XLConnect.xlsx", create = TRUE)

# Create a worksheet
createSheet(wb, name = "mtcars")

# Create a name reference
createName(wb, name = "mtcars", formula = "mtcars!$C$5")

# Write built-in data.frame 'mtcars' to the specified named region
writeNamedRegion(wb, mtcars, name = "mtcars")

# Save workbook
saveWorkbook(wb)

# clean up 
file.remove("XLConnect.xlsx")
}
}
