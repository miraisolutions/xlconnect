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
\tabular{ll}{
Package: \tab XLConnect\cr
Type: \tab Package\cr
Version: \tab 0.2-13\cr
Date: \tab 2017-05-10\cr
URL: \tab \url{http://www.mirai-solutions.com}\cr
\tab \url{http://miraisolutions.wordpress.com}\cr
SystemRequirements: \tab java (>= 1.6)\cr
Depends: \tab R (>= 2.10.0), XLConnectJars (== 0.2-12)\cr
Imports: \tab methods, rJava, utils\cr
Suggests: \tab RUnit, lattice, ggplot2 (>= 0.9.3), zoo\cr
License: \tab GPL-3\cr
Copyright: \tab See file COPYRIGHTS\cr
LazyData: \tab yes\cr
}
For an overview over the package please refer to the available demos:\cr
\code{demo(package = "XLConnect")}
}
\author{
Mirai Solutions GmbH, \email{xlconnect@mirai-solutions.com}
}
\references{
Mirai Solutions GmbH: \url{http://www.mirai-solutions.com}\cr
Mirai Solutions on GitHub: \url{https://github.com/miraisolutions}\cr
Mirai Solutions Blog: \url{http://miraisolutions.wordpress.com}\cr
Apache POI: \url{http://poi.apache.org}
}
\keyword{ package }
\examples{
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
}
