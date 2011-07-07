\name{addImage-methods}
\docType{methods}
\alias{addImage}
\alias{addImage-methods}
\alias{addImage,workbook,character,character,logical-method}
\alias{addImage,workbook,character,character,missing-method}
\title{Adding images to a worksheet}
\description{
Adds an image to a worksheet using a named region.
}
\usage{
 \S4method{addImage}{workbook,character,character,logical}(object, filename, name, originalSize)
}

\arguments{
  \item{object}{The \code{\linkS4class{workbook}} to use}
  \item{filename}{Name of the image file. Supported are images of 
  the following formats: JPG/JPEG, PNG, WMF, EMF, BMP, PICT.}
  \item{name}{Name of the named region that the image is set to}
  \item{originalSize}{If \code{originalSize = TRUE}, the image is
  inserted in the top left corner of the named region and not
  scaled. Otherwise, the image is scaled to fit the named region.
  The default value for \code{originalSize} is \code{FALSE}.}
}

\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
There is an known issue in Apache POI with adding images to xls workbooks. The result of adding
images to workbooks that already contain shapes or images may be that previous images are
removed or that existing images are replaced with newly added ones. It is therefore advised that
you use the \code{addImage} functionality only with workbooks that have no existing shapes or images.
Note that this only holds for xls workbooks (Excel 97-2003) and not for xlsx (Excel 2007+). There should
be no issues with xlsx workbooks.
}
\seealso{
\code{\linkS4class{workbook}}, \code{\link[=createName-methods]{createName}}
}
\examples{
## Write an R plot to a specified named region
## This example makes use of the 'Tonga Trench Earthquakes' example

# Load workbook (create if not existing)
wb <- loadWorkbook("earthquake.xlsx", create = TRUE)

# Create a sheet named 'earthquake'
createSheet(wb, name = "earthquake")

# Create a named region called 'earthquake' referring to the sheet
# called 'earthquake' 
createName(wb, name = "earthquake", formula = "earthquake!$B$2")

# Create R plot to a png device
require(lattice)
png(filename = "earthquake.png", width = 800, height = 600)
devAskNewPage(ask = FALSE)

Depth <- equal.count(quakes$depth, number=8, overlap=.1)
xyplot(lat ~ long | Depth, data = quakes)
update(trellis.last.object(),
       strip = strip.custom(strip.names = TRUE, strip.levels = TRUE),
       par.strip.text = list(cex = 0.75),
       aspect = "iso")

dev.off()

# Write image to the named region created above using the image's
# original size; i.e. the image's top left corner will match the
# specified cell's top left corner 
addImage(wb, filename = "earthquake.png", name = "earthquake",
         originalSize = TRUE)

# Save workbook (this actually writes the file to disk)
saveWorkbook(wb)
}
\keyword{methods}
\keyword{utilities}
