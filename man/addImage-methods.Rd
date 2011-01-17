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
\section{Methods}{
\describe{

\item{\code{signature(object = "workbook", filename = "character", name = "character", originalSize = "logical")}}{
Adds the image specified by the \code{filename} to the named region specified by \code{name}. If \code{originalSize = FALSE},
the image will be fitted to the area specified by the named region, otherwise top left corners of the image and the named region
will be aligned and the image will keep its original size.\cr
Supported are images of the following formats: JPG/JPEG, PNG, WMF, EMF, BMP, PICT.
}

\item{\code{signature(object = "workbook", filename = "character", name = "character", originalSize = "missing")}}{
As above with \code{originalSize = FALSE}.
}
}}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
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

# Create a named region called 'earthquake' referring to the sheet called 'earthquake'
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

# Write image to the named region created above using the image's original size;
# i.e. the image's top left corner will match the specified cell's top left corner
addImage(wb, filename = "earthquake.png", name = "earthquake", originalSize = TRUE)
}
\keyword{methods}
\keyword{utilities}
