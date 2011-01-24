\name{workbook-class}
\Rdversion{1.1}
\docType{class}
\alias{workbook-class}
\alias{$,workbook-method}
\alias{addImage,workbook,character,character,logical-method}
\alias{addImage,workbook,character,character,missing-method}
\alias{createCellStyle,workbook,character-method}
\alias{createCellStyle,workbook,missing-method}
\alias{createName,workbook,character,character,logical-method}
\alias{createName,workbook,character,character,missing-method}
\alias{createSheet,workbook,character-method}
\alias{existsName,workbook,character-method}
\alias{existsSheet,workbook,character-method}
\alias{getActiveSheetIndex,workbook-method}
\alias{getActiveSheetName,workbook-method}
\alias{getDefinedNames,workbook,logical-method}
\alias{getDefinedNames,workbook,missing-method}
\alias{getReferenceFormula,workbook,character-method}
\alias{getSheets,workbook-method}
\alias{hideSheet,workbook,character,logical-method}
\alias{hideSheet,workbook,character,missing-method}
\alias{hideSheet,workbook,numeric,logical-method}
\alias{hideSheet,workbook,numeric,missing-method}
\alias{isSheetHidden,workbook,character-method}
\alias{isSheetHidden,workbook,numeric-method}
\alias{isSheetVeryHidden,workbook,character-method}
\alias{isSheetVeryHidden,workbook,numeric-method}
\alias{isSheetVisible,workbook,character-method}
\alias{isSheetVisible,workbook,numeric-method}
\alias{print,workbook-method}
\alias{readNamedRegion,workbook,character,logical-method}
\alias{readNamedRegion,workbook,character,missing-method}
\alias{readWorksheet,workbook,character,missing,missing,missing,missing,logical-method}
\alias{readWorksheet,workbook,character,missing,missing,missing,missing,missing-method}
\alias{readWorksheet,workbook,character,numeric,numeric,numeric,numeric,logical-method}
\alias{readWorksheet,workbook,character,numeric,numeric,numeric,numeric,missing-method}
\alias{readWorksheet,workbook,numeric,missing,missing,missing,missing,logical-method}
\alias{readWorksheet,workbook,numeric,missing,missing,missing,missing,missing-method}
\alias{readWorksheet,workbook,numeric,numeric,numeric,numeric,numeric,logical-method}
\alias{readWorksheet,workbook,numeric,numeric,numeric,numeric,numeric,missing-method}
\alias{removeName,workbook,character-method}
\alias{removeSheet,workbook,character-method}
\alias{removeSheet,workbook,numeric-method}
\alias{saveWorkbook,workbook-method}
\alias{setActiveSheet,workbook,character-method}
\alias{setActiveSheet,workbook,numeric-method}
\alias{setCellStyle,workbook,character,numeric,numeric,cellstyle-method}
\alias{setCellStyle,workbook,numeric,numeric,numeric,cellstyle-method}
\alias{setColumnWidth,workbook,character,numeric,numeric-method}
\alias{setColumnWidth,workbook,numeric,numeric,numeric-method}
\alias{setRowHeight,workbook,character,numeric,numeric-method}
\alias{setRowHeight,workbook,numeric,numeric,numeric-method}
\alias{setStyleAction,workbook,character-method}
\alias{setStyleNamePrefix,workbook,character-method}
\alias{show,workbook-method}
\alias{summary,workbook-method}
\alias{unhideSheet,workbook,character-method}
\alias{unhideSheet,workbook,numeric-method}
\alias{writeNamedRegion,workbook,ANY,character,logical-method}
\alias{writeNamedRegion,workbook,ANY,character,missing-method}
\alias{writeWorksheet,workbook,ANY,character,missing,missing,logical-method}
\alias{writeWorksheet,workbook,ANY,character,missing,missing,missing-method}
\alias{writeWorksheet,workbook,ANY,character,numeric,numeric,logical-method}
\alias{writeWorksheet,workbook,ANY,character,numeric,numeric,missing-method}
\alias{writeWorksheet,workbook,ANY,numeric,missing,missing,logical-method}
\alias{writeWorksheet,workbook,ANY,numeric,missing,missing,missing-method}
\alias{writeWorksheet,workbook,ANY,numeric,numeric,numeric,logical-method}
\title{Class "workbook"}
\description{
This is \pkg{XLConnect}'s main entity representing a Microsoft Excel workbook. S4 objects of this class
and corresponding methods are used to manipulate the underlying Excel workbook instances.
}
\section{Objects from the Class}{
Objects can be created by calls of the form \code{\link{loadWorkbook}(filename, create)}. This is
a shortcut form of \code{new("workbook", filename, create)} with some additional error checking.
}
\section{Slots}{
  \describe{
    \item{\code{filename}:}{Object of class \code{character} which represents the filename of the
    underlying Microsoft Excel workbook.}
    \item{\code{jobj}:}{Object of class \code{jobjRef} (see package \pkg{rJava}) which represents a Java object reference that is used
    in the back-end to manipulate the underlying Excel workbook instance.}
  }
  These slots should not be accessed directly in user code. \code{workbook} objects should only be manipulated
  via the corresponding methods.
}
\references{
\url{http://en.wikipedia.org/wiki/Office_Open_XML}
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\note{
\pkg{XLConnect} supports both Excel '97 (*.xls) and OOXML (Excel 2007+, *.xlsx) file formats.\cr\cr

A \code{workbook}'s underlying Excel file is not saved (or being created in case the file did not exist and \code{create = TRUE} has been
specified) unless the \code{\link[=saveWorkbook-methods]{saveWorkbook}} method has been called on the object. This provides more flexibility to the user to decide
when changes are saved and also provides better performance in that several changes can be written in one go (normally at the end, rather than 
after every operation causing the file to be rewritten again completely each time). This is due to the fact that workbooks are manipulated in-memory and
are only written to disk with specifically calling \code{\link[=saveWorkbook-methods]{saveWorkbook}}.
}
\seealso{
\code{\link{loadWorkbook}}, \code{\link[=saveWorkbook-methods]{saveWorkbook}}
}
\examples{
# Create a new workbook 'myWorkbook.xlsx' (assuming the file to not exist already)
wb <- loadWorkbook("myWorkbook.xlsx", create = TRUE)
# Create a worksheet called 'mtcars'
createSheet(wb, name = "mtcars")
# Write built-in dataset 'mtcars' to sheet 'mtcars' created above
writeWorksheet(wb, mtcars, sheet = "mtcars")
# Save workbook - this actually writes the file 'myWorkbook.xlsx' to disk
saveWorkbook(wb)
}
\keyword{classes}
