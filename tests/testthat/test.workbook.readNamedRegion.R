test_that("test.workbook.readNamedRegion", {
  # Create workbooks
  wb.xls <- loadWorkbook(test_path("resources/testWorkbookReadNamedRegion.xls"), create = FALSE)
  wb.xlsx <- loadWorkbook(test_path("resources/testWorkbookReadNamedRegion.xlsx"), create = FALSE)
  checkDf <- data.frame(
    NumericColumn = c(-23.63, NA, NA, 5.8, 3),
    StringColumn = c("Hello", NA, NA, NA, "World"),
    BooleanColumn = c(TRUE, FALSE, FALSE, NA, NA),
    DateTimeColumn = as.POSIXct(c(NA, NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07")),
    stringsAsFactors = F
  )
  options(XLConnect.setCustomAttributes = FALSE)

  # Check that the read named region equals the defined data.frame (*.xls)
  res <- readNamedRegion(wb.xls, "Test", header = TRUE)
  expect_equal(checkDf, res)

  # Check that the read named region equals the defined data.frame (*.xlsx)
  res <- readNamedRegion(wb.xlsx, "Test", header = TRUE)
  expect_equal(checkDf, res)

  # Check that the same works when explicitly specifying global scope (*.xls)
  res <- readNamedRegion(wb.xls, "Test", header = TRUE, worksheetScope = "")
  expect_equal(checkDf, res)

  # Check that the same works when explicitly specifying global scope (*.xlsx)
  res <- readNamedRegion(wb.xlsx, "Test", header = TRUE, worksheetScope = "")
  expect_equal(checkDf, res)

  # check that global ranges are not found when specifying a worksheet name (*.xls)
  expect_error(readNamedRegion(wb.xls, name = "Test", header = TRUE, worksheetScope = "Test"))

  # check that global ranges are not found when specifying a worksheet name (*.xlsx)
  expect_error(readNamedRegion(wb.xlsx, name = "Test", header = TRUE, worksheetScope = "Test"))

  # Check that attempting to read a non-existing named region throws an exception (*.xls)
  expect_error(readNamedRegion(wb.xls, "NameThatDoesNotExist"))

  # Check that attempting to read a non-existing named region throws an exception (*.xlsx)
  expect_error(readNamedRegion(wb.xlsx, "NameThatDoesNotExist"))
  targetNoForce <- data.frame(
    AAA = c(NA, NA, NA, 780.9, NA),
    BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
    CCC = c(TRUE, NA, NA, NA, NA),
    DDD = as.POSIXct(c("1984-01-11 12:00:00", NA, NA, NA, NA)),
    stringsAsFactors = FALSE
  )
  targetForce <- data.frame(
    AAA = c(-14.65, NA, 11.7, 780.9, NA),
    BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
    CCC = c(TRUE, TRUE, NA, FALSE, FALSE),
    DDD = as.POSIXct(c("1984-01-11 12:00:00", "2012-02-06 16:15:23", "1984-01-11 12:00:00", NA, "1900-12-22 16:04:48")),
    stringsAsFactors = FALSE
  )
  # Check that conversion performs ok (without forcing conversion; *.xls)
  res <- suppressWarnings(readNamedRegion(
    wb.xls,
    name = "Conversion",
    header = TRUE,
    colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
    forceConversion = FALSE,
    dateTimeFormat = "%d.%m.%Y %H:%M:%S"
  ))
  expect_equal(targetNoForce, res)

  # Check that conversion performs ok (without forcing conversion; *.xlsx)
  res <- suppressWarnings(readNamedRegion(
    wb.xlsx,
    name = "Conversion",
    header = TRUE,
    colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
    forceConversion = FALSE,
    dateTimeFormat = "%d.%m.%Y %H:%M:%S"
  ))
  expect_equal(targetNoForce, res)

  # Check that conversion performs ok (with forcing conversion; *.xls)
  res <- suppressWarnings(readNamedRegion(
    wb.xls,
    name = "Conversion",
    header = TRUE,
    colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
    forceConversion = TRUE,
    dateTimeFormat = "%d.%m.%Y %H:%M:%S"
  ))
  expect_equal(targetForce, res)

  # Check that conversion performs ok (with forcing conversion; *.xlsx)
  res <- suppressWarnings(readNamedRegion(
    wb.xlsx,
    name = "Conversion",
    header = TRUE,
    colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
    forceConversion = TRUE,
    dateTimeFormat = "%d.%m.%Y %H:%M:%S"
  ))
  expect_equal(targetForce, res)
  target <- list(
    AAA = data.frame(A = 1:3, B = letters[1:3], C = c(TRUE, TRUE, FALSE), stringsAsFactors = FALSE),
    BBB = data.frame(D = 4:6, E = letters[4:6], F = c(FALSE, TRUE, TRUE), stringsAsFactors = FALSE)
  )
  # Check that reading multiple named regions returns a named list (*.xls)
  res <- readNamedRegion(wb.xls, name = c("AAA", "BBB"), header = TRUE)
  expect_equal(target, res)

  # Check that reading multiple named regions returns a named list (*.xlsx)
  res <- readNamedRegion(wb.xlsx, name = c("AAA", "BBB"), header = TRUE)
  expect_equal(target, res)
  target <- data.frame(
    `With whitespace` = 1:4,
    `And some other funky characters: _=?^~!$@#%§` = letters[1:4],
    check.names = FALSE,
    stringsAsFactors = FALSE
  )
  # Check that reading named regions with check.names = FALSE works (*.xls)
  res <- readNamedRegion(wb.xls, name = "VariableNames", header = TRUE, check.names = FALSE)
  expect_equal(target, res)

  # Check that reading named regions with check.names = FALSE works (*.xlsx)
  res <- readNamedRegion(wb.xlsx, name = "VariableNames", header = TRUE, check.names = FALSE)
  expect_equal(target, res)
  # Check that attempting to specify both keep and drop throws an exception (*.xls)
  expect_error(readNamedRegion(wb.xls, "Conversion", header = TRUE, keep = c("AAA", "BBB"), drop = c("CCC", "DDD")))

  # Check that attempting to specify both keep and drop throws an exception (*.xlsx)
  expect_error(readNamedRegion(wb.xlsx, "Conversion", header = TRUE, keep = c("AAA", "BBB"), drop = c("CCC", "DDD")))

  # Check that attempting to keep a non-existing column (indicated by header name) throws an exception (*.xls)
  expect_error(readNamedRegion(wb.xls, "Conversion", header = TRUE, keep = c("AAA", "BBB", "ZZZ")))

  # Check that attempting to keep a non-existing column (indicated by header name) throws an exception (*.xlsx)
  expect_error(readNamedRegion(wb.xlsx, "Conversion", header = TRUE, keep = c("AAA", "BBB", "ZZZ")))

  # Check that attempting to keep a column (indicated by index) out of named region bounds throws an exception (*.xls)
  expect_error(readNamedRegion(wb.xls, "Conversion", header = TRUE, keep = c(1, 2, 5)))

  # Check that attempting to keep a column (indicated by index) out of named region bounds throws an exception (*.xlsx)
  expect_error(readNamedRegion(wb.xlsx, "Conversion", header = TRUE, keep = c(1, 2, 5)))

  # Check that attempting to drop a non-existing column (indicated by header name) throws an exception (*.xls)
  expect_error(readNamedRegion(wb.xls, "Conversion", header = TRUE, drop = c("AAA", "BBB", "ZZZ")))

  # Check that attempting to drop a non-existing column (indicated by header name) throws an exception (*.xlsx)
  expect_error(readNamedRegion(wb.xlsx, "Conversion", header = TRUE, drop = c("AAA", "BBB", "ZZZ")))

  # Check that attempting to drop a column (indicated by index) out of named region bounds throws an exception (*.xls)
  expect_error(readNamedRegion(wb.xls, "Conversion", header = TRUE, drop = c(1, 2, 5)))

  # Check that attempting to drop a column (indicated by index) out of named region bounds throws an exception (*.xlsx)
  expect_error(readNamedRegion(wb.xlsx, "Conversion", header = TRUE, drop = c(1, 2, 5)))
  AAA <- data.frame(A = 1:3, B = letters[1:3], C = c(TRUE, TRUE, FALSE), stringsAsFactors = FALSE)
  BBB <- data.frame(D = 4:6, E = letters[4:6], F = c(FALSE, TRUE, TRUE), stringsAsFactors = FALSE)
  # Check that keeping columns "NumericColumn" and "BooleanColumn" (= by name) with header=FALSE throws an exception (*.xls)
  expect_error(readNamedRegion(wb.xls, "Test", header = FALSE, keep = c("NumericColumn", "BooleanColumn")))

  # Check that keeping columns "NumericColumn" and "BooleanColumn" (= by name) with header=FALSE throws an exception (*.xlsx)
  expect_error(readNamedRegion(wb.xlsx, "Test", header = FALSE, keep = c("NumericColumn", "BooleanColumn")))

  # Check that dropping columns "StringColumn" and "DateTimeColumn" (= by name) with header=FALSE throws an exception (*.xls)
  expect_error(readNamedRegion(wb.xls, "Test", header = FALSE, drop = c("StringColumn", "DateTimeColumn")))

  # Check that dropping columns "StringColumn" and "DateTimeColumn" (= by name) with header=FALSE throws an exception (*.xlsx)
  expect_error(readNamedRegion(wb.xlsx, "Test", header = FALSE, drop = c("StringColumn", "DateTimeColumn")))
  # Keeping the same columns from multiple named regions (*.xls)
  res <- readNamedRegion(wb.xls, name = c("Test", "AAA", "BBB"), header = TRUE, keep = c(1, 3))
  expect_equal(list(Test = checkDf[, c(1, 3)], AAA = AAA[, c(1, 3)], BBB = BBB[, c(1, 3)]), res)

  # Keeping the same columns from multiple named regions (*.xlsx)
  res <- readNamedRegion(wb.xlsx, name = c("Test", "AAA", "BBB"), header = TRUE, keep = c(1, 3))
  expect_equal(list(Test = checkDf[, c(1, 3)], AAA = AAA[, c(1, 3)], BBB = BBB[, c(1, 3)]), res)
  # Testing the correct replication of the keep argument (reading from 3 named regions, while keep has length 2) (*.xls)
  res <- readNamedRegion(wb.xls, name = c("Test", "AAA", "BBB"), header = TRUE, keep = list(1, 3))
  expect_equal(list(Test = checkDf[1], AAA = AAA[3], BBB = BBB[1]), res)

  # Testing the correct replication of the keep argument (reading from 3 named regions, while keep has length 2) (*.xlsx)
  res <- readNamedRegion(wb.xlsx, name = c("Test", "AAA", "BBB"), header = TRUE, keep = list(1, 3))
  expect_equal(list(Test = checkDf[1], AAA = AAA[3], BBB = BBB[1]), res)
  # Keeping different columns from multiple named regions (*.xls)
  res <- readNamedRegion(wb.xls, name = c("Test", "AAA", "BBB"), header = TRUE, keep = list(c(1, 2), c(2, 3), c(1, 3)))
  expect_equal(list(Test = checkDf[, c(1, 2)], AAA = AAA[, c(2, 3)], BBB = BBB[, c(1, 3)]), res)

  # Keeping different columns from multiple named regions (*.xlsx)
  res <- readNamedRegion(wb.xlsx, name = c("Test", "AAA", "BBB"), header = TRUE, keep = list(c(1, 2), c(2, 3), c(1, 3)))
  expect_equal(list(Test = checkDf[, c(1, 2)], AAA = AAA[, c(2, 3)], BBB = BBB[, c(1, 3)]), res)
  # Keeping different columns from multiple named regions (2 keep list elements for 4 named regions) (*.xls)
  res <- readNamedRegion(wb.xls, name = c("Test", "AAA", "BBB", "Test"), header = TRUE, keep = list(c(1, 2), c(2, 3)))
  expect_equal(
    list(Test = checkDf[, c(1, 2)], AAA = AAA[, c(2, 3)], BBB = BBB[, c(1, 2)], Test = checkDf[, c(2, 3)]),
    res
  )

  # Keeping different columns from multiple named regions (2 keep list elements for 4 named regions) (*.xlsx)
  res <- readNamedRegion(wb.xlsx, name = c("Test", "AAA", "BBB", "Test"), header = TRUE, keep = list(c(1, 2), c(2, 3)))
  expect_equal(
    list(Test = checkDf[, c(1, 2)], AAA = AAA[, c(2, 3)], BBB = BBB[, c(1, 2)], Test = checkDf[, c(2, 3)]),
    res
  )
  # Dropping the same columns from multiple named regions (*.xls)
  res <- readNamedRegion(wb.xls, name = c("Test", "AAA", "BBB"), header = TRUE, drop = c(1, 2))
  expect_equal(
    list(
      Test = checkDf[, c(3, 4)],
      AAA = data.frame(C = AAA[, 3], stringsAsFactors = F),
      BBB = data.frame(F = BBB[, 3], stringsAsFactors = F)
    ),
    res
  )

  # Dropping the same columns from multiple named regions (*.xlsx)
  res <- readNamedRegion(wb.xlsx, name = c("Test", "AAA", "BBB"), header = TRUE, drop = c(1, 2))
  expect_equal(
    list(
      Test = checkDf[, c(3, 4)],
      AAA = data.frame(C = AAA[, 3], stringsAsFactors = F),
      BBB = data.frame(F = BBB[, 3], stringsAsFactors = F)
    ),
    res
  )
  # Testing the correct replication of the drop argument (reading from 3 named regions, while drop has length 2) (*.xls)
  res <- readNamedRegion(wb.xls, name = c("Test", "AAA", "BBB"), header = TRUE, drop = list(1, 2))
  expect_equal(list(Test = checkDf[, c(2, 3, 4)], AAA = AAA[, c(1, 3)], BBB = BBB[, c(2, 3)]), res)

  # Testing the correct replication of the drop argument (reading from 3 named regions, while drop has length 2) (*.xlsx)
  res <- readNamedRegion(wb.xlsx, name = c("Test", "AAA", "BBB"), header = TRUE, drop = list(1, 2))
  expect_equal(list(Test = checkDf[, c(2, 3, 4)], AAA = AAA[, c(1, 3)], BBB = BBB[, c(2, 3)]), res)
  # Dropping different columns from multiple named regions (*.xls)
  res <- readNamedRegion(wb.xls, name = c("Test", "AAA", "BBB"), header = TRUE, drop = list(c(1, 2), c(2, 3), c(1, 3)))
  expect_equal(
    list(
      Test = checkDf[, c(3, 4)],
      AAA = data.frame(A = AAA[, 1], stringsAsFactors = F),
      BBB = data.frame(E = BBB[, 2], stringsAsFactors = F)
    ),
    res
  )

  # Dropping different columns from multiple named regions (*.xlsx)
  res <- readNamedRegion(wb.xlsx, name = c("Test", "AAA", "BBB"), header = TRUE, drop = list(c(1, 2), c(2, 3), c(1, 3)))
  expect_equal(
    list(
      Test = checkDf[, c(3, 4)],
      AAA = data.frame(A = AAA[, 1], stringsAsFactors = F),
      BBB = data.frame(E = BBB[, 2], stringsAsFactors = F)
    ),
    res
  )
  # Dropping different columns from multiple named regions (2 drop list elements for 4 named regions) (*.xls)
  res <- readNamedRegion(wb.xls, name = c("Test", "AAA", "BBB", "Test"), header = TRUE, drop = list(c(1, 2), c(2, 3)))
  expect_equal(
    list(
      Test = checkDf[, c(3, 4)],
      AAA = data.frame(A = AAA[, 1], stringsAsFactors = F),
      BBB = data.frame(F = BBB[, 3], stringsAsFactors = F),
      Test = checkDf[, c(1, 4)]
    ),
    res
  )

  # Dropping different columns from multiple named regions (2 drop list elements for 4 named regions) (*.xlsx)
  res <- readNamedRegion(wb.xlsx, name = c("Test", "AAA", "BBB", "Test"), header = TRUE, drop = list(c(1, 2), c(2, 3)))
  expect_equal(
    list(
      Test = checkDf[, c(3, 4)],
      AAA = data.frame(A = AAA[, 1], stringsAsFactors = F),
      BBB = data.frame(F = BBB[, 3], stringsAsFactors = F),
      Test = checkDf[, c(1, 4)]
    ),
    res
  )
  targetNoForceSubset <- data.frame(
    BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
    DDD = as.POSIXct(c("1984-01-11 12:00:00", NA, NA, NA, NA)),
    stringsAsFactors = FALSE
  )
  # Check that conversion performs ok (without forcing conversion, keeping columns BBB and DDD; *.xls)
  res <- readNamedRegion(
    wb.xls,
    name = "Conversion",
    header = TRUE,
    colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
    forceConversion = FALSE,
    dateTimeFormat = "%d.%m.%Y %H:%M:%S",
    keep = c("BBB", "DDD")
  )
  expect_equal(targetNoForceSubset, res)

  # Check that conversion performs ok (without forcing conversion, keeping columns BBB and DDD; *.xlsx)
  res <- readNamedRegion(
    wb.xlsx,
    name = "Conversion",
    header = TRUE,
    colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
    forceConversion = FALSE,
    dateTimeFormat = "%d.%m.%Y %H:%M:%S",
    keep = c("BBB", "DDD")
  )
  expect_equal(targetNoForceSubset, res)
  # Check that conversion performs ok (without forcing conversion, dropping columns AAA and CCC; *.xls)
  res <- readNamedRegion(
    wb.xls,
    name = "Conversion",
    header = TRUE,
    colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
    forceConversion = FALSE,
    dateTimeFormat = "%d.%m.%Y %H:%M:%S",
    drop = c("AAA", "CCC")
  )
  expect_equal(targetNoForceSubset, res)

  # Check that conversion performs ok (without forcing conversion, dropping columns AAA and CCC; *.xlsx)
  res <- readNamedRegion(
    wb.xlsx,
    name = "Conversion",
    header = TRUE,
    colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
    forceConversion = FALSE,
    dateTimeFormat = "%d.%m.%Y %H:%M:%S",
    drop = c("AAA", "CCC")
  )
  expect_equal(targetNoForceSubset, res)
  # Check that conversion performs ok (without forcing conversion, keeping columns 2 and 4; *.xls)
  res <- readNamedRegion(
    wb.xls,
    name = "Conversion",
    header = TRUE,
    colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
    forceConversion = FALSE,
    dateTimeFormat = "%d.%m.%Y %H:%M:%S",
    keep = c(2, 4)
  )
  expect_equal(targetNoForceSubset, res)

  # Check that conversion performs ok (without forcing conversion, keeping columns 2 and 4; *.xlsx)
  res <- readNamedRegion(
    wb.xlsx,
    name = "Conversion",
    header = TRUE,
    colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
    forceConversion = FALSE,
    dateTimeFormat = "%d.%m.%Y %H:%M:%S",
    keep = c(2, 4)
  )
  expect_equal(targetNoForceSubset, res)
  # Check that conversion performs ok (without forcing conversion, dropping columns 1 and 3; *.xls)
  res <- suppressWarnings(readNamedRegion(
    wb.xls,
    name = "Conversion",
    header = TRUE,
    colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
    forceConversion = FALSE,
    dateTimeFormat = "%d.%m.%Y %H:%M:%S",
    drop = c(1, 3)
  ))
  expect_equal(targetNoForceSubset, res)

  # Check that conversion performs ok (without forcing conversion, dropping columns 1 and 3; *.xlsx)
  res <- suppressWarnings(readNamedRegion(
    wb.xlsx,
    name = "Conversion",
    header = TRUE,
    colTypes = c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME),
    forceConversion = FALSE,
    dateTimeFormat = "%d.%m.%Y %H:%M:%S",
    drop = c(1, 3)
  ))
  expect_equal(targetNoForceSubset, res)
  # Check that simplification works as expected (*.xls)
  res <- readNamedRegion(wb.xls, name = "Simplify1", header = TRUE, simplify = TRUE)
  expect_equal(1:10, res)
  res <- readNamedRegion(wb.xls, name = "Simplify2", header = TRUE, simplify = TRUE)
  expect_equal(1:4, res)
  res <- readNamedRegion(wb.xls, name = "Simplify3", header = TRUE, simplify = TRUE)
  expect_equal(c(TRUE, FALSE, FALSE, TRUE), res)
  res <- readNamedRegion(wb.xls, name = "Simplify4", header = TRUE, simplify = TRUE)
  expect_equal(c("one", "two", "three", "four", "five"), res)

  # Check that simplification works as expected (*.xlsx)
  res <- readNamedRegion(wb.xlsx, name = "Simplify1", header = TRUE, simplify = TRUE)
  expect_equal(1:10, res)
  res <- readNamedRegion(wb.xlsx, name = "Simplify2", header = TRUE, simplify = TRUE)
  expect_equal(1:4, res)
  res <- readNamedRegion(wb.xlsx, name = "Simplify3", header = TRUE, simplify = TRUE)
  expect_equal(c(TRUE, FALSE, FALSE, TRUE), res)
  res <- readNamedRegion(wb.xlsx, name = "Simplify4", header = TRUE, simplify = TRUE)
  expect_equal(c("one", "two", "three", "four", "five"), res)
  # Cached value tests: Create workbook
  wb.xls <- loadWorkbook(test_path("resources/testCachedValues.xls"), create = FALSE)
  wb.xlsx <- loadWorkbook(test_path("resources/testCachedValues.xlsx"), create = FALSE)
  # "AllLocal" contains no formulae
  ref.xls.uncached <- readNamedRegion(wb.xls, "AllLocal", useCachedValues = FALSE)
  ref.xls.cached <- readNamedRegion(wb.xls, "AllLocal", useCachedValues = TRUE)
  # cached and uncached results should be identical
  expect_equal(ref.xls.cached, ref.xls.uncached)

  ref.xlsx.uncached <- readNamedRegion(wb.xlsx, "AllLocal", useCachedValues = FALSE)
  ref.xlsx.cached <- readNamedRegion(wb.xlsx, "AllLocal", useCachedValues = TRUE)
  expect_equal(ref.xlsx.cached, ref.xlsx.uncached)

  # XLS and XLSX results should be identical
  expect_equal(ref.xls.uncached, ref.xlsx.uncached)
  # the other three named regions reference external worksheets and can't be read
  # with useCachedValues=FALSE
  onErrorCell(wb.xls, XLC$ERROR.STOP)
  expect_error(readNamedRegion(wb.xls, "HeaderRemote", useCachedValues = FALSE))
  expect_error(readNamedRegion(wb.xls, "BodyRemote", useCachedValues = FALSE))
  expect_error(readNamedRegion(wb.xls, "AllRemote", useCachedValues = FALSE))

  onErrorCell(wb.xlsx, XLC$ERROR.STOP)
  expect_error(readNamedRegion(wb.xlsx, "HeaderRemote", useCachedValues = FALSE))
  expect_error(readNamedRegion(wb.xlsx, "BodyRemote", useCachedValues = FALSE))
  expect_error(readNamedRegion(wb.xlsx, "AllRemote", useCachedValues = FALSE))
  res <- readNamedRegion(wb.xls, "HeadersRemote", useCachedValues = TRUE)
  expect_equal(res, ref.xls.uncached)
  res <- readNamedRegion(wb.xls, "BodyRemote", useCachedValues = TRUE)
  expect_equal(res, ref.xls.uncached)
  res <- readNamedRegion(wb.xls, "BothRemote", useCachedValues = TRUE)
  expect_equal(res, ref.xls.uncached)
  res <- readNamedRegion(wb.xlsx, "HeadersRemote", useCachedValues = TRUE)
  expect_equal(res, ref.xls.uncached)
  res <- readNamedRegion(wb.xlsx, "BodyRemote", useCachedValues = TRUE)
  expect_equal(res, ref.xls.uncached)
  res <- readNamedRegion(wb.xlsx, "BothRemote", useCachedValues = TRUE)
  expect_equal(res, ref.xls.uncached)
  # Check that dimensionality is not dropped when reading in a named region with rownames = x
  # (see github issue #49)
  expected <- data.frame(B = 1:5, row.names = letters[1:5])
  res <- readNamedRegionFromFile(test_path("resources/testBug49.xlsx"), name = "test", rownames = "A")
  expect_equal(res, expected)
  options(XLConnect.setCustomAttributes = TRUE)

  # Check that named region can be read within a worksheet scope (see github issue #37)

  name <- "Bla"
  wb37xlsx <- loadWorkbook(test_path("resources/test37.xlsx"), create = FALSE)
  read1 <- readNamedRegion(wb37xlsx, name, worksheetScope = "Sheet1")
  expect_equal("Sheet1", attr(read1, "worksheetScope", exact = TRUE))
  read2 <- readNamedRegion(wb37xlsx, name, worksheetScope = "Sheet2")
  expect_equal("Sheet2", attr(read2, "worksheetScope", exact = TRUE))
  expect_equal(c("bla1"), colnames(read1))
  expect_equal(c("bla2"), colnames(read2))
  readBoth <- readNamedRegion(wb37xlsx, name, worksheetScope = c("Sheet1", "Sheet2"))
  expect_equal("bla1", colnames(readBoth[[1]]))
  expect_equal("bla2", colnames(readBoth[[2]]))

  expect_error(readNamedRegion(wb37xlsx, name, worksheetScope = "Sheet3"))
})
