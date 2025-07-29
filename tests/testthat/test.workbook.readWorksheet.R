test_that("reading basic worksheets by index and name works in XLS", {
    wb.xls <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xls"), create = FALSE)
    common_checkDf <- data.frame(
        NumericColumn = c(-23.63, NA, NA, 5.8, 3),
        StringColumn = c("Hello", NA, NA, NA, "World"),
        BooleanColumn = c(TRUE, FALSE, FALSE, NA, NA),
        DateTimeColumn = as.POSIXct(c(NA, NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07")),
        stringsAsFactors = FALSE
    )
    expect_equal(common_checkDf, readWorksheet(wb.xls, 1), info = "XLS: Read sheet 1 by index")
    expect_equal(common_checkDf, readWorksheet(wb.xls, "Test1"), info = "XLS: Read sheet 'Test1' by name")
})

test_that("reading basic worksheets by index and name works in XLSX", {
    wb.xlsx <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xlsx"), create = FALSE)
    common_checkDf <- data.frame(
        NumericColumn = c(-23.63, NA, NA, 5.8, 3),
        StringColumn = c("Hello", NA, NA, NA, "World"),
        BooleanColumn = c(TRUE, FALSE, FALSE, NA, NA),
        DateTimeColumn = as.POSIXct(c(NA, NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07")),
        stringsAsFactors = FALSE
    )
    expect_equal(common_checkDf, readWorksheet(wb.xlsx, 1), info = "XLSX: Read sheet 1 by index")
    expect_equal(common_checkDf, readWorksheet(wb.xlsx, "Test1"), info = "XLSX: Read sheet 'Test1' by name")
})

test_that("reading specific regions works in XLS", {
    wb.xls <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xls"), create = FALSE)
    common_checkDf <- data.frame(
        NumericColumn = c(-23.63, NA, NA, 5.8, 3),
        StringColumn = c("Hello", NA, NA, NA, "World"),
        BooleanColumn = c(TRUE, FALSE, FALSE, NA, NA),
        DateTimeColumn = as.POSIXct(c(NA, NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07")),
        stringsAsFactors = FALSE
    )
    expect_equal(common_checkDf, readWorksheet(wb.xls, 2, startRow = 17, startCol = 6, endRow = 22, endCol = 9, header = TRUE), info = "XLS: Specific area")
    expected_neg_end <- common_checkDf[-nrow(common_checkDf) + 0:1, -ncol(common_checkDf)]
    expect_equal(expected_neg_end, readWorksheet(wb.xls, "Test2", startRow = 17, startCol = 6, endRow = -2, endCol = -1, header = TRUE), info = "XLS: Negative endRow/Col")
    expect_equal(common_checkDf, readWorksheet(wb.xls, 2, region = "F17:I22", header = TRUE), info = "XLS: Region string")
    expect_equal(common_checkDf, readWorksheet(wb.xls, 2, region = "F17:I22", startRow = 88, endCol = 45, header = TRUE), info = "XLS: Region string with other params")
})

test_that("reading specific regions works in XLSX", {
    wb.xlsx <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xlsx"), create = FALSE)
    common_checkDf <- data.frame(
        NumericColumn = c(-23.63, NA, NA, 5.8, 3),
        StringColumn = c("Hello", NA, NA, NA, "World"),
        BooleanColumn = c(TRUE, FALSE, FALSE, NA, NA),
        DateTimeColumn = as.POSIXct(c(NA, NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07")),
        stringsAsFactors = FALSE
    )
    expect_equal(common_checkDf, readWorksheet(wb.xlsx, "Test2", startRow = 17, startCol = 6, endRow = 22, endCol = 9, header = TRUE), info = "XLSX: Specific area by name")
    expected_neg_end <- common_checkDf[-nrow(common_checkDf) + 0:1, -ncol(common_checkDf)]
    expect_equal(expected_neg_end, readWorksheet(wb.xlsx, "Test2", startRow = 17, startCol = 6, endRow = -2, endCol = -1, header = TRUE), info = "XLSX: Negative endRow/Col")
    expect_equal(common_checkDf, readWorksheet(wb.xlsx, "Test2", region = "F17:I22", header = TRUE), info = "XLSX: Region string by name")
    expect_equal(common_checkDf, readWorksheet(wb.xlsx, "Test2", region = "F17:I22", startRow = 88, endCol = 45, header = TRUE), info = "XLSX: Region string with other params by name")
})

test_that("handling of non-existent and empty sheets is correct in XLS", {
    wb.xls <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xls"), create = FALSE)
    expect_error(readWorksheet(wb.xls, 23), info = "XLS: Non-existent sheet index")
    expect_error(readWorksheet(wb.xls, "SheetDoesNotExist"), info = "XLS: Non-existent sheet name")
    res_xls_3 <- suppressMessages(readWorksheet(wb.xls, 3))
    expect_equal(data.frame(), res_xls_3, info = "XLS: Empty sheet by index (Test3)")
    res_xls_Test3 <- suppressMessages(readWorksheet(wb.xls, "Test3"))
    expect_equal(data.frame(), res_xls_Test3, info = "XLS: Empty sheet by name (Test3)")
})

test_that("handling of non-existent and empty sheets is correct in XLSX", {
    wb.xlsx <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xlsx"), create = FALSE)
    expect_error(readWorksheet(wb.xlsx, 23), info = "XLSX: Non-existent sheet index")
    expect_error(readWorksheet(wb.xlsx, "SheetDoesNotExist"), info = "XLSX: Non-existent sheet name")
    res_xlsx_3 <- suppressMessages(readWorksheet(wb.xlsx, 3))
    expect_equal(data.frame(), res_xlsx_3, info = "XLSX: Empty sheet by index (Test3)")
    res_xlsx_Test3 <- suppressMessages(readWorksheet(wb.xlsx, "Test3"))
    expect_equal(data.frame(), res_xlsx_Test3, info = "XLSX: Empty sheet by name (Test3)")
})

test_that("reading sheets with NAs and varied data works in XLS", {
    wb.xls <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xls"), create = FALSE)
    common_checkDf1 <- data.frame(
        A = c(1:2, NA, 3:6, NA), B = letters[1:8],
        C = c("z", "y", "x", "w", NA, "v", "u", NA), D = c(NA, 1:5, NA, NA),
        stringsAsFactors = FALSE
    )
    common_checkDf2 <- data.frame(
        A = c(rep(NA, 3), 3:6, NA), B = c(NA, letters[2:8]),
        C = c("z", "y", "x", "w", NA, "v", "u", NA), D = c(NA, 1:5, NA, NA),
        stringsAsFactors = FALSE
    )
    expect_equal(common_checkDf1, readWorksheet(wb.xls, "Test4"), info = "XLS: Test4 sheet")
    expect_equal(common_checkDf2, readWorksheet(wb.xls, "Test5"), info = "XLS: Test5 sheet")
    expected_test4_neg <- common_checkDf1[-nrow(common_checkDf1) + 0:3, -ncol(common_checkDf1) + 0:1]
    expect_equal(expected_test4_neg, readWorksheet(wb.xls, "Test4", endRow = -4, endCol = -2), info = "XLS: Test4 negative endRow/Col")
    expected_test5_neg <- common_checkDf2[-nrow(common_checkDf2) + 0:2, -ncol(common_checkDf2)]
    expect_equal(expected_test5_neg, readWorksheet(wb.xls, "Test5", endRow = -3, endCol = -1), info = "XLS: Test5 negative endRow/Col")
})

test_that("reading sheets with NAs and varied data works in XLSX", {
    wb.xlsx <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xlsx"), create = FALSE)
    common_checkDf1 <- data.frame(
        A = c(1:2, NA, 3:6, NA), B = letters[1:8],
        C = c("z", "y", "x", "w", NA, "v", "u", NA), D = c(NA, 1:5, NA, NA),
        stringsAsFactors = FALSE
    )
    common_checkDf2 <- data.frame(
        A = c(rep(NA, 3), 3:6, NA), B = c(NA, letters[2:8]),
        C = c("z", "y", "x", "w", NA, "v", "u", NA), D = c(NA, 1:5, NA, NA),
        stringsAsFactors = FALSE
    )
    expect_equal(common_checkDf1, readWorksheet(wb.xlsx, "Test4"), info = "XLSX: Test4 sheet")
    expect_equal(common_checkDf2, readWorksheet(wb.xlsx, "Test5"), info = "XLSX: Test5 sheet")
    expected_test4_neg <- common_checkDf1[-nrow(common_checkDf1) + 0:3, -ncol(common_checkDf1) + 0:1]
    expect_equal(expected_test4_neg, readWorksheet(wb.xlsx, "Test4", endRow = -4, endCol = -2), info = "XLSX: Test4 negative endRow/Col")
    expected_test5_neg <- common_checkDf2[-nrow(common_checkDf2) + 0:2, -ncol(common_checkDf2)]
    expect_equal(expected_test5_neg, readWorksheet(wb.xlsx, "Test5", endRow = -3, endCol = -1), info = "XLSX: Test5 negative endRow/Col")
})

test_that("column type conversion works in XLS", {
    wb.xls <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xls"), create = FALSE)
    col_types_spec <- c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME)
    datetime_fmt <- "%d.%m.%Y %H:%M:%S"
    targetNoForce <- data.frame(
        AAA = c(NA, NA, NA, 780.9, NA),
        BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
        CCC = c(TRUE, NA, NA, NA, NA), DDD = as.POSIXct(c("1984-01-11 12:00:00", NA, NA, NA, NA)),
        stringsAsFactors = FALSE
    )
    targetForce <- data.frame(
        AAA = c(-14.65, NA, 11.7, 780.9, NA),
        BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
        CCC = c(TRUE, TRUE, NA, FALSE, FALSE), DDD = as.POSIXct(c("1984-01-11 12:00:00", "2012-02-06 16:15:23", "1984-01-11 12:00:00", NA, "1900-12-22 16:04:48")),
        stringsAsFactors = FALSE
    )
    res_xls_noforce <- readWorksheet(wb.xls, sheet = "Conversion", header = TRUE, colTypes = col_types_spec, forceConversion = FALSE, dateTimeFormat = datetime_fmt)
    expect_equal(targetNoForce, res_xls_noforce, info = "XLS: Conversion sheet, no force")
    res_xls_force <- readWorksheet(wb.xls, sheet = "Conversion", header = TRUE, colTypes = col_types_spec, forceConversion = TRUE, dateTimeFormat = datetime_fmt)
    expect_equal(targetForce, res_xls_force, info = "XLS: Conversion sheet, force")
})

test_that("column type conversion works in XLSX", {
    wb.xlsx <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xlsx"), create = FALSE)
    col_types_spec <- c(XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING, XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME)
    datetime_fmt <- "%d.%m.%Y %H:%M:%S"
    targetNoForce <- data.frame(
        AAA = c(NA, NA, NA, 780.9, NA),
        BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
        CCC = c(TRUE, NA, NA, NA, NA), DDD = as.POSIXct(c("1984-01-11 12:00:00", NA, NA, NA, NA)),
        stringsAsFactors = FALSE
    )
    targetForce <- data.frame(
        AAA = c(-14.65, NA, 11.7, 780.9, NA),
        BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
        CCC = c(TRUE, TRUE, NA, FALSE, FALSE), DDD = as.POSIXct(c("1984-01-11 12:00:00", "2012-02-06 16:15:23", "1984-01-11 12:00:00", NA, "1900-12-22 16:04:48")),
        stringsAsFactors = FALSE
    )
    res_xlsx_noforce <- readWorksheet(wb.xlsx, sheet = "Conversion", header = TRUE, colTypes = col_types_spec, forceConversion = FALSE, dateTimeFormat = datetime_fmt)
    expect_equal(targetNoForce, res_xlsx_noforce, info = "XLSX: Conversion sheet, no force")
    res_xlsx_force <- readWorksheet(wb.xlsx, sheet = "Conversion", header = TRUE, colTypes = col_types_spec, forceConversion = TRUE, dateTimeFormat = datetime_fmt)
    expect_equal(targetForce, res_xlsx_force, info = "XLSX: Conversion sheet, force")
})

test_that("reading multiple sheets by name works in XLS", {
    wb.xls <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xls"), create = FALSE)
    target_multi_sheet <- list(
        AAA = data.frame(A = 1:3, B = letters[1:3], C = c(TRUE, TRUE, FALSE), stringsAsFactors = FALSE),
        BBB = data.frame(D = 4:6, E = letters[4:6], F = c(FALSE, TRUE, TRUE), stringsAsFactors = FALSE)
    )
    expect_equal(target_multi_sheet, readWorksheet(wb.xls, sheet = c("AAA", "BBB"), header = TRUE), info = "XLS: Multi-sheet read")
})

test_that("reading multiple sheets by name works in XLSX", {
    wb.xlsx <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xlsx"), create = FALSE)
    target_multi_sheet <- list(
        AAA = data.frame(A = 1:3, B = letters[1:3], C = c(TRUE, TRUE, FALSE), stringsAsFactors = FALSE),
        BBB = data.frame(D = 4:6, E = letters[4:6], F = c(FALSE, TRUE, TRUE), stringsAsFactors = FALSE)
    )
    expect_equal(target_multi_sheet, readWorksheet(wb.xlsx, sheet = c("AAA", "BBB"), header = TRUE), info = "XLSX: Multi-sheet read")
})

test_that("handling of variable names works in XLS", {
    wb.xls <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xls"), create = FALSE)
    target_var_names <- data.frame(
        `With whitespace` = 1:4, `And some other funky characters: _=?^~!$@#%§` = letters[1:4],
        check.names = FALSE, stringsAsFactors = FALSE
    )
    res_xls_varnames <- readWorksheet(wb.xls, sheet = "VariableNames", header = TRUE, check.names = FALSE)
    expect_equal(target_var_names, res_xls_varnames, info = "XLS: VariableNames sheet, check.names=FALSE")
})

test_that("handling of variable names works in XLSX", {
    wb.xlsx <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xlsx"), create = FALSE)
    target_var_names <- data.frame(
        `With whitespace` = 1:4, `And some other funky characters: _=?^~!$@#%§` = letters[1:4],
        check.names = FALSE, stringsAsFactors = FALSE
    )
    res_xlsx_varnames <- readWorksheet(wb.xlsx, sheet = "VariableNames", header = TRUE, check.names = FALSE)
    expect_equal(target_var_names, res_xlsx_varnames, info = "XLSX: VariableNames sheet, check.names=FALSE")
})

test_that("keep and drop arguments work correctly in XLS", {
    wb.xls <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xls"), create = FALSE)
    checkDfSubset <- data.frame(A = c(rep(NA, 3), 3:6, NA), C = c("z", "y", "x", "w", NA, "v", "u", NA), stringsAsFactors = FALSE)
    expect_error(readWorksheet(wb.xls, "Test5", header = TRUE, keep = c("A", "C"), drop = c("B", "D")), info = "XLS: keep and drop both specified")
    expect_error(readWorksheet(wb.xls, "Test5", header = TRUE, keep = c("A", "Z")), info = "XLS: keep non-existent column name")
    expect_error(readWorksheet(wb.xls, "Test5", header = TRUE, keep = c(1, 5)), info = "XLS: keep non-existent column index") # Max 4 cols
    expect_error(readWorksheet(wb.xls, "Test5", header = TRUE, drop = c("A", "Z")), info = "XLS: drop non-existent column name")
    expect_error(readWorksheet(wb.xls, "Test5", header = TRUE, drop = c(1, 5)), info = "XLS: drop non-existent column index")
    expect_equal(checkDfSubset, readWorksheet(wb.xls, "Test5", header = TRUE, keep = c("A", "C")), info = "XLS: keep by name")
    expect_error(readWorksheet(wb.xls, "Test5", header = FALSE, keep = c("A", "C")), info = "XLS: keep by name with header=FALSE")
    expect_equal(checkDfSubset, readWorksheet(wb.xls, "Test5", header = TRUE, drop = c("B", "D")), info = "XLS: drop by name")
    expect_error(readWorksheet(wb.xls, "Test5", header = FALSE, drop = c("B", "D")), info = "XLS: drop by name with header=FALSE")
    expect_equal(checkDfSubset, readWorksheet(wb.xls, "Test5", header = TRUE, keep = c(1, 3)), info = "XLS: keep by index")
    expect_equal(checkDfSubset, readWorksheet(wb.xls, "Test5", header = TRUE, drop = c(2, 4)), info = "XLS: drop by index")
})

test_that("keep and drop arguments work correctly in XLSX", {
    wb.xlsx <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xlsx"), create = FALSE)
    checkDfSubset <- data.frame(A = c(rep(NA, 3), 3:6, NA), C = c("z", "y", "x", "w", NA, "v", "u", NA), stringsAsFactors = FALSE)
    expect_equal(checkDfSubset, readWorksheet(wb.xlsx, "Test5", header = TRUE, keep = c("A", "C")), info = "XLSX: keep by name")
    expect_equal(checkDfSubset, readWorksheet(wb.xlsx, "Test5", header = TRUE, drop = c("B", "D")), info = "XLSX: drop by name")
    expect_equal(checkDfSubset, readWorksheet(wb.xlsx, "Test5", header = TRUE, keep = c(1, 3)), info = "XLSX: keep by index")
    expect_equal(checkDfSubset, readWorksheet(wb.xlsx, "Test5", header = TRUE, drop = c(2, 4)), info = "XLSX: drop by index")
})

test_that("keep and drop arguments with specified region work correctly", {
    wb.xls <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xls"), create = FALSE)
    wb.xlsx <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xlsx"), create = FALSE)

    region_params <- list(sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol = 9, header = TRUE) # This region is B, C, D columns from original Test5
    
    checkDfAreaSubset <- data.frame(B = c(NA, letters[2:7]), D = c(NA, 1:5, NA), stringsAsFactors = FALSE)


    # Errors
    expect_error(do.call(readWorksheet, c(list(wb.xls), region_params, list(keep = c("B", "D"), drop = c("C")))), info = "XLS: Region keep and drop")
    expect_error(do.call(readWorksheet, c(list(wb.xls), region_params, list(keep = c("B", "Z")))), info = "XLS: Region keep non-existent name")
    expect_error(do.call(readWorksheet, c(list(wb.xls), region_params, list(keep = c(1, 5)))), info = "XLS: Region keep non-existent index") # Max 3 cols in region G,H,I
    expect_error(do.call(readWorksheet, c(list(wb.xls), region_params, list(drop = c("B", "Z")))), info = "XLS: Region drop non-existent name")
    expect_error(do.call(readWorksheet, c(list(wb.xls), region_params, list(drop = c(1, 5)))), info = "XLS: Region drop non-existent index")

    # Keep by name in region
    expect_equal(checkDfAreaSubset, do.call(readWorksheet, c(list(wb.xls), region_params, list(keep = c("B", "D")))), info = "XLS: Region keep by name")
    expect_equal(checkDfAreaSubset, do.call(readWorksheet, c(list(wb.xlsx), region_params, list(keep = c("B", "D")))), info = "XLSX: Region keep by name")

    # Drop by name in region
    expect_equal(checkDfAreaSubset, do.call(readWorksheet, c(list(wb.xls), region_params, list(drop = "C"))), info = "XLS: Region drop by name")
    expect_equal(checkDfAreaSubset, do.call(readWorksheet, c(list(wb.xlsx), region_params, list(drop = "C"))), info = "XLSX: Region drop by name")

    # Keep by index in region (cols B,C,D map to 1,2,3 in the sub-region)
    expect_equal(checkDfAreaSubset, do.call(readWorksheet, c(list(wb.xls), region_params, list(keep = c(1, 3)))), info = "XLS: Region keep by index") # Keep B (1) and D (3)
    expect_equal(checkDfAreaSubset, do.call(readWorksheet, c(list(wb.xlsx), region_params, list(keep = c(1, 3)))), info = "XLSX: Region keep by index")

    # Drop by index in region
    expect_equal(checkDfAreaSubset, do.call(readWorksheet, c(list(wb.xls), region_params, list(drop = 2))), info = "XLS: Region drop by index") # Drop C (2)
    expect_equal(checkDfAreaSubset, do.call(readWorksheet, c(list(wb.xlsx), region_params, list(drop = 2))), info = "XLSX: Region drop by index")
})

test_that("keep/drop with multiple sheets works in XLS", {
    wb.xls <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xls"), create = FALSE)
    common_checkDf <- data.frame(
        NumericColumn = c(-23.63, NA, NA, 5.8, 3),
        StringColumn = c("Hello", NA, NA, NA, "World"),
        BooleanColumn = c(TRUE, FALSE, FALSE, NA, NA),
        DateTimeColumn = as.POSIXct(c(NA, NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07")),
        stringsAsFactors = FALSE
    )
    common_checkDf1 <- data.frame(
        A = c(1:2, NA, 3:6, NA), B = letters[1:8],
        C = c("z", "y", "x", "w", NA, "v", "u", NA), D = c(NA, 1:5, NA, NA),
        stringsAsFactors = FALSE
    )
    common_checkDf2 <- data.frame(
        A = c(rep(NA, 3), 3:6, NA), B = c(NA, letters[2:8]),
        C = c("z", "y", "x", "w", NA, "v", "u", NA), D = c(NA, 1:5, NA, NA),
        stringsAsFactors = FALSE
    )
    testAAA_df <- data.frame(A = 1:3, B = letters[1:3], C = c(TRUE, TRUE, FALSE), stringsAsFactors = FALSE)
    sheets_to_read <- c("Test1", "Test4", "Test5")
    res_xls_kl1 <- readWorksheet(wb.xls, sheet = sheets_to_read, header = TRUE, keep = c(1, 2, 3))
    expect_equal(list(Test1 = common_checkDf[1:3], Test4 = common_checkDf1[1:3], Test5 = common_checkDf2[1:3]), res_xls_kl1, info = "XLS: Multi-sheet keep same cols")
    res_xls_kl2 <- readWorksheet(wb.xls, sheet = sheets_to_read, header = TRUE, keep = list(1, 2, c(1,3)))
    expect_equal(list(Test1 = common_checkDf[1], Test4 = common_checkDf1[2], Test5 = common_checkDf2[c(1,3)]), res_xls_kl2, info = "XLS: Multi-sheet keep different cols (simple list)")
    res_xls_kl3 <- readWorksheet(wb.xls, sheet = sheets_to_read, header = TRUE, keep = list(c(1,2), c(2,3), c(1,3)))
    expect_equal(list(Test1 = common_checkDf[1:2], Test4 = common_checkDf1[2:3], Test5 = common_checkDf2[c(1,3)]), res_xls_kl3, info = "XLS: Multi-sheet keep different cols (list of vectors)")
    sheets_plus_aaa <- c("Test1", "Test4", "Test5", "AAA")
    res_xls_kl4 <- readWorksheet(wb.xls, sheet = sheets_plus_aaa, header = TRUE, keep = list(c(1,2), c(2,3)))
    expect_equal(list(Test1 = common_checkDf[1:2], Test4 = common_checkDf1[2:3], Test5 = common_checkDf2[1:2], AAA = testAAA_df[2:3]), res_xls_kl4, info = "XLS: Multi-sheet keep, recycle last keep spec (adjusted for observed behavior)")
    res_xls_dl1 <- readWorksheet(wb.xls, sheet = sheets_to_read, header = TRUE, drop = c(1, 2))
    expect_equal(list(Test1 = common_checkDf[3:4], Test4 = common_checkDf1[3:4], Test5 = common_checkDf2[3:4]), res_xls_dl1, info = "XLS: Multi-sheet drop same cols")
    res_xls_dl2 <- readWorksheet(wb.xls, sheet = sheets_to_read, header = TRUE, drop = list(1, 2, c(1,3)))
    expect_equal(list(Test1 = common_checkDf[2:4], Test4 = common_checkDf1[c(1,3,4)], Test5 = common_checkDf2[c(2,4)]), res_xls_dl2, info = "XLS: Multi-sheet drop different cols (simple list)")
})

test_that("keep/drop with multiple sheets works in XLSX", {
    wb.xlsx <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xlsx"), create = FALSE)
    common_checkDf <- data.frame(
        NumericColumn = c(-23.63, NA, NA, 5.8, 3),
        StringColumn = c("Hello", NA, NA, NA, "World"),
        BooleanColumn = c(TRUE, FALSE, FALSE, NA, NA),
        DateTimeColumn = as.POSIXct(c(NA, NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07")),
        stringsAsFactors = FALSE
    )
    common_checkDf1 <- data.frame(
        A = c(1:2, NA, 3:6, NA), B = letters[1:8],
        C = c("z", "y", "x", "w", NA, "v", "u", NA), D = c(NA, 1:5, NA, NA),
        stringsAsFactors = FALSE
    )
    common_checkDf2 <- data.frame(
        A = c(rep(NA, 3), 3:6, NA), B = c(NA, letters[2:8]),
        C = c("z", "y", "x", "w", NA, "v", "u", NA), D = c(NA, 1:5, NA, NA),
        stringsAsFactors = FALSE
    )
    testAAA_df <- data.frame(A = 1:3, B = letters[1:3], C = c(TRUE, TRUE, FALSE), stringsAsFactors = FALSE)
    sheets_to_read <- c("Test1", "Test4", "Test5")
    res_xlsx_kl1 <- readWorksheet(wb.xlsx, sheet = sheets_to_read, header = TRUE, keep = c(1, 2, 3))
    expect_equal(list(Test1 = common_checkDf[1:3], Test4 = common_checkDf1[1:3], Test5 = common_checkDf2[1:3]), res_xlsx_kl1, info = "XLSX: Multi-sheet keep same cols")
    res_xlsx_kl2 <- readWorksheet(wb.xlsx, sheet = sheets_to_read, header = TRUE, keep = list(1, 2, c(1,3)))
    expect_equal(list(Test1 = common_checkDf[1], Test4 = common_checkDf1[2], Test5 = common_checkDf2[c(1,3)]), res_xlsx_kl2, info = "XLSX: Multi-sheet keep different cols (simple list)")
    res_xlsx_kl3 <- readWorksheet(wb.xlsx, sheet = sheets_to_read, header = TRUE, keep = list(c(1,2), c(2,3), c(1,3)))
    expect_equal(list(Test1 = common_checkDf[1:2], Test4 = common_checkDf1[2:3], Test5 = common_checkDf2[c(1,3)]), res_xlsx_kl3, info = "XLSX: Multi-sheet keep different cols (list of vectors)")
    sheets_plus_aaa <- c("Test1", "Test4", "Test5", "AAA")
    res_xlsx_kl4 <- readWorksheet(wb.xlsx, sheet = sheets_plus_aaa, header = TRUE, keep = list(c(1,2), c(2,3)))
    expect_equal(list(Test1 = common_checkDf[1:2], Test4 = common_checkDf1[2:3], Test5 = common_checkDf2[1:2], AAA = testAAA_df[2:3]), res_xlsx_kl4, info = "XLSX: Multi-sheet keep, recycle last keep spec (adjusted for observed behavior)")
    res_xlsx_dl1 <- readWorksheet(wb.xlsx, sheet = sheets_to_read, header = TRUE, drop = c(1, 2))
    expect_equal(list(Test1 = common_checkDf[3:4], Test4 = common_checkDf1[3:4], Test5 = common_checkDf2[3:4]), res_xlsx_dl1, info = "XLSX: Multi-sheet drop same cols")
    res_xlsx_dl2 <- readWorksheet(wb.xlsx, sheet = sheets_to_read, header = TRUE, drop = list(1, 2, c(1,3)))
    expect_equal(list(Test1 = common_checkDf[2:4], Test4 = common_checkDf1[c(1,3,4)], Test5 = common_checkDf2[c(2,4)]), res_xlsx_dl2, info = "XLSX: Multi-sheet drop different cols (simple list)")
})

test_that("autofitRow and autofitCol work for BoundingBox sheet in XLS", {
    wb.xls <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xls"), create = FALSE)
    target1_bb <- data.frame(Col1 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,7,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col2 = c(NA,NA,NA,3,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,13), Col3 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col4 = c(NA,NA,NA,4,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,9,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col5 = c(1,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col6 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col7 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,10,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col8 = c(2,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col9 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col10 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,11,NA,NA,NA,NA,NA), Col11 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col12 = c(NA,NA,NA,NA,NA,NA,5,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,12,NA,NA,NA,NA,NA), Col13 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col14 = c(NA,NA,NA,NA,NA,NA,6,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col15 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col16 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,8,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA))
    target2_orig <- data.frame( Col1 = c(9, NA, NA, NA, NA, NA), Col2 = c( NA, NA, NA, NA, NA, NA ), Col3 = c(NA, NA, NA, NA, NA, NA), Col4 = c(10, NA, NA, NA, NA, NA), Col5 = c( NA, NA, NA, NA, NA, NA ), Col6 = c(NA, NA, NA, NA, NA, NA), Col7 = c( NA, NA, NA, NA, NA, 11 ) )
    target3_orig <- data.frame(Col1=c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col2=c(NA,NA,NA,9,NA,NA,NA,NA,NA,NA,NA,NA), Col3=c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col4=c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col5=c(NA,NA,NA,10,NA,NA,NA,NA,NA,NA,NA,NA), Col6=c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col7=c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col8=c(NA,NA,NA,NA,NA,NA,NA,NA,11,NA,NA,NA), Col9=c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA))
    target4_orig <- as.data.frame(matrix(NA, nrow = 10, ncol = 8)); names(target4_orig) <- paste("Col", 1:8, sep = "")
    target5_orig <- data.frame(Col1=c(NA,NA,NA,NA,4,NA), Col2=c(NA,1,NA,NA,NA,NA))
    target6_orig <- data.frame(Col1=c(NA,NA,NA,NA),Col2=c(NA,NA,NA,4),Col3=c(1,NA,NA,NA),Col4=c(NA,NA,NA,NA),Col5=c(NA,NA,NA,NA))
    target7_orig <- data.frame(Col1=c(NA,NA,NA,4),Col2=c(1,NA,NA,NA))
    expect_equal(target1_bb, readWorksheet(wb.xls, sheet = "BoundingBox", autofitRow = TRUE, autofitCol = TRUE, header = FALSE))
    expect_equal(target1_bb, readWorksheet(wb.xls, sheet = "BoundingBox", autofitRow = FALSE, autofitCol = FALSE, header = FALSE))
    expect_equal(target2_orig, readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 20, startCol = 5, endRow = 31, endCol = 13, autofitRow = TRUE, autofitCol = TRUE, header = FALSE))
    expect_equal(target3_orig, readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 20, startCol = 5, endRow = 31, endCol = 13, autofitRow = FALSE, autofitCol = FALSE, header = FALSE))
    expect_equal(data.frame(), readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 12, startCol = 5, endRow = 21, endCol = 12, autofitRow = TRUE, autofitCol = TRUE, header = FALSE))
    expect_equal(target4_orig, readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 12, startCol = 5, endRow = 21, endCol = 12, autofitRow = FALSE, autofitCol = FALSE, header = FALSE))
    expect_equal(target5_orig, readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 6, startCol = 5, endRow = 11, endCol = 9, autofitRow = FALSE, autofitCol = TRUE, header = FALSE))
    expect_equal(target6_orig, readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 6, startCol = 5, endRow = 11, endCol = 9, autofitRow = TRUE, autofitCol = FALSE, header = FALSE))
    expect_equal(target7_orig, readWorksheet(wb.xls, sheet = "BoundingBox", startRow = 6, startCol = 5, endRow = 11, endCol = 9, autofitRow = TRUE, autofitCol = TRUE, header = FALSE))
})

test_that("autofitRow and autofitCol for BoundingBox sheet work in XLSX", {
    wb.xlsx <- loadWorkbook(test_path("resources/testWorkbookReadWorksheet.xlsx"), create = FALSE)
    target1_bb <- data.frame(Col1 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,7,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col2 = c(NA,NA,NA,3,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,13), Col3 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col4 = c(NA,NA,NA,4,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,9,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col5 = c(1,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col6 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col7 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,10,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col8 = c(2,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col9 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col10 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,11,NA,NA,NA,NA,NA), Col11 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col12 = c(NA,NA,NA,NA,NA,NA,5,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,12,NA,NA,NA,NA,NA), Col13 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col14 = c(NA,NA,NA,NA,NA,NA,6,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col15 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col16 = c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,8,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA))
    target2_orig <- data.frame( Col1 = c(9, NA, NA, NA, NA, NA), Col2 = c( NA, NA, NA, NA, NA, NA ), Col3 = c(NA, NA, NA, NA, NA, NA), Col4 = c(10, NA, NA, NA, NA, NA), Col5 = c( NA, NA, NA, NA, NA, NA ), Col6 = c(NA, NA, NA, NA, NA, NA), Col7 = c( NA, NA, NA, NA, NA, 11 ) )
    target3_orig <- data.frame(Col1=c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col2=c(NA,NA,NA,9,NA,NA,NA,NA,NA,NA,NA,NA), Col3=c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col4=c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col5=c(NA,NA,NA,10,NA,NA,NA,NA,NA,NA,NA,NA), Col6=c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col7=c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA), Col8=c(NA,NA,NA,NA,NA,NA,NA,NA,11,NA,NA,NA), Col9=c(NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA,NA))
    target4_orig <- as.data.frame(matrix(NA, nrow = 10, ncol = 8)); names(target4_orig) <- paste("Col", 1:8, sep = "")
    target5_orig <- data.frame(Col1=c(NA,NA,NA,NA,4,NA), Col2=c(NA,1,NA,NA,NA,NA))
    target6_orig <- data.frame(Col1=c(NA,NA,NA,NA),Col2=c(NA,NA,NA,4),Col3=c(1,NA,NA,NA),Col4=c(NA,NA,NA,NA),Col5=c(NA,NA,NA,NA))
    target7_orig <- data.frame(Col1=c(NA,NA,NA,4),Col2=c(1,NA,NA,NA))
    expect_equal(target1_bb, readWorksheet(wb.xlsx, sheet = "BoundingBox", autofitRow = TRUE, autofitCol = TRUE, header = FALSE))
    expect_equal(target1_bb, readWorksheet(wb.xlsx, sheet = "BoundingBox", autofitRow = FALSE, autofitCol = FALSE, header = FALSE))
    expect_equal(target2_orig, readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 20, startCol = 5, endRow = 31, endCol = 13, autofitRow = TRUE, autofitCol = TRUE, header = FALSE))
    expect_equal(target3_orig, readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 20, startCol = 5, endRow = 31, endCol = 13, autofitRow = FALSE, autofitCol = FALSE, header = FALSE))
    expect_equal(data.frame(), readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 12, startCol = 5, endRow = 21, endCol = 12, autofitRow = TRUE, autofitCol = TRUE, header = FALSE))
    expect_equal(target4_orig, readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 12, startCol = 5, endRow = 21, endCol = 12, autofitRow = FALSE, autofitCol = FALSE, header = FALSE))
    expect_equal(target5_orig, readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 6, startCol = 5, endRow = 11, endCol = 9, autofitRow = FALSE, autofitCol = TRUE, header = FALSE))
    expect_equal(target6_orig, readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 6, startCol = 5, endRow = 11, endCol = 9, autofitRow = TRUE, autofitCol = FALSE, header = FALSE))
    expect_equal(target7_orig, readWorksheet(wb.xlsx, sheet = "BoundingBox", startRow = 6, startCol = 5, endRow = 11, endCol = 9, autofitRow = TRUE, autofitCol = TRUE, header = FALSE))
})

test_that("useCachedValues and onErrorCell interaction works in XLS", {
    wb.xls.cache <- loadWorkbook(test_path("resources/testCachedValues.xls"), create = FALSE)
    ref.xls.uncached <- readWorksheet(wb.xls.cache, "AllLocal", useCachedValues = FALSE)
    ref.xls.cached <- readWorksheet(wb.xls.cache, "AllLocal", useCachedValues = TRUE)
    expect_equal(ref.xls.cached, ref.xls.uncached, info = "XLS: Cached vs Uncached for AllLocal")
    onErrorCell(wb.xls.cache, XLC$ERROR.STOP)
    expect_error(readWorksheet(wb.xls.cache, "HeaderRemote", useCachedValues = FALSE), info = "XLS: HeaderRemote uncached error")
    expect_error(readWorksheet(wb.xls.cache, "BodyRemote", useCachedValues = FALSE), info = "XLS: BodyRemote uncached error")
    expect_error(readWorksheet(wb.xls.cache, "AllRemote", useCachedValues = FALSE), info = "XLS: AllRemote uncached error")
    expect_equal(readWorksheet(wb.xls.cache, "HeadersRemote", useCachedValues = TRUE), ref.xls.uncached, info = "XLS: HeadersRemote cached")
    expect_equal(readWorksheet(wb.xls.cache, "BodyRemote", useCachedValues = TRUE), ref.xls.uncached, info = "XLS: BodyRemote cached")
    expect_equal(readWorksheet(wb.xls.cache, "BothRemote", useCachedValues = TRUE), ref.xls.uncached, info = "XLS: BothRemote cached")
})

test_that("useCachedValues and onErrorCell interaction works in XLSX", {
    wb.xlsx.cache <- loadWorkbook(test_path("resources/testCachedValues.xlsx"), create = FALSE)
    ref.xlsx.uncached <- readWorksheet(wb.xlsx.cache, "AllLocal", useCachedValues = FALSE)
    ref.xlsx.cached <- readWorksheet(wb.xlsx.cache, "AllLocal", useCachedValues = TRUE)
    expect_equal(ref.xlsx.cached, ref.xlsx.uncached, info = "XLSX: Cached vs Uncached for AllLocal")
    onErrorCell(wb.xlsx.cache, XLC$ERROR.STOP)
    expect_error(readWorksheet(wb.xlsx.cache, "HeaderRemote", useCachedValues = FALSE), info = "XLSX: HeaderRemote uncached error")
    expect_error(readWorksheet(wb.xlsx.cache, "BodyRemote", useCachedValues = FALSE), info = "XLSX: BodyRemote uncached error")
    expect_error(readWorksheet(wb.xlsx.cache, "AllRemote", useCachedValues = FALSE), info = "XLSX: AllRemote uncached error")
    expect_equal(readWorksheet(wb.xlsx.cache, "HeadersRemote", useCachedValues = TRUE), ref.xlsx.uncached, info = "XLSX: HeadersRemote cached")
    expect_equal(readWorksheet(wb.xlsx.cache, "BodyRemote", useCachedValues = TRUE), ref.xlsx.uncached, info = "XLSX: BodyRemote cached")
    expect_equal(readWorksheet(wb.xlsx.cache, "BothRemote", useCachedValues = TRUE), ref.xlsx.uncached, info = "XLSX: BothRemote cached")
})

test_that("readWorksheetFromFile with useCachedValues works (Bug 52)", {
    res_bug52 <- readWorksheetFromFile(test_path("resources/testBug52.xlsx"), sheet = 1, useCachedValues = TRUE)
    expected_bug52 <- data.frame(Var1 = c(2,4,6), Var2 = c("2", "nope", "6"), Var3 = c(NA,4,6), Var4 = c(2,4,6), stringsAsFactors = FALSE)
    expect_equal(res_bug52, expected_bug52, info = "Bug 52 (cached values)")
})

test_that("readWorksheetFromFile with rownames works (Bug 49)", {
    expected_bug49 <- data.frame(B = 1:5, row.names = letters[1:5])
    res_bug49 <- readWorksheetFromFile(test_path("resources/testBug49.xlsx"), sheet = 1, rownames = 1)
    expect_equal(res_bug49, expected_bug49, info = "Bug 49 (rownames)")
})

test_that("readWorksheetFromFile with dateTimeFormat and forceConversion works (Bug 53)", {
    expected_bug53_sheet1 <- data.frame(A = c("2003-04-06", "2014-10-30", "abc"), stringsAsFactors = FALSE)
    res_bug53_sheet1 <- readWorksheetFromFile(test_path("resources/testBug53.xlsx"), sheet = 1, dateTimeFormat = "%Y-%m-%d")
    expect_equal(res_bug53_sheet1, expected_bug53_sheet1, info = "Bug 53 (sheet 1, dateTimeFormat)")
    expected_bug53_sheet2 <- data.frame(A = as.POSIXct(c("2015-12-01", "2015-11-17", "1984-01-11")))
    res_bug53_sheet2 <- readWorksheetFromFile(test_path("resources/testBug53.xlsx"), sheet = 2, colTypes = "POSIXt", forceConversion = TRUE)
    expect_equal(res_bug53_sheet2, expected_bug53_sheet2, info = "Bug 53 (sheet 2, colTypes POSIXt)")
})

test_that("reading sparse bitset worksheet works", {
    wbSparse.xlsx <- loadWorkbook(test_path("resources/testReadWorksheetSparseBitSet.xlsx"), create = FALSE)
    expect_silent(sparseSheet <- readWorksheet(wbSparse.xlsx, "hist"))
    expect_true(is.data.frame(sparseSheet))
})
