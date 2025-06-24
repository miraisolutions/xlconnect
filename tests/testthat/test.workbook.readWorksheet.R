test_that("test.workbook.readWorksheet", {
    wb.xls <- loadWorkbook("resources/testWorkbookReadWorksheet.xls",
        create = FALSE
    )
    wb.xlsx <- loadWorkbook("resources/testWorkbookReadWorksheet.xlsx",
        create = FALSE
    )
    checkDf <- data.frame(
        NumericColumn = c(
            -23.63, NA, NA, 5.8,
            3
        ), StringColumn = c("Hello", NA, NA, NA, "World"), BooleanColumn = c(
            TRUE,
            FALSE, FALSE, NA, NA
        ), DateTimeColumn = as.POSIXct(c(
            NA,
            NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07"
        )),
        stringsAsFactors = FALSE
    )
    res.index <- readWorksheet(wb.xls, 1)
    expect_equal(checkDf, res.index)
    res.name <- readWorksheet(wb.xls, "Test1")
    expect_equal(checkDf, res.name)
    res.index <- readWorksheet(wb.xlsx, 1)
    expect_equal(checkDf, res.index)
    res.name <- readWorksheet(wb.xlsx, "Test1")
    expect_equal(checkDf, res.name)
    res.index <- readWorksheet(wb.xls, 2,
        startRow = 17, startCol = 6,
        endRow = 22, endCol = 9, header = TRUE
    )
    expect_equal(checkDf, res.index)
    res.name <- readWorksheet(wb.xls, "Test2",
        startRow = 17,
        startCol = 6, endRow = 22, endCol = 9, header = TRUE
    )
    expect_equal(checkDf, res.name)
    res.name <- readWorksheet(wb.xls, "Test2",
        startRow = 17,
        startCol = 6, endRow = -2, endCol = -1, header = TRUE
    )
    expect_equal(
        checkDf[-nrow(checkDf) + 0:1, -ncol(checkDf)],
        res.name
    )
    res.index <- readWorksheet(wb.xlsx, 2,
        startRow = 17, startCol = 6,
        endRow = 22, endCol = 9, header = TRUE
    )
    expect_equal(checkDf, res.index)
    res.name <- readWorksheet(wb.xlsx, "Test2",
        startRow = 17,
        startCol = 6, endRow = 22, endCol = 9, header = TRUE
    )
    expect_equal(checkDf, res.name)
    res.name <- readWorksheet(wb.xlsx, "Test2",
        startRow = 17,
        startCol = 6, endRow = -2, endCol = -1, header = TRUE
    )
    expect_equal(
        checkDf[-nrow(checkDf) + 0:1, -ncol(checkDf)],
        res.name
    )
    res.index <- readWorksheet(wb.xls, 2,
        region = "F17:I22",
        header = TRUE
    )
    expect_equal(checkDf, res.index)
    res.name <- readWorksheet(wb.xls, "Test2",
        region = "F17:I22",
        header = TRUE
    )
    expect_equal(checkDf, res.name)
    res.index <- readWorksheet(wb.xlsx, 2,
        region = "F17:I22",
        header = TRUE
    )
    expect_equal(checkDf, res.index)
    res.name <- readWorksheet(wb.xlsx, "Test2",
        region = "F17:I22",
        header = TRUE
    )
    expect_equal(checkDf, res.name)
    res.index <- readWorksheet(wb.xls, 2,
        region = "F17:I22",
        startRow = 88, endCol = 45, header = TRUE
    )
    expect_equal(checkDf, res.index)
    res.name <- readWorksheet(wb.xls, "Test2",
        region = "F17:I22",
        startRow = 88, endCol = 45, header = TRUE
    )
    expect_equal(checkDf, res.name)
    res.index <- readWorksheet(wb.xlsx, 2,
        region = "F17:I22",
        startRow = 88, endCol = 45, header = TRUE
    )
    expect_equal(checkDf, res.index)
    res.name <- readWorksheet(wb.xlsx, "Test2",
        region = "F17:I22",
        startRow = 88, endCol = 45, header = TRUE
    )
    expect_equal(checkDf, res.name)
    expect_error(readWorksheet(wb.xls, 23))
    expect_error(readWorksheet(wb.xls, "SheetDoesNotExist"))
    expect_error(readWorksheet(wb.xlsx, 23))
    expect_error(readWorksheet(wb.xlsx, "SheetDoesNotExist"))
    expect_equal(data.frame(), readWorksheet(wb.xls, 3))
    expect_equal(data.frame(), readWorksheet(wb.xls, "Test3"))
    expect_equal(data.frame(), readWorksheet(wb.xlsx, 3))
    expect_equal(data.frame(), readWorksheet(wb.xlsx, "Test3"))
    checkDf1 <- data.frame(
        A = c(1:2, NA, 3:6, NA), B = letters[1:8],
        C = c("z", "y", "x", "w", NA, "v", "u", NA), D = c(
            NA,
            1:5, NA, NA
        ), stringsAsFactors = FALSE
    )
    checkDf2 <- data.frame(A = c(rep(NA, 3), 3:6, NA), B = c(
        NA,
        letters[2:8]
    ), C = c(
        "z", "y", "x", "w", NA, "v", "u",
        NA
    ), D = c(NA, 1:5, NA, NA), stringsAsFactors = FALSE)
    res <- readWorksheet(wb.xls, "Test4")
    expect_equal(checkDf1, res)
    res <- readWorksheet(wb.xls, "Test5")
    expect_equal(checkDf2, res)
    res <- readWorksheet(wb.xls, "Test4", endRow = -4, endCol = -2)
    expect_equal(checkDf1[-nrow(checkDf1) + 0:3, -ncol(checkDf) +
        0:1], res)
    res <- readWorksheet(wb.xls, "Test5", endRow = -3, endCol = -1)
    expect_equal(
        checkDf2[-nrow(checkDf2) + 0:2, -ncol(checkDf)],
        res
    )
    res <- readWorksheet(wb.xlsx, "Test4")
    expect_equal(checkDf1, res)
    res <- readWorksheet(wb.xlsx, "Test5")
    expect_equal(checkDf2, res)
    res <- readWorksheet(wb.xlsx, "Test4", endRow = -4, endCol = -2)
    expect_equal(checkDf1[-nrow(checkDf1) + 0:3, -ncol(checkDf) +
        0:1], res)
    res <- readWorksheet(wb.xlsx, "Test5", endRow = -3, endCol = -1)
    expect_equal(
        checkDf2[-nrow(checkDf2) + 0:2, -ncol(checkDf)],
        res
    )
    targetNoForce <- data.frame(
        AAA = c(NA, NA, NA, 780.9, NA),
        BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
        CCC = c(TRUE, NA, NA, NA, NA), DDD = as.POSIXct(c(
            "1984-01-11 12:00:00",
            NA, NA, NA, NA
        )), stringsAsFactors = FALSE
    )
    targetForce <- data.frame(
        AAA = c(
            -14.65, NA, 11.7, 780.9,
            NA
        ), BBB = c("hello", "42.24", "true", NA, "11.01.1984 12:00:00"),
        CCC = c(TRUE, TRUE, NA, FALSE, FALSE), DDD = as.POSIXct(c(
            "1984-01-11 12:00:00",
            "2012-02-06 16:15:23", "1984-01-11 12:00:00", NA,
            "1900-12-22 16:04:48"
        )), stringsAsFactors = FALSE
    )
    res <- readWorksheet(wb.xls,
        sheet = "Conversion", header = TRUE,
        colTypes = c(
            XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING,
            XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME
        ), forceConversion = FALSE,
        dateTimeFormat = "%d.%m.%Y %H:%M:%S"
    )
    expect_equal(targetNoForce, res)
    res <- readWorksheet(wb.xlsx,
        sheet = "Conversion", header = TRUE,
        colTypes = c(
            XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING,
            XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME
        ), forceConversion = FALSE,
        dateTimeFormat = "%d.%m.%Y %H:%M:%S"
    )
    expect_equal(targetNoForce, res)
    res <- readWorksheet(wb.xls,
        sheet = "Conversion", header = TRUE,
        colTypes = c(
            XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING,
            XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME
        ), forceConversion = TRUE,
        dateTimeFormat = "%d.%m.%Y %H:%M:%S"
    )
    expect_equal(targetForce, res)
    res <- readWorksheet(wb.xlsx,
        sheet = "Conversion", header = TRUE,
        colTypes = c(
            XLC$DATA_TYPE.NUMERIC, XLC$DATA_TYPE.STRING,
            XLC$DATA_TYPE.BOOLEAN, XLC$DATA_TYPE.DATETIME
        ), forceConversion = TRUE,
        dateTimeFormat = "%d.%m.%Y %H:%M:%S"
    )
    expect_equal(targetForce, res)
    target <- list(
        AAA = data.frame(
            A = 1:3, B = letters[1:3],
            C = c(TRUE, TRUE, FALSE), stringsAsFactors = FALSE
        ),
        BBB = data.frame(D = 4:6, E = letters[4:6], F = c(
            FALSE,
            TRUE, TRUE
        ), stringsAsFactors = FALSE)
    )
    res <- readWorksheet(wb.xls, sheet = c("AAA", "BBB"), header = TRUE)
    expect_equal(target, res)
    res <- readWorksheet(wb.xlsx, sheet = c("AAA", "BBB"), header = TRUE)
    expect_equal(target, res)
    target <- data.frame(
        `With whitespace` = 1:4, `And some other funky characters: _=?^~!$@#%§` = letters[1:4],
        check.names = FALSE, stringsAsFactors = FALSE
    )
    res <- readWorksheet(wb.xls,
        sheet = "VariableNames", header = TRUE,
        check.names = FALSE
    )
    expect_equal(target, res)
    res <- readWorksheet(wb.xlsx,
        sheet = "VariableNames", header = TRUE,
        check.names = FALSE
    )
    expect_equal(target, res)
    expect_error(readWorksheet(wb.xls,
        sheet = "Test5", header = TRUE,
        keep = c("A", "C"), drop = c("B", "D")
    ))
    expect_error(readWorksheet(wb.xlsx,
        sheet = "Test5", header = TRUE,
        keep = c("A", "C"), drop = c("B", "D")
    ))
    expect_error(readWorksheet(wb.xls,
        sheet = "Test5", header = TRUE,
        keep = c("A", "Z")
    ))
    expect_error(readWorksheet(wb.xlsx,
        sheet = "Test5", header = TRUE,
        keep = c("A", "Z")
    ))
    expect_error(readWorksheet(wb.xls,
        sheet = "Test5", header = TRUE,
        keep = c(1, 5)
    ))
    expect_error(readWorksheet(wb.xlsx,
        sheet = "Test5", header = TRUE,
        keep = c(1, 5)
    ))
    expect_error(readWorksheet(wb.xls,
        sheet = "Test5", header = TRUE,
        drop = c("A", "Z")
    ))
    expect_error(readWorksheet(wb.xlsx,
        sheet = "Test5", header = TRUE,
        drop = c("A", "Z")
    ))
    expect_error(readWorksheet(wb.xls,
        sheet = "Test5", header = TRUE,
        drop = c(1, 5)
    ))
    expect_error(readWorksheet(wb.xlsx,
        sheet = "Test5", header = TRUE,
        drop = c(1, 5)
    ))
    checkDfSubset <- data.frame(A = c(rep(NA, 3), 3:6, NA), C = c(
        "z",
        "y", "x", "w", NA, "v", "u", NA
    ), stringsAsFactors = FALSE)
    res <- readWorksheet(wb.xls, "Test5", header = TRUE, keep = c(
        "A",
        "C"
    ))
    expect_equal(checkDfSubset, res)
    res <- readWorksheet(wb.xlsx, "Test5", header = TRUE, keep = c(
        "A",
        "C"
    ))
    expect_equal(checkDfSubset, res)
    expect_error(readWorksheet(wb.xls, "Test5",
        header = FALSE,
        keep = c("A", "C")
    ))
    expect_error(readWorksheet(wb.xlsx, "Test5",
        header = FALSE,
        keep = c("A", "C")
    ))
    res <- readWorksheet(wb.xls, "Test5", header = TRUE, drop = c(
        "B",
        "D"
    ))
    expect_equal(checkDfSubset, res)
    res <- readWorksheet(wb.xlsx, "Test5", header = TRUE, drop = c(
        "B",
        "D"
    ))
    expect_equal(checkDfSubset, res)
    expect_error(readWorksheet(wb.xls, "Test5",
        header = FALSE,
        drop = c("B", "D")
    ))
    expect_error(readWorksheet(wb.xlsx, "Test5",
        header = FALSE,
        drop = c("B", "D")
    ))
    res <- readWorksheet(wb.xls, "Test5", header = TRUE, keep = c(
        1,
        3
    ))
    expect_equal(checkDfSubset, res)
    res <- readWorksheet(wb.xlsx, "Test5", header = TRUE, keep = c(
        1,
        3
    ))
    expect_equal(checkDfSubset, res)
    res <- readWorksheet(wb.xls, "Test5", header = TRUE, drop = c(
        2,
        4
    ))
    expect_equal(checkDfSubset, res)
    res <- readWorksheet(wb.xlsx, "Test5", header = TRUE, drop = c(
        2,
        4
    ))
    expect_equal(checkDfSubset, res)
    expect_error(readWorksheet(wb.xls,
        sheet = "Test5", startRow = 17,
        startCol = 7, endRow = 24, endCol = 9, header = TRUE,
        keep = c("B", "D"), drop = c("C")
    ))
    expect_error(readWorksheet(wb.xlsx,
        sheet = "Test5", startRow = 17,
        startCol = 7, endRow = 24, endCol = 9, header = TRUE,
        keep = c("B", "D"), drop = c("C")
    ))
    expect_error(readWorksheet(wb.xls,
        sheet = "Test5", startRow = 17,
        startCol = 7, endRow = 24, endCol = 9, header = TRUE,
        keep = c("B", "Z")
    ))
    expect_error(readWorksheet(wb.xlsx,
        sheet = "Test5", startRow = 17,
        startCol = 7, endRow = 24, endCol = 9, header = TRUE,
        keep = c("B", "Z")
    ))
    expect_error(readWorksheet(wb.xls,
        sheet = "Test5", startRow = 17,
        startCol = 7, endRow = 24, endCol = 9, header = TRUE,
        keep = c(1, 5)
    ))
    expect_error(readWorksheet(wb.xlsx,
        sheet = "Test5", startRow = 17,
        startCol = 7, endRow = 24, endCol = 9, header = TRUE,
        keep = c(1, 5)
    ))
    expect_error(readWorksheet(wb.xls,
        sheet = "Test5", startRow = 17,
        startCol = 7, endRow = 24, endCol = 9, header = TRUE,
        drop = c("B", "Z")
    ))
    expect_error(readWorksheet(wb.xlsx,
        sheet = "Test5", startRow = 17,
        startCol = 7, endRow = 24, endCol = 9, header = TRUE,
        drop = c("B", "Z")
    ))
    expect_error(readWorksheet(wb.xls,
        sheet = "Test5", startRow = 17,
        startCol = 7, endRow = 24, endCol = 9, header = TRUE,
        drop = c(1, 5)
    ))
    expect_error(readWorksheet(wb.xlsx,
        sheet = "Test5", startRow = 17,
        startCol = 7, endRow = 24, endCol = 9, header = TRUE,
        drop = c(1, 5)
    ))
    checkDfAreaSubset <- data.frame(
        B = c(NA, letters[2:7]),
        D = c(NA, 1:5, NA), stringsAsFactors = FALSE
    )
    res <- readWorksheet(wb.xls, "Test5",
        startRow = 17, startCol = 7,
        endRow = 24, endCol = 9, header = TRUE, keep = c(
            "B",
            "D"
        )
    )
    expect_equal(checkDfAreaSubset, res)
    res <- readWorksheet(wb.xlsx, "Test5",
        startRow = 17, startCol = 7,
        endRow = 24, endCol = 9, header = TRUE, keep = c(
            "B",
            "D"
        )
    )
    expect_equal(checkDfAreaSubset, res)
    res <- readWorksheet(wb.xls, "Test5",
        startRow = 17, startCol = 7,
        endRow = 24, endCol = 9, header = TRUE, drop = c("C")
    )
    expect_equal(checkDfAreaSubset, res)
    res <- readWorksheet(wb.xlsx, "Test5",
        startRow = 17, startCol = 7,
        endRow = 24, endCol = 9, header = TRUE, drop = c("C")
    )
    expect_equal(checkDfAreaSubset, res)
    res <- readWorksheet(wb.xls, "Test5",
        startRow = 17, startCol = 7,
        endRow = 24, endCol = 9, header = TRUE, keep = c(1, 3)
    )
    expect_equal(checkDfAreaSubset, res)
    res <- readWorksheet(wb.xlsx, "Test5",
        startRow = 17, startCol = 7,
        endRow = 24, endCol = 9, header = TRUE, keep = c(1, 3)
    )
    expect_equal(checkDfAreaSubset, res)
    res <- readWorksheet(wb.xls, "Test5",
        startRow = 17, startCol = 7,
        endRow = 24, endCol = 9, header = TRUE, drop = c(2)
    )
    expect_equal(checkDfAreaSubset, res)
    res <- readWorksheet(wb.xlsx, "Test5",
        startRow = 17, startCol = 7,
        endRow = 24, endCol = 9, header = TRUE, drop = c(2)
    )
    expect_equal(checkDfAreaSubset, res)
    res <- readWorksheet(wb.xls, sheet = c(
        "Test1", "Test4",
        "Test5"
    ), header = TRUE, keep = c(1, 2, 3))
    expect_equal(list(
        Test1 = checkDf[1:3], Test4 = checkDf1[1:3],
        Test5 = checkDf2[1:3]
    ), res)
    res <- readWorksheet(wb.xlsx, sheet = c(
        "Test1", "Test4",
        "Test5"
    ), header = TRUE, keep = c(1, 2, 3))
    expect_equal(list(
        Test1 = checkDf[1:3], Test4 = checkDf1[1:3],
        Test5 = checkDf2[1:3]
    ), res)
    res <- readWorksheet(wb.xls, sheet = c(
        "Test1", "Test4",
        "Test5"
    ), header = TRUE, keep = list(1, 2))
    expect_equal(list(
        Test1 = checkDf[1], Test4 = checkDf1[2],
        Test5 = checkDf2[1]
    ), res)
    res <- readWorksheet(wb.xlsx, sheet = c(
        "Test1", "Test4",
        "Test5"
    ), header = TRUE, keep = list(1, 2))
    expect_equal(list(
        Test1 = checkDf[1], Test4 = checkDf1[2],
        Test5 = checkDf2[1]
    ), res)
    res <- readWorksheet(wb.xls, sheet = c(
        "Test1", "Test4",
        "Test5"
    ), header = TRUE, keep = list(
        c(1, 2), c(2, 3),
        c(1, 3)
    ))
    expect_equal(list(
        Test1 = checkDf[1:2], Test4 = checkDf1[2:3],
        Test5 = checkDf2[c(1, 3)]
    ), res)
    res <- readWorksheet(wb.xlsx, sheet = c(
        "Test1", "Test4",
        "Test5"
    ), header = TRUE, keep = list(
        c(1, 2), c(2, 3),
        c(1, 3)
    ))
    expect_equal(list(
        Test1 = checkDf[1:2], Test4 = checkDf1[2:3],
        Test5 = checkDf2[c(1, 3)]
    ), res)
    testAAA <- data.frame(A = 1:3, B = letters[1:3], C = c(
        TRUE,
        TRUE, FALSE
    ), stringsAsFactors = FALSE)
    res <- readWorksheet(wb.xls, sheet = c(
        "Test1", "Test4",
        "Test5", "AAA"
    ), header = TRUE, keep = list(
        c(1, 2),
        c(2, 3)
    ))
    expect_equal(list(
        Test1 = checkDf[1:2], Test4 = checkDf1[2:3],
        Test5 = checkDf2[1:2], AAA = testAAA[2:3]
    ), res)
    res <- readWorksheet(wb.xlsx, sheet = c(
        "Test1", "Test4",
        "Test5", "AAA"
    ), header = TRUE, keep = list(
        c(1, 2),
        c(2, 3)
    ))
    expect_equal(list(
        Test1 = checkDf[1:2], Test4 = checkDf1[2:3],
        Test5 = checkDf2[1:2], AAA = testAAA[2:3]
    ), res)
    res <- readWorksheet(wb.xls, sheet = c(
        "Test1", "Test4",
        "Test5"
    ), header = TRUE, drop = c(1, 2))
    expect_equal(list(
        Test1 = checkDf[3:4], Test4 = checkDf1[3:4],
        Test5 = checkDf2[3:4]
    ), res)
    res <- readWorksheet(wb.xlsx, sheet = c(
        "Test1", "Test4",
        "Test5"
    ), header = TRUE, drop = c(1, 2))
    expect_equal(list(
        Test1 = checkDf[3:4], Test4 = checkDf1[3:4],
        Test5 = checkDf2[3:4]
    ), res)
    res <- readWorksheet(wb.xls, sheet = c(
        "Test1", "Test4",
        "Test5"
    ), header = TRUE, drop = list(1, 2))
    expect_equal(list(Test1 = checkDf[2:4], Test4 = checkDf1[c(
        1,
        3, 4
    )], Test5 = checkDf2[2:4]), res)
    res <- readWorksheet(wb.xlsx, sheet = c(
        "Test1", "Test4",
        "Test5"
    ), header = TRUE, drop = list(1, 2))
    expect_equal(list(Test1 = checkDf[2:4], Test4 = checkDf1[c(
        1,
        3, 4
    )], Test5 = checkDf2[2:4]), res)
    res <- readWorksheet(wb.xls, sheet = c(
        "Test1", "Test4",
        "Test5"
    ), header = TRUE, drop = list(
        c(1, 2), c(2, 3),
        c(1, 3)
    ))
    expect_equal(list(Test1 = checkDf[3:4], Test4 = checkDf1[c(
        1,
        4
    )], Test5 = checkDf2[c(2, 4)]), res)
    res <- readWorksheet(wb.xlsx, sheet = c(
        "Test1", "Test4",
        "Test5"
    ), header = TRUE, drop = list(
        c(1, 2), c(2, 3),
        c(1, 3)
    ))
    expect_equal(list(Test1 = checkDf[3:4], Test4 = checkDf1[c(
        1,
        4
    )], Test5 = checkDf2[c(2, 4)]), res)
    res <- readWorksheet(wb.xls, sheet = c(
        "Test1", "Test4",
        "Test5", "AAA"
    ), header = TRUE, drop = list(
        c(1, 2),
        c(2, 3)
    ))
    expect_equal(list(Test1 = checkDf[3:4], Test4 = checkDf1[c(
        1,
        4
    )], Test5 = checkDf2[3:4], AAA = testAAA[c(1)]), res)
    res <- readWorksheet(wb.xlsx, sheet = c(
        "Test1", "Test4",
        "Test5", "AAA"
    ), header = TRUE, drop = list(
        c(1, 2),
        c(2, 3)
    ))
    expect_equal(list(Test1 = checkDf[3:4], Test4 = checkDf1[c(
        1,
        4
    )], Test5 = checkDf2[3:4], AAA = testAAA[c(1)]), res)
    target1 <- data.frame(Col1 = c(
        NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, 7, NA, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA
    ), Col2 = c(
        NA, NA, NA, 3, NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA, 13
    ), Col3 = c(
        NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col4 = c(
        NA,
        NA, NA, 4, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, 9, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col5 = c(
        1,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col6 = c(
        NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col7 = c(
        NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, 10, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col8 = c(
        2,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col9 = c(
        NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col10 = c(
        NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, 11, NA, NA, NA, NA, NA
    ), Col11 = c(
        NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col12 = c(
        NA,
        NA, NA, NA, NA, NA, 5, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, 12, NA, NA, NA, NA, NA
    ), Col13 = c(
        NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col14 = c(
        NA,
        NA, NA, NA, NA, NA, 6, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col15 = c(
        NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col16 = c(
        NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, 8, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ))
    target2 <- data.frame(
        Col1 = c(9, NA, NA, NA, NA, NA), Col2 = c(
            NA,
            NA, NA, NA, NA, NA
        ), Col3 = c(NA, NA, NA, NA, NA, NA),
        Col4 = c(10, NA, NA, NA, NA, NA), Col5 = c(
            NA, NA, NA,
            NA, NA, NA
        ), Col6 = c(NA, NA, NA, NA, NA, NA), Col7 = c(
            NA,
            NA, NA, NA, NA, 11
        )
    )
    target3 <- data.frame(Col1 = c(
        NA, NA, NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA
    ), Col2 = c(
        NA, NA, NA, 9, NA, NA,
        NA, NA, NA, NA, NA, NA
    ), Col3 = c(
        NA, NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA
    ), Col4 = c(
        NA, NA, NA, NA,
        NA, NA, NA, NA, NA, NA, NA, NA
    ), Col5 = c(
        NA, NA, NA,
        10, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col6 = c(
        NA, NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col7 = c(
        NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ), Col8 = c(
        NA,
        NA, NA, NA, NA, NA, NA, NA, 11, NA, NA, NA
    ), Col9 = c(
        NA,
        NA, NA, NA, NA, NA, NA, NA, NA, NA, NA, NA
    ))
    target4 <- as.data.frame(matrix(NA, nrow = 10, ncol = 8))
    names(target4) <- paste("Col", 1:8, sep = "")
    target5 <- data.frame(Col1 = c(NA, NA, NA, NA, 4, NA), Col2 = c(
        NA,
        1, NA, NA, NA, NA
    ))
    target6 <- data.frame(Col1 = c(NA, NA, NA, NA), Col2 = c(
        NA,
        NA, NA, 4
    ), Col3 = c(1, NA, NA, NA), Col4 = c(
        NA, NA,
        NA, NA
    ), Col5 = c(NA, NA, NA, NA))
    target7 <- data.frame(Col1 = c(NA, NA, NA, 4), Col2 = c(
        1,
        NA, NA, NA
    ))
    res <- readWorksheet(wb.xls,
        sheet = "BoundingBox", autofitRow = TRUE,
        autofitCol = TRUE, header = FALSE
    )
    expect_equal(target1, res)
    res <- readWorksheet(wb.xls,
        sheet = "BoundingBox", autofitRow = FALSE,
        autofitCol = FALSE, header = FALSE
    )
    expect_equal(target1, res)
    res <- readWorksheet(wb.xls,
        sheet = "BoundingBox", startRow = 20,
        startCol = 5, endRow = 31, endCol = 13, autofitRow = TRUE,
        autofitCol = TRUE, header = FALSE
    )
    expect_equal(target2, res)
    res <- readWorksheet(wb.xls,
        sheet = "BoundingBox", startRow = 20,
        startCol = 5, endRow = 31, endCol = 13, autofitRow = FALSE,
        autofitCol = FALSE, header = FALSE
    )
    expect_equal(target3, res)
    res <- readWorksheet(wb.xls,
        sheet = "BoundingBox", startRow = 12,
        startCol = 5, endRow = 21, endCol = 12, autofitRow = TRUE,
        autofitCol = TRUE, header = FALSE
    )
    expect_equal(data.frame(), res)
    res <- readWorksheet(wb.xls,
        sheet = "BoundingBox", startRow = 12,
        startCol = 5, endRow = 21, endCol = 12, autofitRow = FALSE,
        autofitCol = FALSE, header = FALSE
    )
    expect_equal(target4, res)
    res <- readWorksheet(wb.xls,
        sheet = "BoundingBox", startRow = 6,
        startCol = 5, endRow = 11, endCol = 9, autofitRow = FALSE,
        autofitCol = TRUE, header = FALSE
    )
    expect_equal(target5, res)
    res <- readWorksheet(wb.xls,
        sheet = "BoundingBox", startRow = 6,
        startCol = 5, endRow = 11, endCol = 9, autofitRow = TRUE,
        autofitCol = FALSE, header = FALSE
    )
    expect_equal(target6, res)
    res <- readWorksheet(wb.xls,
        sheet = "BoundingBox", startRow = 6,
        startCol = 5, endRow = 11, endCol = 9, autofitRow = TRUE,
        autofitCol = TRUE, header = FALSE
    )
    expect_equal(target7, res)
    res <- readWorksheet(wb.xlsx,
        sheet = "BoundingBox", autofitRow = TRUE,
        autofitCol = TRUE, header = FALSE
    )
    expect_equal(target1, res)
    res <- readWorksheet(wb.xlsx,
        sheet = "BoundingBox", autofitRow = FALSE,
        autofitCol = FALSE, header = FALSE
    )
    expect_equal(target1, res)
    res <- readWorksheet(wb.xlsx,
        sheet = "BoundingBox", startRow = 20,
        startCol = 5, endRow = 31, endCol = 13, autofitRow = TRUE,
        autofitCol = TRUE, header = FALSE
    )
    expect_equal(target2, res)
    res <- readWorksheet(wb.xlsx,
        sheet = "BoundingBox", startRow = 20,
        startCol = 5, endRow = 31, endCol = 13, autofitRow = FALSE,
        autofitCol = FALSE, header = FALSE
    )
    expect_equal(target3, res)
    res <- readWorksheet(wb.xlsx,
        sheet = "BoundingBox", startRow = 12,
        startCol = 5, endRow = 21, endCol = 12, autofitRow = TRUE,
        autofitCol = TRUE, header = FALSE
    )
    expect_equal(data.frame(), res)
    res <- readWorksheet(wb.xlsx,
        sheet = "BoundingBox", startRow = 12,
        startCol = 5, endRow = 21, endCol = 12, autofitRow = FALSE,
        autofitCol = FALSE, header = FALSE
    )
    expect_equal(target4, res)
    res <- readWorksheet(wb.xlsx,
        sheet = "BoundingBox", startRow = 6,
        startCol = 5, endRow = 11, endCol = 9, autofitRow = FALSE,
        autofitCol = TRUE, header = FALSE
    )
    expect_equal(target5, res)
    res <- readWorksheet(wb.xlsx,
        sheet = "BoundingBox", startRow = 6,
        startCol = 5, endRow = 11, endCol = 9, autofitRow = TRUE,
        autofitCol = FALSE, header = FALSE
    )
    expect_equal(target6, res)
    res <- readWorksheet(wb.xlsx,
        sheet = "BoundingBox", startRow = 6,
        startCol = 5, endRow = 11, endCol = 9, autofitRow = TRUE,
        autofitCol = TRUE, header = FALSE
    )
    expect_equal(target7, res)
    wb.xls <- loadWorkbook("resources/testCachedValues.xls",
        create = FALSE
    )
    wb.xlsx <- loadWorkbook("resources/testCachedValues.xlsx",
        create = FALSE
    )
    ref.xls.uncached <- readWorksheet(wb.xls, "AllLocal", useCachedValues = FALSE)
    ref.xls.cached <- readWorksheet(wb.xls, "AllLocal", useCachedValues = TRUE)
    expect_equal(ref.xls.cached, ref.xls.uncached)
    ref.xlsx.uncached <- readWorksheet(wb.xlsx, "AllLocal", useCachedValues = FALSE)
    ref.xlsx.cached <- readWorksheet(wb.xlsx, "AllLocal", useCachedValues = TRUE)
    expect_equal(ref.xlsx.cached, ref.xlsx.uncached)
    expect_equal(ref.xls.uncached, ref.xlsx.uncached)
    onErrorCell(wb.xls, XLC$ERROR.STOP)
    expect_error(readWorksheet(wb.xls, "HeaderRemote", useCachedValues = FALSE))
    expect_error(readWorksheet(wb.xls, "BodyRemote", useCachedValues = FALSE))
    expect_error(readWorksheet(wb.xls, "AllRemote", useCachedValues = FALSE))
    onErrorCell(wb.xlsx, XLC$ERROR.STOP)
    expect_error(readWorksheet(wb.xlsx, "HeaderRemote", useCachedValues = FALSE))
    expect_error(readWorksheet(wb.xlsx, "BodyRemote", useCachedValues = FALSE))
    expect_error(readWorksheet(wb.xlsx, "AllRemote", useCachedValues = FALSE))
    res <- readWorksheet(wb.xls, "HeadersRemote", useCachedValues = TRUE)
    expect_equal(res, ref.xls.uncached)
    res <- readWorksheet(wb.xls, "BodyRemote", useCachedValues = TRUE)
    expect_equal(res, ref.xls.uncached)
    res <- readWorksheet(wb.xls, "BothRemote", useCachedValues = TRUE)
    expect_equal(res, ref.xls.uncached)
    res <- readWorksheet(wb.xlsx, "HeadersRemote", useCachedValues = TRUE)
    expect_equal(res, ref.xls.uncached)
    res <- readWorksheet(wb.xlsx, "BodyRemote", useCachedValues = TRUE)
    expect_equal(res, ref.xls.uncached)
    res <- readWorksheet(wb.xlsx, "BothRemote", useCachedValues = TRUE)
    expect_equal(res, ref.xls.uncached)
    res <- readWorksheetFromFile("resources/testBug52.xlsx",
        sheet = 1, useCachedValues = TRUE
    )
    expected <- data.frame(Var1 = c(2, 4, 6), Var2 = c(
        "2", "nope",
        "6"
    ), Var3 = c(NA, 4, 6), Var4 = c(2, 4, 6), stringsAsFactors = FALSE)
    expect_equal(res, expected)
    expected <- data.frame(B = 1:5, row.names = letters[1:5])
    res <- readWorksheetFromFile("resources/testBug49.xlsx",
        sheet = 1, rownames = 1
    )
    expect_equal(res, expected)
    expected <- data.frame(
        A = c("2003-04-06", "2014-10-30", "abc"),
        stringsAsFactors = FALSE
    )
    res <- readWorksheetFromFile("resources/testBug53.xlsx",
        sheet = 1, dateTimeFormat = "%Y-%m-%d"
    )
    expect_equal(res, expected)
    expected <- data.frame(A = as.POSIXct(c(
        "2015-12-01", "2015-11-17",
        "1984-01-11"
    )))
    res <- readWorksheetFromFile("resources/testBug53.xlsx",
        sheet = 2, colTypes = "POSIXt", forceConversion = TRUE
    )
    expect_equal(res, expected)
    wbSparse.xlsx <- loadWorkbook("resources/testReadWorksheetSparseBitSet.xlsx",
        create = FALSE
    )
    sparseSheet <- readWorksheet(wbSparse.xlsx, "hist")
})
