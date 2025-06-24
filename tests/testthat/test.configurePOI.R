test_that("test.configurePOI", {
    configurePOI(zip_max_files = 1L)
    expect_error(loadWorkbook("resources/testLoadWorkbook.xlsx"))
    configurePOI(zip_min_inflate_ratio = 0.99)
    expect_error(readWorksheetFromFile("resources/testZipBomb.xlsx",
        sheet = 1
    ))
    configurePOI(zip_max_entry_size = 1L)
    expect_error(readWorksheetFromFile("resources/testWorkbookReadWorksheet.xlsx", sheet = 1))
    configurePOI(zip_entry_threshold_bytes = 0L)
    readWorksheetFromFile("resources/testWorkbookReadWorksheet.xlsx", sheet = 1)
    configurePOI()
})
