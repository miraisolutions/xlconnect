test_that("test.workbook.getReferenceFormula", {
    wb.xls <- loadWorkbook("resources/testWorkbookReferenceFormula.xls"),
        create = FALSE)
    wb.xlsx <- loadWorkbook("resources/testWorkbookReferenceFormula.xlsx"),
        create = FALSE)
    expect_true(getReferenceFormula(wb.xls, "FirstName") == "Tabelle1!$A$1")
    expect_true(substring(getReferenceFormula(wb.xls, "SecondName"), 
        1, 5) == "#REF!")
    expect_true(getReferenceFormula(wb.xlsx, "FirstName") == 
        "Tabelle1!$A$1")
    expect_true(substring(getReferenceFormula(wb.xlsx, "SecondName"), 
        1, 5) == "#REF!")
    expect_formula <- "Tabelle1!$A$1"
    attributes(expect_formula) <- list(worksheetScope = "")
    expect_equal(expect_formula, getReferenceFormula(wb.xls, 
        "FirstName"))
    expect_formula <- "#REF!"
    attributes(expect_formula) <- list(worksheetScope = "")
    expect_equal(expect_formula, substring(getReferenceFormula(wb.xls, 
        "SecondName"), 1, 5))
    expect_formula <- "Tabelle1!$A$1"
    attributes(expect_formula) <- list(worksheetScope = "")
    expect_equal(expect_formula, getReferenceFormula(wb.xlsx, 
        "FirstName"))
    expect_formula <- "#REF!"
    attributes(expect_formula) <- list(worksheetScope = "")
    expect_equal(expect_formula, substring(getReferenceFormula(wb.xlsx, 
        "SecondName"), 1, 5))
})

