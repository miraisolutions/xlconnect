test_that("test.dumpAndRestore", {
    skip_on_cran() # Skip this test on CRAN due to memory issues
    if (getOption("FULL.TEST.SUITE")) {
        require(datasets)
        pos = "package:datasets"
        objs = ls(pos = pos)
        idx = sapply(objs, function(obj) is.data.frame(get(obj, 
            pos = pos)))
        objs = objs[idx]
        for (file in c("testDumpAndRestore.xls", "testDumpAndRestore.xlsx")) {
            out = xlcDump(objs, file = file, pos = pos, overwrite = TRUE)
            xlcRestore(file = file, pos = globalenv(), overwrite = TRUE)
            sapply(names(out)[out], function(obj) {
                data.orig = normalizeDataframe(get(obj, pos = pos))
                data.restored = get(obj)
                expect_equal(data.orig, data.restored, check.attributes = FALSE,
                  check.names = TRUE)
                expect_equal(attr(data.orig, "row.names"), attr(data.restored,
                  "row.names"))
            })
        }
    }
})

