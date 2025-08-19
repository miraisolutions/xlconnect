test_that("test.dumpAndRestore - full test suite only", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")

  require(datasets)
  pos = "package:datasets"
  objs = ls(pos = pos)
  idx = sapply(objs, function(obj) is.data.frame(get(obj, pos = pos)))
  objs = objs[idx]
  for (file in c("testDumpAndRestore.xls", "testDumpAndRestore.xlsx")) {
    out = xlcDump(objs, file = file, pos = pos, overwrite = TRUE)
    xlcRestore(file = file, pos = globalenv(), overwrite = TRUE)
    sapply(names(out)[out], function(obj) {
      data.orig = normalizeDataframe(get(obj, pos = pos))
      data.restored = get(obj)
      expect_equal(data.orig, data.restored, ignore_attr = c("worksheetScope"))
      expect_equal(attr(data.orig, "row.names"), attr(data.restored, "row.names"))
    })
  }
})
