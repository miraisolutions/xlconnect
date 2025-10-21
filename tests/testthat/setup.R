# This file sets up global options for the test suite environment.
# Options set here using withr::local_options and teardown_env()
# will be automatically reset after the test suite finishes.

# Define locale to be set locally for the test environment
withr::local_locale(c(LC_NUMERIC = "C"), .local_envir = testthat::teardown_env())
# Set Java Locale to US
library(rJava)
jlocale <- J("java.util.Locale")
jlocale$setDefault(jlocale$US)
if (!getOption("FULL.TEST.SUITE", default = FALSE)) {
  Sys.setenv("OMP_THREAD_LIMIT" = 1)
}
withr::local_options(.new = list(XLConnect.setCustomAttributes = TRUE), .local_envir = testthat::teardown_env())
# Clean up variables from this script's environment
rm(jlocale)
