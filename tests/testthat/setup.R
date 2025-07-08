# This file sets up global options for the test suite environment.
# Options set here using withr::local_options and teardown_env()
# will be automatically reset after the test suite finishes.

# Ensure withr is available. It's a dependency of testthat.
# testthat::teardown_env() is the correct environment for withr functions
# when used within testthat's setup files.

# Get current Java parameters to append to them, not overwrite others.
options("FULL.TEST.SUITE" = Sys.getenv("FULL_TEST_SUITE") == "1")

original_java_params <- getOption("java.parameters")
j_params <- c(
  original_java_params,
  "-XX:+UseParallelGC",
  "-XX:ParallelGCThreads=1",
  paste0("-Duser.timezone=", Sys.timezone())
)
if (!getOption("FULL.TEST.SUITE")) {
  j_params = c(j_params, "-XX:ActiveProcessorCount=1")
}

# Define options and locale to be set locally for the test environment
options_to_set <- list(
  java.parameters = j_params, # Keep java.parameters under withr for reset
  encoding = "UTF-8"
)

withr::local_options(options_to_set, .local_envir = testthat::teardown_env())
withr::local_locale(
  c(LC_NUMERIC = "C"),
  .local_envir = testthat::teardown_env()
)

# Load library built by R CMD check
library(package = "XLConnect", character.only = TRUE)
require(rJava)

# Set Java Locale to US
jlocale <- J("java.util.Locale")
jlocale$setDefault(jlocale$US)

# Clean up variables from this script's environment
# Note: new_java_params was already not defined, removing it from rm()
# j_params is defined earlier and used in options_to_set, so it's cleaned up here.
rm(original_java_params, j_params, options_to_set, locale_to_set, jlocale)
