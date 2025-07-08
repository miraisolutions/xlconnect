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
options(java.parameters = j_params)

# Load library built by R CMD check
library(package = "XLConnect", character.only = TRUE)
require(rJava)

# Define all options to be set
options_to_set <- list(
  java.parameters = j_params
)

# Apply these options locally for the test environment
# No need to check for withr, testthat depends on it.
withr::local_options(options_to_set, .local_envir = testthat::teardown_env())

# Clean up variables from this script's environment
rm(original_java_params, new_java_params, options_to_set)
