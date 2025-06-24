# This file sets up global options for the test suite environment.
# Options set here using withr::local_options and teardown_env()
# will be automatically reset after the test suite finishes.

# Ensure withr is available. It's a dependency of testthat.
# testthat::teardown_env() is the correct environment for withr functions
# when used within testthat's setup files.

# Get current Java parameters to append to them, not overwrite others.
original_java_params <- getOption("java.parameters")
new_java_params <- c(original_java_params, "-Xmx1024m")

# Define all options to be set
options_to_set <- list(
  FULL.TEST.SUITE = TRUE,
  java.parameters = new_java_params
)

# Apply these options locally for the test environment
# No need to check for withr, testthat depends on it.
withr::local_options(options_to_set, .local_envir = testthat::teardown_env())

# Clean up variables from this script's environment
rm(original_java_params, new_java_params, options_to_set)
