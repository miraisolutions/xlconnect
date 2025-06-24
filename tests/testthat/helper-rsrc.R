# Define the rsrc function for testthat tests
# Copied from tests/run_tests.R

# Path to unit tests for 'R CMD check' and as part of public API
# Note: 'inst/unitTests' is the path relative to the package root
# when running tests with devtools or R CMD check.
# For testthat, the current directory is tests/testthat,
# so we need to go up two levels to the package root.
.pathToUnitTests <- file.path("..", "..", "inst", "unitTests")

rsrc <- function(resource) {
  file.path(.pathToUnitTests, resource)
}
