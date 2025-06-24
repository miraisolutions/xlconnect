# Define the rsrc function for testthat tests

# The test_path() function from testthat automatically creates paths
# relative to the tests/testthat directory.
# The resources for these tests are in tests/testthat/resources.
# This function will now construct paths to files within that subdirectory.

rsrc <- function(resource_name) {
  testthat::test_path("resources", resource_name)
}
