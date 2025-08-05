require(testthat)

options("FULL.TEST.SUITE" = Sys.getenv("FULL_TEST_SUITE") == "1")

# Get current Java parameters to append to them, not overwrite others.
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

options_to_set <- list(
  java.parameters = j_params, # Keep java.parameters under withr for reset
  encoding = "UTF-8"
)

options(options_to_set)

rm(original_java_params, j_params, options_to_set)
# Load library built by R CMD check
library(package = "XLConnect", character.only = TRUE)

test_check("XLConnect")
