# Set global options for testthat tests
options(FULL.TEST.SUITE = TRUE)

# Increase Java heap space
current_java_params <- getOption("java.parameters")
options(java.parameters = c(current_java_params, "-Xmx1024m"))
