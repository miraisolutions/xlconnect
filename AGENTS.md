# instructions for agents

## running R commands

You can run R commands by using the `Rscript` command. You can put statements on a single line, separated by a semicolon.

```sh
> Rscript -e "print('hello world'); print(1+1)"
[1] "hello world"
[1] 2
```

You need to call `library(<some_package>)` before using any function from that package.

## Setup

Try to run `Rscript -e 'library(XLConnect)'. If this command is successful, the setup has already been run once.

Prerequisite packages, including installing the checked out package itelf from source. This is only required once; it may have already been performed by a setup script.

```sh
Rscript -e "install.packages(c('rJava'))"
R CMD INSTALL .
```

Run this before you start any work, and then check that the setup worked by running `Rscript -e 'library(XLConnect)'

## Common commands

### Install the local source package

This must be run to make source code changes (under `R/`) effective (i.e. before rerunning tests).

```sh
R CMD INSTALL .
```

### Load the package

Must be done as a first command each time you run tests

```R
library(XLConnect)
```

### Run tests with testthat

Run all tests:

```R
options(FULL.TEST.SUITE=TRUE)
testthat::test_local(load_package='none')
```

Run a specific test file:

```R
options(FULL.TEST.SUITE=TRUE)
testthat::test_file("tests/testthat/test.loadWorkbook.R")
```

When making multiple changes, run the tests after each change and fix any failures before moving on to the next change.
In that case it's ok to run only the tests that you expect to be impacted.

### Run RUnit tests

```sh
export FULL_TEST_SUITE=1
Rscript tests/run_tests-old.R
```

### Clean up after tests

```sh
git clean -fdx **.xls **.xlsx *.xls *.xlsx
```

### Explore available functionality

You can list the functions / objects in a package with `ls` or `lsf`:

```R
> library(testthat)
> lsf("package:testthat")
  [1] "%>%"                       "announce_snapshot_file"   
  [3] "auto_test"                 "auto_test_package"        
  [5] "capture_condition"         "capture_error"            
  [7] "capture_expectation"       "capture_message"          
  [9] "capture_messages"          "capture_output"
  ...
```

You can then get help on a specific function by using `?`,
for example for `testthat::test_file`:

```R
?testthat::test_file"
```

This should open a read view in the terminal which can be scrolled with arrow keys or exited with `q`.

## Checks

- no xls(x) files that are tracked by git should be modified after running the tests:

   ```sh
   > git status | grep -E "xls|xlsx" | wc -l
   0
   ```

## Code style

Only write comments if they explain a non-obvious aspect of the code, or the rationale behind it.
If you have come to a certain conclusion, only write the salient points and the conclusion itself.
If a comment doesn't make sense or appears to be out of date, err on the side of removing it.

### Format the code

When you modify R code, you should use `air` to format it:

```sh
air format tests/testthat
```

formats all testthat tests;

```sh
air format tests/test-all.R
```

formats the test-all.R file (the test runner).

### Write tests

Use `withr` functionality to change the R environment for the duration of a test:

- create temporary files: `local_tempfile(...)`
- set R options: `local_options(...)`
...
