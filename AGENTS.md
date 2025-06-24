# instructions for agents

## running R commands

You can run R commands by using the `Rscript` command. You can put statements on a single line, separated by a semicolon.

```sh
> Rscript -e "print('hello world'); print(1+1)"
[1] "hello world"
[1] 2
```

You need to call `library(<some_package>)` before using any function from that package.

## IMPORTANT Setup - REQUIRED

Prerequisite packages, including installing the checked out package itelf from source:

```sh
Rscript -e "install.packages(c('rJava', 'RUnit', 'devtools', 'testthat'))"
Rscript -e "{\
   library(devtools);\
   devtools::install();\
}"
```

## Common commands

### Install the local source package

```R
devtools::install()
```

### Load the package

Must be done as a first command each time you run tests

```R
library(XLConnect)
```

### Run tests with testthat

Run all tests:

```R
testthat::test_local(load_package='none')
```

Run a specific test file:

```R
testthat::test_file("tests/testthat/test.loadWorkbook.R")
```

### Run RUnit tests

```sh
Rscript tests/run_tests.R
```

### Clean up after tests

```sh
git clean -fdx *.xls *.xlsx
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

- no xls(x) files that are tracked by git should be modifiied after running the tests:

   ```sh
   > git status | grep -E "xls|xlsx" | wc -l
   0
   ```
