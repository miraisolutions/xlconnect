# instructions for agents

## running R commands

You can run R commands by using the `Rscript` command. You can put statements on a single line, separated by a semicolon.

```sh
> Rscript -e "print('hello world'); print(1+1)"
[1] "hello world"
[1] 2
```

You need to call `library(<some_package>)` before using any function from that package.

## common commands

### install the local source package

```R
devtools::install()
```

### Load the package

Must be done as a first command each time you run tests

```R
library(XLConnect)
```

### Run tests with testthat

```R
testthat::test_dir("tests/testthat")
```

### Run RUnit tests

```sh
Rscript tests/run_tests.R
```
