<img src="man/figures/logo.png" align="right" width="15%" height="15%"/>

XLConnect: Excel Connector for R
================================

[![CRAN\_Status\_Badge](https://www.r-pkg.org/badges/version/XLConnect)](https://cran.r-project.org/package=XLConnect)
[![rdocumentation.org](https://www.rdocumentation.org/badges/version/XLConnect)](https://www.rdocumentation.org/packages/XLConnect)
[![CranLogsTotal](https://cranlogs.r-pkg.org/badges/grand-total/XLConnect?color=yellow)](#xlconnect-excel-connector-for-r)
[![CranLogsMonthly](https://cranlogs.r-pkg.org/badges/XLConnect?color=blue)](#xlconnect-excel-connector-for-r)
[![CranLogsWeekly](https://cranlogs.r-pkg.org/badges/last-week/XLConnect?color=green)](#xlconnect-excel-connector-for-r)
[![R](https://github.com/miraisolutions/xlconnect/workflows/R/badge.svg?branch=master&event=push)](https://github.com/miraisolutions/xlconnect/actions?query=workflow%3AR+branch%3Amaster)
[![codecov](https://codecov.io/gh/miraisolutions/xlconnect/branch/master/graph/badge.svg)](https://app.codecov.io/gh/miraisolutions/xlconnect)

XLConnect is a comprehensive and cross-platform R package for manipulating Microsoft Excel files from within R. XLConnect differs from other related R packages in that it is completely cross-platform and as such runs under Windows, Unix/Linux and Mac (32- and 64-bit). Moreover, it does not require any installation of Microsoft Excel or any other special drivers to be able to read & write Excel files. The only requirement is a recent version of a Java Runtime Environment (JRE).

The package can easily be installed from <a href="https://cran.r-project.org/package=XLConnect">CRAN</a> via `install.packages("XLConnect")`. In order to get started have a look at the <a href="https://cran.r-project.org/package=XLConnect/vignettes/XLConnect.pdf">XLConnect</a> and <a href="https://cran.r-project.org/package=XLConnect/vignettes/XLConnectImpatient.pdf">XLConnect for the Impatient</a> package vignettes, the numerous demos available via `demo(package = "XLConnect")` or browse through the comprehensive <a href="https://cran.r-project.org/package=XLConnect/XLConnect.pdf">reference manual</a>.

Alternatively, you may install XLConnect directly from our <a href="https://github.com/miraisolutions/xlconnect">github repository</a> using the excellent <a href="https://github.com/r-lib/devtools">devtools</a> package:

```r
require(devtools)

# Installs the master branch of XLConnect (= current development version)
install_github("miraisolutions/xlconnect")

# Installs XLConnect with the given version, e.g. 1.0.2
install_github("miraisolutions/xlconnect", ref = "<version>")
```

Please log any enhancement requests or bug reports with a simple and self-contained reproducible example as an issue on our <a href="https://github.com/miraisolutions/xlconnect">github repository</a>.
For other questions you may also use <a href="https://stackoverflow.com/questions/tagged/xlconnect">Stackoverflow</a>.

Build for release on CRAN
-------------------------

```r
devtools::build(args=c("--compact-vignettes=gs+qpdf"))
```
