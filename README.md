<img src="man/figures/logo.png" align="right" width="15%" height="15%"/>

XLConnect: Excel Connector for R
================================

[![CRAN\_Status\_Badge](https://www.r-pkg.org/badges/version/XLConnect)](https://cran.r-project.org/package=XLConnect)
[![rdocumentation.org](https://api.rdocumentation.org/badges/version/XLConnect)](https://www.rdocumentation.org/packages/XLConnect)
[![CranLogsTotal](https://cranlogs.r-pkg.org/badges/grand-total/XLConnect?color=yellow)](#xlconnect-excel-connector-for-r)
[![CranLogsMonthly](https://cranlogs.r-pkg.org/badges/XLConnect?color=blue)](#xlconnect-excel-connector-for-r)
[![CranLogsWeekly](https://cranlogs.r-pkg.org/badges/last-week/XLConnect?color=green)](#xlconnect-excel-connector-for-r)
[![codecov](https://codecov.io/gh/miraisolutions/xlconnect/branch/master/graph/badge.svg)](https://app.codecov.io/gh/miraisolutions/xlconnect)

XLConnect is a comprehensive and cross-platform R package for manipulating Microsoft Excel files from within R. XLConnect differs from other related R packages in that it is completely cross-platform and as such runs under Windows, Unix/Linux and Mac (32- and 64-bit). Moreover, it does not require any installation of Microsoft Excel or any other special drivers to be able to read & write Excel files. The only requirement is a recent version of a Java Runtime Environment (JRE).

The package can easily be installed from [CRAN](https://cran.r-project.org/package=XLConnect) via `install.packages("XLConnect")`. In order to get started have a look at the [XLConnect](https://cran.r-project.org/package=XLConnect/vignettes/XLConnect.pdf) and [XLConnect for the Impatient](https://cran.r-project.org/package=XLConnect/vignettes/XLConnectImpatient.pdf) package vignettes, the numerous demos available via `demo(package = "XLConnect")` or browse through the comprehensive [reference manual](https://cran.r-project.org/package=XLConnect/XLConnect.pdf).

Alternatively, you may install XLConnect directly from our [github repository](https://github.com/miraisolutions/xlconnect) using the excellent [devtools](https://github.com/r-lib/devtools) package:

```r
require(devtools)

# Installs the master branch of XLConnect (= current development version)
install_github("miraisolutions/xlconnect")

# Installs XLConnect with the given version, e.g. 1.0.2
install_github("miraisolutions/xlconnect", ref = "<version>")
```

Please log any enhancement requests or bug reports with a simple and self-contained reproducible example as an issue on our [github repository](https://github.com/miraisolutions/xlconnect).
For other questions you may also use [Stackoverflow](https://stackoverflow.com/questions/tagged/xlconnect).

Build for release on CRAN
-------------------------

```r
devtools::build(args=c("--compact-vignettes=gs+qpdf"))
```
