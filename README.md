XLConnect: Excel Connector for R
================================

XLConnect is a comprehensive and cross-platform R package for manipulating Microsoft Excel files from within R. XLConnect differs from other related R packages in that it is completely cross-platform and as such runs under Windows, Unix/Linux and Mac (32- and 64-bit). Moreover, it does not require any installation of Microsoft Excel or any other special drivers to be able to read & write Excel files. The only requirement is a recent version of a Java Runtime Environment (JRE).

The package can easily be installed from <a href="http://cran.r-project.org/web/packages/XLConnect">CRAN</a> via `install.packages("XLConnect")` (use `install.packages("XLConnect", type = "source")` on Mac OS X). In order to get started have a look at the <a href="http://cran.r-project.org/web/packages/XLConnect/vignettes/XLConnect.pdf">XLConnect</a> and <a href="http://cran.r-project.org/web/packages/XLConnect/vignettes/XLConnectImpatient.pdf">XLConnect for the Impatient</a> package vignettes, the numerous demos available via `demo(package = "XLConnect")` or browse through the comprehensive <a href="http://cran.r-project.org/web/packages/XLConnect/XLConnect.pdf">reference manual</a>.

Alternatively, you may install XLConnect directly from github using the excellent <a href="https://github.com/hadley/devtools">devtools</a> package:

```
require(devtools)

# Installs the master branch of XLConnect (= current development version)
install_github("xlconnect", username = "miraisolutions", ref = "master")

# Installs XLConnect 0.2-4
install_github("xlconnect", username = "miraisolutions", ref = "0.2-4")
```
