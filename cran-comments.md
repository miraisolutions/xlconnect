** This submission should be published shortly after XLConnectJars 1.0.0 **

# Test Environments
* Ubuntu 18.04 on Travis CI
* Ubuntu 16.04 on Travis CI
* Mac OS X 10.13.3 on Travis CI
  * This was run on a previous git commit (705ac99..., see https://github.com/miraisolutions/xlconnect/commits/master). The
  only difference with the present release   is a change of ubuntu versions in the travis CI file. Travis's OSX images were
  recently updated to version 13 of the JDK, which we do not support yet.

# Dependencies
This package depends on XLConnectJars 1.0.0, which was submitted simultaneously. The new XLConnectJars version is
required to install this package - it may be installed e.g. using
`remotes::install_github("miraisolutions/xlconnectjars@1.0.0")`

# R CMD check results
There were no ERRORs or WARNINGs

There was 1 NOTE:
> checking installed package size ... NOTE
  installed size is  7.3Mb
  sub-directories of 1Mb or more:
    java   4.7Mb

Justification: XLConnect uses a java component which we maintain in a separate project, as well as Apache POI 4.1.1,
present as a JAR file.

# Downstream dependencies
All downstream dependencies were installed successfully with XLConnect 1.0.0 installed and loaded.

The following downstream dependencies could be checked with R CMD (ignoring suggested packages):

* growthPheno
* imageData

The following downstream dependencies were not available as source packages as installed, therefore could not be checked:

* Dominance
* LLSR
* table1xls
* upwaver

# Miscellaneous
I'm submitting this update because the maintainer has reduced availability at the moment due to professional duties.
You may of course e-mail him for confirmation.
