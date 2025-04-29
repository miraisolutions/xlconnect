# XLConnect 1.2.1

## Test Environments

* Ubuntu 24.04.2 LTS on Github Actions
* Ubuntu 22.04.5 LTS on Github Actions
* Windows Server 2019 on Github Actions
* macOS Sonoma 14.7.5 on Github Actions

## R CMD check results

There were no ERRORs or WARNINGs

There was 1 INFO:

```sh
* checking installed package size ... INFO
  installed size is 29.0Mb
  sub-directories of 1Mb or more:
    java  26.5Mb
```

Justification: XLConnect uses a Java component which we maintain in a separate project, as well as Apache POI 5.4.x and its dependencies. At install time, the presence of these dependencies in the correct version is checked; if missing, they are downloaded into XLConnect's installation directory. Apache POI 5.4.x is not yet available from major distributions' package managers at the time of writing. In addition, _poi-ooxml-full-<version>.jar_ is required, which is not distributed via package managers. See point 3. of [the POI FAQ](https://poi.apache.org/help/faq.html) for more information.

## revdepcheck results

We checked 6 reverse dependencies (1 from CRAN + 5 from Bioconductor), comparing R CMD check results across CRAN and dev versions of this package.

 * We saw 0 new problems
 * We failed to check 0 packages
