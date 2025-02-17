# XLConnect 1.2.0

## Test Environments

* Ubuntu 22.04 on Github Actions
* Ubuntu 20.04 on Github Actions
* Windows 10 (2019) on Github Actions
* macOS 12.7.4 on Github Actions

## R CMD check results

There were no ERRORs or WARNINGs

There was 1 NOTE:

```sh
* checking installed package size ... NOTE
  installed size is 28.7Mb
  sub-directories of 1Mb or more:
    java  26.1Mb
```

Justification: XLConnect uses a java component which we maintain in a separate project, as well as Apache POI 5.4.x and its dependencies. At install time, the presence of these dependencies in the correct version is checked; if missing, they are downloaded into XLConnect's installation directory. Apache POI 5.4.x is not yet available from major distributions' package managers at the time of writing. In addition, _poi-ooxml-full-<version>.jar_ is required, which is not distributed via package managers. See point 3. of [the POI FAQ](https://poi.apache.org/help/faq.html) for more information.

## Downstream dependencies

_revdepcheck_ found 7 downstream dependencies; they were all checked successfully.
