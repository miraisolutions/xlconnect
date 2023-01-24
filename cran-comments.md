XLConnect 1.0.7

# Test Environments

* Ubuntu 18.04 on Github Actions
* Ubuntu 20.04 on Github Actions
* Windows 10 (2019) on Github Actions
* macOS 10.15 on Github Actions

# R CMD check results

There were no ERRORs or WARNINGs

There was 1 NOTE:

```sh
> checking installed package size ... NOTE
    installed size is 27.0Mb
    sub-directories of 1Mb or more:
      java  24.5Mb
```

Justification: XLConnect uses a java component which we maintain in a separate project, as well as Apache POI 5.2.x and its dependencies. At install time, the presence of these dependencies in the correct version is checked; if missing, they are downloaded into XLConnect's installation directory. Apache POI 5.2.x is not yet available from major distributions' package managers at the time of writing. In addition, the full ooxml-schemas-4.1.0.jar is required, which is not distributed via package managers. See [The POI FAQ](https://poi.apache.org/help/faq.html#faq-N10109) for more information.

# Downstream dependencies

_revdepcheck_ found 10 downstream dependencies; they were all checked successfully.
