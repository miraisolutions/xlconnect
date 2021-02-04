# Test Environments

* Ubuntu 18.04 on Github Actions
* Ubuntu 16.04 on Github Actions
* Windows 10 (2019) on Github Actions
* macOS 10.15 on Github Actions

# R CMD check results

There were no ERRORs or WARNINGs

There was 1 NOTE:
> checking installed package size ... NOTE
    installed size is 36.3Mb
    sub-directories of 1Mb or more:
      java  34.0Mb

Justification: XLConnect uses a java component which we maintain in a separate project, as well as Apache POI 4.1.x and
its dependencies. At install time, the presence of these dependencies in the correct version is checked; if missing,
they are downloaded into XLConnect's installation directory. Apache POI 4.1.x is not yet available from major 
distributions' package managers at the time of writing. In addition, the full ooxml-schemas-1.4.jar is required, which 
is not distributed via package managers. See http://poi.apache.org/help/faq.html#faq-N10109 for more information.

# Downstream dependencies

_revdepcheck_ found 10 downstream dependencies; they were all checked successfully.
