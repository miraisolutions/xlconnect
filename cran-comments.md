This is a resubmission of XLConnect 1.0.2 to fix obsolete URLs in the package documentation.
# Test Environments

* Ubuntu 18.04 on Github Actions
* Ubuntu 16.04 on Github Actions
* Windows 10 (2019) on Github Actions
* macOS 10.15 on Github Actions

# R CMD check results

There were no ERRORs or WARNINGs

There were 2 NOTEs:
> checking installed package size ... NOTE
    installed size is 36.3Mb
    sub-directories of 1Mb or more:
      java  34.0Mb

Justification: XLConnect uses a java component which we maintain in a separate project, as well as Apache POI 4.1.x and its dependencies. At install time, the presence of these dependencies in the correct version is checked; if missing, they are downloaded into XLConnect's installation directory. Apache POI 4.1.x is not yet available from major distributions' package managers at the time of writing. In addition, the full ooxml-schemas-1.4.jar is required, which is not distributed via package managers. See https://poi.apache.org/help/faq.html#faq-N10109 for more information.

> checking CRAN incoming feasibility ... NOTE
  Maintainer: ‘Martin Studer <martin.studer@mirai-solutions.com>’
  
  Found the following (possibly) invalid URLs:
    URL: https://www.ozgrid.com/Excel/CustomFormats.htm
      From: man/setDataFormat-methods.Rd
            man/setDataFormatForType-methods.Rd
      Status: 301
      Message: Moved Permanently

Justification: the linked website seems to uses cloudflare DDoS protection, which results in a redirect when accessing the URL.

In addition, there was a NOTE on the Debian machine used by CRAN to run tests.

> Found the following files/directories:
  ‘calibre_5.10.1_tmp_2jgz_y1p’ ‘calibre_5.10.1_tmp_qry3rapb’

This seems unique to that machine and already appeared when submitting the previous version.
# Downstream dependencies

_revdepcheck_ found 10 downstream dependencies; they were all checked successfully.
