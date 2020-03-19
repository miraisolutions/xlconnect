# Resubmission
This is a resubmission. In this version I have:

* removed temporary debug code
* (previously) extended the installation process to check for the presence of required JAR files,
and to download these if missing.

# Test Environments

* Ubuntu 18.04 on Travis CI
* Ubuntu 16.04 on Travis CI
* CRAN winbuilder
* Ubuntu Linux 16.04 LTS, R-release, GCC on r-hub

# R CMD check results

There were no ERRORs or WARNINGs

There were 2 NOTEs:
> checking installed package size ... NOTE
  installed size is 28.8Mb
  sub-directories of 1Mb or more:
    java  26.2Mb

Justification: XLConnect uses a java component which we maintain in a separate project, as well as Apache POI 4.1.1 and
its dependencies. At install time, the presence of these dependencies in the correct version is checked; if missing,
they are downloaded into XLConnect's installation directory. XLConnect depends on Apache POI 4.1.x, which is not yet
available from major distributions' package managers at the time of writing. In addition, it requires the full
ooxml-schemas-1.4.jar, which is not distributed via package managers. See http://poi.apache.org/help/faq.html#faq-N10109
for more information.

> Found the following (possibly) invalid URLs:
  URL: http://office.microsoft.com/en-001/excel-help/overview-of-excel-tables-HA010048546.aspx
    From: man/readTable-methods.Rd
    Status: Error
    Message: libcurl error code 52:
      Empty reply from server

Justification: this URL is working when accessed in a browser.

In addition, CRAN finds 3 NOTES:

> Found the following (possibly) invalid file URI:
  URI: xlconnect@mirai-solutions.com
    From: README.md

Justification: this is an email address, not a file URI

> Found the following files/directories:
    'XLConnect.xlsx' 'autofilter.xlsx' 'cellstyles.xlsx' [...]

Justification: these are examples part of our documentation and tests.

> Check: for detritus in the temp directory, Result: NOTE
  Found the following files/directories:
    'calibre_4.99.4_tmp_47pf1vxy' 'calibre_4.99.4_tmp_ol_3_tqm'
    'runtime-hornik'

Justification: these files do not appear in any other systems where I run devtools check / install.For example, none of
the local environments I use even have _calibre_ installed. This seems related to the specific machine CRAN runs this on.

# Downstream dependencies

_revdepcheck_ found 9 downstream dependencies; they were all checked successfully.

# Miscellaneous

I'm submitting this update because the maintainer has reduced availability at the moment due to professional duties.
You may of course e-mail him for confirmation.
