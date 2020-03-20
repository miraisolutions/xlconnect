# Patch
This is a patch to add a java version check as required by the CRAN team (mail from Brian Ripley).

# Test Environments

* Ubuntu 18.04 on Travis CI
* Ubuntu 16.04 on Travis CI
* CRAN winbuilder
* Ubuntu Linux 16.04 LTS, R-release, GCC on r-hub

# R CMD check results

There were no ERRORs or WARNINGs

There was 1 NOTE:
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

Justification: this URL is working when accessed in a browser and with httr::GET()

In addition, CRAN found 1 NOTE in the submission presently being patched:

> Check: for detritus in the temp directory, Result: NOTE
  Found the following files/directories:
    'calibre_4.99.4_tmp_47pf1vxy' 'calibre_4.99.4_tmp_ol_3_tqm'
    'runtime-hornik'

Justification: this was tolerated on the previous version that this is patching (mail by Uwe Ligges).

# Downstream dependencies

_revdepcheck_ found 9 downstream dependencies; they were all checked successfully.

# Miscellaneous

I'm submitting this update because the maintainer has reduced availability at the moment due to professional duties.
You may of course e-mail him for confirmation.
