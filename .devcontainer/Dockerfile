# rocker/verse:latest tracks the current R release (4.6.1 at the time of writing),
# which is what the "release" rows of the CI matrix and CRAN check against.
# Pin to an explicit tag (e.g. rocker/verse:4.6.1) to freeze the R version.
#
# The base image already provides pandoc, quarto, TeX Live, the tinytex R
# package, ghostscript, qpdf, devtools, testthat, ggplot2 and rJava, so only
# what is missing on top of it is installed below.
FROM rocker/verse:latest

ENV DEBIAN_FRONTEND=noninteractive

# tidy is needed by `R CMD check --as-cran` to validate the generated HTML.
# Java 25 is the newest JDK covered by the CI matrix and supersedes the
# default-jdk (21) shipped with rocker/verse.
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        gpg \
        openjdk-25-jdk-headless \
        tidy && \
    rm -rf /var/lib/apt/lists/*

# Point R at the new JDK and rebuild rJava from source against it: the copy in
# the base image is linked against the older default-jdk. The assertion turns a
# broken Java setup into a build failure rather than a puzzling runtime error.
RUN R CMD javareconf && \
    install2.r --error -r https://cloud.r-project.org rJava && \
    Rscript -e "{\
        library(rJava);\
        .jinit();\
        version <- .jcall('java/lang/System', 'S', 'getProperty', 'java.version');\
        cat('rJava runs on Java', version, '\n');\
        stopifnot(as.integer(sub('[.].*', '', version)) >= 25);\
        }"

# Package dependencies and check tooling not covered by the base image.
RUN install2.r --error --skipinstalled qpdf withr zoo
RUN installGithub.r r-lib/revdepcheck

# (more latex libs would be required to include pdf related checks)
RUN Rscript -e "{\
    library(tinytex);\
    tlmgr_update(self=TRUE, all=FALSE);\
    tlmgr_install('datetime');\
    tlmgr_install('hyperref');\
    tlmgr_install('url');\
    tlmgr_install('fmtcount');\
    tlmgr_install('float');\
    tlmgr_install('natbib');\
    tinytex:::install_yihui_pkgs();\
    }"
