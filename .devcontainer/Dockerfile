FROM rocker/verse:latest

RUN apt-get update
RUN apt-get install -y libxml2-dev libssl-dev gpg
# texlive
RUN apt-get install -y libqpdf29t64 qpdf ghostscript tidy
RUN apt-get install -y openjdk-21-jdk-headless
RUN R CMD javareconf
# RUN tlmgr install xcolor  # (more latex libs would be required to include pdf related checks)
RUN install2.r --error devtools rJava testthat zoo qpdf
RUN installGithub.r r-lib/revdepcheck
RUN apt-get install -y pandoc
RUN install2.r --error ggplot2 tinytex
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