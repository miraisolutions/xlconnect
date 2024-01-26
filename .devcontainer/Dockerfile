FROM rocker/verse:devel

RUN apt-get update
RUN apt-get install -y libxml2-dev libssl-dev gpg
# texlive
RUN apt-get install -y libqpdf28 qpdf ghostscript tidy
# RUN tlmgr install xcolor  # (more latex libs would be required to include pdf related checks)
RUN install2.r --error devtools rJava RUnit zoo qpdf languageserver
RUN installGithub.r r-lib/revdepcheck
RUN apt-get install -y pandoc
RUN install2.r --error ggplot2 tinytex
RUN Rscript -e "{\
    library(tinytex);\
    tlmgr_install('datetime');\
    tlmgr_install('inputenc');\
    tlmgr_install('hyperref');\
    tlmgr_install('url');\
    tlmgr_install('fmtcount');\
    tlmgr_install('float');\
    tlmgr_install('natbib');\
    tinytex:::install_yihui_pkgs();\
}"
