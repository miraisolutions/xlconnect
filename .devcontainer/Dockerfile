FROM rocker/verse:devel 
RUN apt-get update
RUN apt-get install -y libxml2-dev libssl-dev gpg
# texlive
RUN apt-get install -y libqpdf28 qpdf ghostscript
RUN tlmgr install \
    xcolor hyperref pdftexcmds infwarerr kvoptions etoolbox datetime fmtcount texi2dvi amsmath ulem pgf
RUN install2.r --error devtools rJava RUnit zoo qpdf languageserver
RUN installGithub.r r-lib/revdepcheck
RUN apt-get install -y pandoc
RUN install2.r --error ggplot2
