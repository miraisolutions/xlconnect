FROM rocker/verse:devel

RUN apt-get update
RUN apt-get install -y libxml2-dev libssl-dev gpg
# texlive
RUN apt-get install -y libqpdf26 qpdf ghostscript
RUN install2.r --error devtools rJava RUnit zoo qpdf languageserver
RUN apt-get install -y pandoc
RUN install2.r --error ggplot2
