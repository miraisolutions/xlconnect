apt-get update
apt install -y --no-install-recommends xzdec pandoc qpdf
tlmgr init-usertree
tlmgr install datetime etoolbox fmtcount ulem pgf
install2.r rJava RUnit ggplot2 zoo

mkdir -p /tmp/results
cd /tmp/results
git clone --depth=1 git://github.com/miraisolutions/xlconnectjars.git
git clone --depth=1 git://github.com/miraisolutions/xlconnect.git

R CMD check --as-cran xlconnectjars
R CMD build --md5 xlconnectjars
R CMD INSTALL XLConnectJars_*.tar.gz
		
R CMD check --as-cran xlconnect
R CMD build --compact-vignettes --md5 xlconnect
R CMD INSTALL XLConnect_*.tar.gz
		
cd /tmp
tar -zcvf /out/results.tar.gz results
