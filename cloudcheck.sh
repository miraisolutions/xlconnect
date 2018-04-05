apt-get update
apt install -y --no-install-recommends xzdec pandoc qpdf
tlmgr init-usertree
tlmgr install datetime etoolbox fmtcount ulem pgf
install2.r rJava RUnit ggplot2 zoo

mkdir -p /tmp/results
cd /tmp/results
git clone --depth=1 git://github.com/miraisolutions/xlconnectjars.git
git clone --depth=1 git://github.com/miraisolutions/xlconnect.git

R CMD build --md5 xlconnectjars
R CMD check --as-cran XLConnectJars_*.tar.gz
R CMD INSTALL XLConnectJars_*.tar.gz

# Do not drop unit tests to run full test suite
cp xlconnect/.Rbuildignore .Rbuildignore.bak
sed -i "/unitTests/d" xlconnect/.Rbuildignore

R CMD build --compact-vignettes --md5 xlconnect
FULL_TEST_SUITE=1 R CMD check --as-cran XLConnect_*.tar.gz
R CMD INSTALL XLConnect_*.tar.gz

mv .Rbuildignore.bak xlconnect/.Rbuildignore
R CMD build --compact-vignettes --md5 xlconnect
		
cd /tmp
tar -zcvf /exchange/results-$1.tar.gz results
