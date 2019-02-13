if [[ $TRAVIS_OS_NAME == 'osx' ]]; then
  R CMD javareconf
  echo hello osx
  #sudo tlmgr init-usertree
  sudo tlmgr install datetime fmtcount
  echo bye osx
else
  # https://www.linuxuprising.com/2018/10/how-to-install-oracle-java-11-in-ubuntu.html

  echo oracle-java11-installer shared/accepted-oracle-license-v1-2 select true | sudo /usr/bin/debconf-set-selections

  # note that java10 is not available through this anymore
  sudo add-apt-repository ppa:linuxuprising/java -y
  sudo apt update -q
  # below triggers oracle-java11-set-default through 'Recommends'
  sudo apt --install-recommends install oracle-java11-installer
fi

java -version
