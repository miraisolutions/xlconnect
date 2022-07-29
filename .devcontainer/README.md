# Purpose of the devcontainer

This devcontainer uses a `Dockerfile` and is meant to allow reproducible R CMD checks of XLConnect, as well as reverse dependency checks before new XLConnect releases.


## How to use the devcontainer

It is recommended to open the XLConnect project in VS Code, as it provides some useful features to work with devcontainers.

The `Dockerfile` will sometimes require updates and it is currently not intended to check vignettes or manual. If we wanted to check that as well, we would first need to install all the required latex packages (see commented out line in `Dockerfile`).

`devcontainer.json` allows to configure the devcontainer (see file and linked resources there).

## Concrete steps to run the checks

0. (Required only once) Checkout and open the XLConnect repo in VS Code. It is recommended to use a separate checkout folder from the XLConnect project opened in RStudio.
1. Switch to the desired branch and open the container (VS Code should have pop up dialogs for that).
2. Open the container (I use `File > Open Recent` to switch between repo and container). It may need to re-build the image, in which case there is no prompt like `root@9348tzh4o:/workspaces/xlconnect#`. I just open the repo and then the container again to get that prompt, which allows interacting with the container and running the actual commands.
3. Run `R CMD build . --no-build-vignettes` to build XLConnect locally and run `R CMD check XLConnect_<version>.tar.gz --ignore-vignettes --no-manual` to check the package (`<version>` varies).
4. Open an R session (using `R`) and if required initialize the reverse depencency check with `usethis::use_revdep()`. Run the actual checks with `revdepcheck::revdep_check(num_workers = 8, timeout = as.difftime(120, units = "mins"))` and use `revdepcheck::revdep_reset()` first in order to run it again.
