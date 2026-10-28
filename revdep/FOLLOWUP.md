# revdepcheck follow-up (POI 6.0.0 upgrade)

> **TODO: delete this `revdep/` directory before merging.** It is committed
> (despite `.gitignore`) only to hand over revdepcheck status between sessions.

Run on 2026-10-27, CRAN XLConnect 1.3.0 (POI 5.5.1) vs dev 1.3.1.9999
(POI 6.0.0, Java >= 17), R 4.6.1, Java 21, Ubuntu 24.04. The other files in
this directory are revdepcheck output, updated by hand for D4TAlink.light.

## Results

Reverse dependencies on CRAN (all via Suggests only, no strong revdeps):

| package              | result                                               |
|:---------------------|:-----------------------------------------------------|
| JGR 1.9-2            | 0 new problems (old/new `00check.log` identical)     |
| psyverse 0.2.6       | 0 new problems (old/new `00check.log` identical)     |
| D4TAlink.light 2.1.23| 0 new problems (old/new `00check.log` identical)     |

Summary: 3 CRAN revdeps, 0 new problems, 0 failed to check.

JGR and psyverse were checked with revdepcheck on 2026-10-27. D4TAlink.light
was checked on 2026-10-28 by hand (see below). Its one WARNING ("cannot set
locale en_US.UTF-8" in the R syntax check) is the same with both versions and
comes from the container, which has no such locale.

None of the revdeps exercise XLConnect in their examples/tests/vignettes, so
their XLConnect calls were also run by hand against both versions:
D4TAlink.light (write path), JGR (`readWorksheet`), psyverse
(`writeWorksheetToFile`) all OK with both. psyverse's
`read_spreadsheet(xlsxPkg = "XLConnect")` fails identically with both versions
(psyverse bug: passes an XLConnect workbook to `openxlsx::readWorkbook`).
D4TAlink.light only uses XLConnect when openxlsx is missing, but openxlsx is
in its Imports, so that code path cannot run in normal use.

## D4TAlink.light check (2026-10-28)

- `bioconductor.org` is still blocked (403), but `www.bioconductor.org` is
  reachable and serves the same repository. Use
  `https://www.bioconductor.org/packages/3.23/bioc` as the `BioCsoft` repo.
- revdepcheck was not used: installing it from a GitHub clone was refused by
  the session's permission checks. Instead, D4TAlink.light's dependencies
  (`tools::package_dependencies(which = "most")`, without XLConnect) went into
  one library, CRAN XLConnect 1.3.0 and dev XLConnect into two more, and
  `R CMD check --no-manual` ran once with each, `_R_CHECK_FORCE_SUGGESTS_=false`.
  The installed `poi.jar` checksums confirm 5.5.1 (old) and 6.0.0 (new).
- System libraries needed: `libharfbuzz-dev libfribidi-dev libfreetype6-dev
  libpng-dev libtiff5-dev libjpeg-dev libfontconfig1-dev libwebp-dev`
  (plus `libcurl4-openssl-dev libssl-dev libxml2-dev`). `libarrow-dev` is not
  in Ubuntu apt and was not needed: `feather` builds without Arrow.

## How it was run (environment notes)

- revdepcheck and crancache are not on CRAN: clone from GitHub
  (`r-lib/revdepcheck`, `r-lib/crancache`), `R CMD INSTALL` into a separate lib.
  Also needed `apt-get install libcurl4-openssl-dev libssl-dev libxml2-dev`.
- Only `https://cran.r-project.org` is reachable (cloud.r-project.org and
  packagemanager.posit.co are blocked): `options(repos = c(CRAN = "https://cran.r-project.org"))`.
- Maven Central rate-limits (HTTP 429), which breaks XLConnect's jar download
  on load. Use a local Maven layout holding both the POI 6.0.0 and POI 5.5.1
  dependency jars and point `XLCONNECT_JAVA_REPO_URL=file:///path/to/m2` at it
  (the POI 6 set is in `~/.m2/repository` after `mvn package` in xlconnect-java).
- revdepcheck looks up Bioconductor's config even with `bioc = FALSE`; set
  `R_BIOC_VERSION` (3.23 was used) if bioconductor.org is unreachable.
- Run `R CMD check` with `JAVA_TOOL_OPTIONS` unset: the "Picked up
  JAVA_TOOL_OPTIONS" line otherwise shows up as spurious WARNINGs/NOTEs.

```r
options(repos = c(CRAN = "https://cran.r-project.org"))
env <- c(revdepcheck::revdep_env_vars(),
         XLCONNECT_JAVA_REPO_URL = "file:///path/to/m2", R_BIOC_VERSION = "3.23")
revdepcheck::revdep_check(".", bioc = FALSE, num_workers = 3,
                          timeout = as.difftime(45, units = "mins"), env = env)
```
