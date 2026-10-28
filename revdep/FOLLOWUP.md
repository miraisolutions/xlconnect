# revdepcheck follow-up (POI 6.0.0 upgrade)

> **TODO: delete this `revdep/` directory before merging.** It is committed
> (despite `.gitignore`) only to hand over revdepcheck status between sessions.

Run on 2026-10-27, CRAN XLConnect 1.3.0 (POI 5.5.1) vs dev 1.3.1.9999
(POI 6.0.0, Java >= 17), R 4.6.1, Java 21, Ubuntu 24.04. The other files in
this directory are the raw revdepcheck output.

## Results

Reverse dependencies on CRAN (all via Suggests only, no strong revdeps):

| package              | result                                               |
|:---------------------|:-----------------------------------------------------|
| JGR 1.9-2            | 0 new problems (old/new `00check.log` identical)     |
| psyverse 0.2.6       | 0 new problems (old/new `00check.log` identical)     |
| D4TAlink.light 2.1.23| **not checked** - see below                          |

`cran.md` is wrong: it reports "failed to check 0" and counts D4TAlink.light
as a Bioconductor package because its version could not be resolved. Correct
summary: 3 CRAN revdeps, 0 new problems, 1 failed to check.

None of the revdeps exercise XLConnect in their examples/tests/vignettes, so
their XLConnect calls were also run by hand against both versions:
D4TAlink.light (write path), JGR (`readWorksheet`), psyverse
(`writeWorksheetToFile`) all OK with both. psyverse's
`read_spreadsheet(xlsxPkg = "XLConnect")` fails identically with both versions
(psyverse bug: passes an XLConnect workbook to `openxlsx::readWorkbook`).

## Open item: D4TAlink.light

Could not be installed in the session's environment:

- It imports Bioconductor's `Biobase`; bioconductor.org was blocked by the
  network policy (403). It has been requested on the allow list - check with
  `curl -sS -o /dev/null -w '%{http_code}' https://bioconductor.org/config.yaml`.
- Some CRAN deps need system libraries: harfbuzz/fribidi (textshaping, ragg,
  rvg, officer, officedown) and Arrow C++ (arrow, feather). Ubuntu apt works:
  e.g. `apt-get install -y libharfbuzz-dev libfribidi-dev libfreetype6-dev
  libpng-dev libtiff5-dev libjpeg-dev` (arrow can be installed with
  `NOT_CRAN=true` / `LIBARROW_BINARY=true` or from apt `libarrow-dev` if available).

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

Once bioconductor.org is reachable, `revdepcheck::revdep_add(".", "D4TAlink.light")`
(or `revdep_reset()` + full rerun) should pick up the remaining package.
