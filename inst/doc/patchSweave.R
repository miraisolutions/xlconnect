if( as.integer(R.Version()[['svn rev']]) <= 52157 ){
  # R 2.11.1 or earlier:
  patchedCodeRunner <- function(evalFunc=RweaveEvalWithOpt)
  {
      ## Return a function suitable as the 'runcode' element
      ## of an Sweave driver.  evalFunc will be used for the
      ## actual evaluation of chunk code.
      RweaveLatexRuncode <- function(object, chunk, options)
        {
            if(!(options$engine %in% c("R", "S"))){
                return(object)
            }

            if(!object$quiet){
                cat(formatC(options$chunknr, width=2), ":")
                if(options$echo) cat(" echo")
                if(options$keep.source) cat(" keep.source")
                if(options$eval){
                    if(options$print) cat(" print")
                    if(options$term) cat(" term")
                    cat("", options$results)
                    if(options$fig){
                        if(options$eps) cat(" eps")
                        if(options$pdf) cat(" pdf")
                    }
                }
                if(!is.null(options$label))
                  cat(" (label=", options$label, ")", sep="")
                cat("\n")
            }

            chunkprefix <- RweaveChunkPrefix(options)

            if(options$split){
                ## [x][[1L]] avoids partial matching of x
                chunkout <- object$chunkout[chunkprefix][[1L]]
                if(is.null(chunkout)){
                    chunkout <- file(paste(chunkprefix, "tex", sep="."), "w")
                    if(!is.null(options$label))
                      object$chunkout[[chunkprefix]] <- chunkout
                }
            }
            else
              chunkout <- object$output
      
      # Pull in non-standard opts.
      options <- c( options, options('wrapinput','wrapoutput') )	

      saveopts <- options(keep.source=options$keep.source)
      on.exit(options(saveopts))

            SweaveHooks(options, run=TRUE)

            chunkexps <- try(parse(text=chunk), silent=TRUE)
            RweaveTryStop(chunkexps, options)
            openSinput <- FALSE
            openSchunk <- FALSE

            if(length(chunkexps) == 0L)
              return(object)

            srclines <- attr(chunk, "srclines")
            linesout <- integer(0L)
            srcline <- srclines[1L]

      srcrefs <- attr(chunkexps, "srcref")
      if (options$expand)
        lastshown <- 0L
      else
        lastshown <- srcline - 1L
      thisline <- 0
            for(nce in 1L:length(chunkexps))
              {
                  ce <- chunkexps[[nce]]
                  if (nce <= length(srcrefs) && !is.null(srcref <- srcrefs[[nce]])) {
                      if (options$expand) {
                    srcfile <- attr(srcref, "srcfile")
                    showfrom <- srcref[1L]
                    showto <- srcref[3L]
                      } else {
                        srcfile <- object$srcfile
                        showfrom <- srclines[srcref[1L]]
                        showto <- srclines[srcref[3L]]
                      }
                      dce <- getSrcLines(srcfile, lastshown+1, showto)
              leading <- showfrom-lastshown
              lastshown <- showto
                      srcline <- srclines[srcref[3L]]
                      while (length(dce) && length(grep("^[[:blank:]]*$", dce[1L]))) {
            dce <- dce[-1L]
            leading <- leading - 1L
              }
          } else {
                      dce <- deparse(ce, width.cutoff=0.75*getOption("width"))
                      leading <- 1L
                  }
                  if(object$debug)
                    cat("\nRnw> ", paste(dce, collapse="\n+  "),"\n")
                  if(options$echo && length(dce)){
                      if(!openSinput){
                          if(!openSchunk){
                               cat("\\begin{tikzCodeBlock}[listing style=sweavechunk]\n",
                                  file=chunkout, append=TRUE)
                              linesout[thisline + 1] <- srcline
                              thisline <- thisline + 1
                              openSchunk <- TRUE
                              firstChunkLine <- TRUE
                          }
                          if( !is.null(options$wrapinput) ){
                            cat( options$wrapinput[1],
                              file=chunkout, append=TRUE)
                          }
                          openSinput <- TRUE
                      }
          if( firstChunkLine ) {
            cat(paste(dce[1L:leading], sep="", collapse="\n"),
              file=chunkout, append=TRUE, sep="")
            firstChunkLine <- FALSE
          }else{
            cat("\n", paste(dce[1L:leading], sep="", collapse="\n"),
              file=chunkout, append=TRUE, sep="")
          }
                      if (length(dce) > leading)
                        cat("\n", paste( dce[-(1L:leading)], sep="", collapse="\n"),
                            file=chunkout, append=TRUE, sep="")
          linesout[thisline + 1L:length(dce)] <- srcline
          thisline <- thisline + length(dce)
                  }

                                          # tmpcon <- textConnection("output", "w")
                                          # avoid the limitations (and overhead) of output text connections
                  tmpcon <- file()
                  sink(file=tmpcon)
                  err <- NULL
                  if(options$eval) err <- evalFunc(ce, options)
                  cat("\n") # make sure final line is complete
                  sink()
                  output <- readLines(tmpcon)
                  close(tmpcon)
                  ## delete empty output
                  if(length(output) == 1L & output[1L] == "") output <- NULL

                  RweaveTryStop(err, options)

                  if(object$debug)
                    cat(paste(output, collapse="\n"))

                  if(length(output) & (options$results != "hide")){

                      if(openSinput){
                          if( !is.null(options$wrapinput) ){
                            cat(options$wrapinput[2], file=chunkout, append=TRUE)
                          }
                          linesout[thisline + 1L:2L] <- srcline
                          thisline <- thisline + 2L
                          openSinput <- FALSE
                      }
                      if(options$results=="verbatim"){
                          if(!openSchunk){
                               cat("\\begin{tikzCodeBlock}[listing style=sweavechunk]\n",
                                  file=chunkout, append=TRUE)
                              linesout[thisline + 1L] <- srcline
                              thisline <- thisline + 1L
                              openSchunk <- TRUE
                              firstChunkLine <- TRUE
                          }
                          if( !is.null(options$wrapoutput) ){
                            cat(options$wrapoutput[1],
                              file=chunkout, append=TRUE)
                          }
                          linesout[thisline + 1L] <- srcline
                          thisline <- thisline + 1L
                      }

                      output <- paste(output,collapse="\n")
                      if(options$strip.white %in% c("all", "true")){
                          output <- sub("^[[:space:]]*\n", "", output)
                          output <- sub("\n[[:space:]]*$", "", output)
                          if(options$strip.white=="all")
                            output <- sub("\n[[:space:]]*\n", "\n", output)
                      }
                      cat(output, file=chunkout, append=TRUE)
                      count <- sum(strsplit(output, NULL)[[1L]] == "\n")
                      if (count > 0L) {
                        linesout[thisline + 1L:count] <- srcline
                        thisline <- thisline + count
                      }

                      remove(output)

                      if(options$results=="verbatim"){
                          if( !is.null(options$wrapoutput) ){
                            cat(options$wrapoutput[2], file=chunkout, append=TRUE)
                          }
                          linesout[thisline + 1L:2] <- srcline
                          thisline <- thisline + 2L
                      }
                  }
              }

            if(openSinput){
                if( !is.null(options$wrapinput) ){
                  cat(options$wrapinput[2], file=chunkout, append=TRUE)
                }
                linesout[thisline + 1L:2L] <- srcline
                thisline <- thisline + 2L
            }

            if(openSchunk){
                cat("\\end{tikzCodeBlock}\n", file=chunkout, append=TRUE)
                linesout[thisline + 1L] <- srcline
                thisline <- thisline + 1L
            }

            if(is.null(options$label) & options$split)
              close(chunkout)

            if(options$split & options$include){
                cat("\\input{", chunkprefix, "}\n", sep="",
                  file=object$output, append=TRUE)
                linesout[thisline + 1L] <- srcline
                thisline <- thisline + 1L
            }

            if(options$fig && options$eval){
                if(options$eps){
                    grDevices::postscript(file=paste(chunkprefix, "eps", sep="."),
                                          width=options$width, height=options$height,
                                          paper="special", horizontal=FALSE)

                    err <- try({SweaveHooks(options, run=TRUE)
                                eval(chunkexps, envir=.GlobalEnv)})
                    grDevices::dev.off()
                    if(inherits(err, "try-error")) stop(err)
                }
                if(options$pdf){
                    grDevices::pdf(file=paste(chunkprefix, "pdf", sep="."),
                                   width=options$width, height=options$height,
                                   version=options$pdf.version,
                                   encoding=options$pdf.encoding)

                    err <- try({SweaveHooks(options, run=TRUE)
                                eval(chunkexps, envir=.GlobalEnv)})
                    grDevices::dev.off()
                    if(inherits(err, "try-error")) stop(err)
                }
                if(options$include) {
                    cat("\\includegraphics{", chunkprefix, "}\n", sep="",
                        file=object$output, append=TRUE)
                    linesout[thisline + 1L] <- srcline
                    thisline <- thisline + 1L
                }
            }
            object$linesout <- c(object$linesout, linesout)
            return(object)
        }
      RweaveLatexRuncode
  }

}else if(as.integer(R.Version()[['svn rev']]) <= 54585) {

  # R 2.12.0
  patchedCodeRunner <- function(evalFunc=RweaveEvalWithOpt)
  {
      ## Return a function suitable as the 'runcode' element
      ## of an Sweave driver.  evalFunc will be used for the
      ## actual evaluation of chunk code.
      RweaveLatexRuncode <- function(object, chunk, options)
        {
            if(!(options$engine %in% c("R", "S"))){
                return(object)
            }

            if(!object$quiet){
                cat(formatC(options$chunknr, width=2), ":")
                if(options$echo) cat(" echo")
                if(options$keep.source) cat(" keep.source")
                if(options$eval){
                    if(options$print) cat(" print")
                    if(options$term) cat(" term")
                    cat("", options$results)
                    if(options$fig){
                        if(options$eps) cat(" eps")
                        if(options$pdf) cat(" pdf")
                    }
                }
                if(!is.null(options$label))
                  cat(" (label=", options$label, ")", sep="")
                cat("\n")
            }

            chunkprefix <- RweaveChunkPrefix(options)

            if(options$split){
                ## [x][[1L]] avoids partial matching of x
                chunkout <- object$chunkout[chunkprefix][[1L]]
                if(is.null(chunkout)){
                    chunkout <- file(paste(chunkprefix, "tex", sep="."), "w")
                    if(!is.null(options$label))
                      object$chunkout[[chunkprefix]] <- chunkout
                }
            }
            else
              chunkout <- object$output

      srcfile <- object$srcfile
            SweaveHooks(options, run=TRUE)

            # Pull in non-standard options.
            options <- c( options, options( 'wrapinput', 'wrapoutput' ) )

            # Note that we edit the error message below, so change both
            # if you change this line:
            chunkexps <- try(parse(text=chunk, srcfile=srcfile), silent=TRUE)
            
            if (inherits(chunkexps, "try-error"))
              chunkexps[1L] <- sub(" parse(text = chunk, srcfile = srcfile) : \n ",
                                  "", chunkexps[1L], fixed = TRUE)
                                  
            RweaveTryStop(chunkexps, options)
            
            # A couple of functions used below...
            
      putSinput <- function(dce){
          if(!openSinput){
          if(!openSchunk){
              cat("\\begin{tikzCodeBlock}[listing style=sweavechunk]\n", file = chunkout, append = TRUE)
              linesout[thisline + 1L] <<- srcline
              thisline <<- thisline + 1L
              openSchunk <<- TRUE
          }
          if (!is.null(options$wrapinput)) {
            cat(options$wrapinput[1], file=chunkout, append=TRUE)
          }
          openSinput <<- TRUE
          }
          cat("\n", paste(getOption("prompt"), dce[1L:leading], sep="", collapse="\n"),
          file=chunkout, append=TRUE, sep="")
          if (length(dce) > leading)
          cat("\n", paste(getOption("continue"), dce[-(1L:leading)], sep="", collapse="\n"),
              file=chunkout, append=TRUE, sep="")
          linesout[thisline + seq_along(dce)] <<- srcline
          thisline <<- thisline + length(dce)
      }	  
      
      trySrcLines <- function(srcfile, showfrom, showto, ce) {
          lines <- try(suppressWarnings(getSrcLines(srcfile, showfrom, showto)), silent=TRUE)
          if (inherits(lines, "try-error")) {
              if (is.null(ce)) lines <- character(0)
              else lines <- deparse(ce, width.cutoff=0.75*getOption("width"))
          }
          lines
      }          

      chunkregexp <- "(.*)#from line#([[:digit:]]+)#"

            openSinput <- FALSE
            openSchunk <- FALSE

            srclines <- attr(chunk, "srclines")
            linesout <- integer(0L) # maintains concordance
            srcline <- srclines[1L] # current input line
      thisline <- 0L          # current output line
      lastshown <- srcline    # last line already displayed;
              # at this point it's the <<>>= line
      leading <- 1L		  # How many lines get the user prompt
      
      srcrefs <- attr(chunkexps, "srcref")
      
            for(nce in seq_along(chunkexps)) {
      ce <- chunkexps[[nce]]
                  if (options$keep.source && nce <= length(srcrefs) && !is.null(srcref <- srcrefs[[nce]])) {
          srcfile <- attr(srcref, "srcfile")
          showfrom <- srcref[1L]
          showto <- srcref[3L]
          refline <- srcfile$refline
                      if (is.null(refline)) {
                        if (grepl(chunkregexp, srcfile$filename)) {
                            refline <- as.integer(sub(chunkregexp, "\\2", srcfile$filename))
                            srcfile$filename <- sub(chunkregexp, "\\1", srcfile$filename)
                        } else 
                            refline <- NA
                        srcfile$refline <- refline
                      }
                      if (!options$expand && !is.na(refline)) 
                        showfrom <- showto <- refline
                        
          if (!is.na(refline) || is.na(lastshown)) { 
        # We expanded a named chunk for this expression or the previous
        # one
        dce <- trySrcLines(srcfile, showfrom, showto, ce)
        leading <- 1L
        if (!is.na(refline))
            lastshown <- NA
        else
            lastshown <- showto
          } else {
        dce <- trySrcLines(srcfile, lastshown+1L, showto, ce)
        leading <- showfrom-lastshown
        lastshown <- showto
          }
                      srcline <- showto
                      while (length(dce) && length(grep("^[[:blank:]]*$", dce[1L]))) {
            dce <- dce[-1L]
            leading <- leading - 1L
              }
          } else {
                      dce <- deparse(ce, width.cutoff=0.75*getOption("width"))
                      leading <- 1L
                  }
                  if(object$debug)
                      cat("\nRnw> ", paste(dce, collapse="\n+  "),"\n")
                      
                  if(options$echo && length(dce))
                      putSinput(dce)
                      
                    # tmpcon <- textConnection("output", "w")
                    # avoid the limitations (and overhead) of output text connections
                  tmpcon <- file()
                  sink(file=tmpcon)
                  err <- NULL
                  if(options$eval) err <- evalFunc(ce, options)
                  cat("\n") # make sure final line is complete
                  sink()
                  output <- readLines(tmpcon)
                  close(tmpcon)
                  ## delete empty output
                  if(length(output) == 1L & output[1L] == "") output <- NULL

                  RweaveTryStop(err, options)

                  if(object$debug)
                    cat(paste(output, collapse="\n"))

                  if(length(output) & (options$results != "hide")){

                      if(openSinput){
                        if(!is.null(options$wrapinput)) {
                          cat(options$wrapinput[2], file=chunkout, append=TRUE)
                        }
                          linesout[thisline + 1L:2L] <- srcline
                          thisline <- thisline + 2L
                          openSinput <- FALSE
                      }
                      if(options$results=="verbatim"){
                          if(!openSchunk){
                              cat("\\begin{tikzCodeBlock}[listing style=sweavechunk]\n", file = chunkout, append = TRUE)
                              linesout[thisline + 1L] <- srcline
                              thisline <- thisline + 1L
                              openSchunk <- TRUE
                          }
                          if (!is.null(options$wrapoutput)) {
                            cat(options$wrapoutput[1], file=chunkout, append=TRUE)
                          }
                          linesout[thisline + 1L] <- srcline
                          thisline <- thisline + 1L
                      }

                      output <- paste(output,collapse="\n")
                      if(options$strip.white %in% c("all", "true")){
                          output <- sub("^[[:space:]]*\n", "", output)
                          output <- sub("\n[[:space:]]*$", "", output)
                          if(options$strip.white=="all")
                            output <- sub("\n[[:space:]]*\n", "\n", output)
                      }
                      cat(output, file=chunkout, append=TRUE)
                      count <- sum(strsplit(output, NULL)[[1L]] == "\n")
                      if (count > 0L) {
                        linesout[thisline + 1L:count] <- srcline
                        thisline <- thisline + count
                      }

                      remove(output)

                      if(options$results=="verbatim"){
                        if(!is.null(options$wrapoutput)) {
                          cat(options$wrapoutput[2], file=chunkout, append=TRUE)
                        }
                          linesout[thisline + 1L:2L] <- srcline
                          thisline <- thisline + 2L
                      }
                  }
              }

      if(options$echo && options$keep.source 
         && !is.na(lastshown) 
         && lastshown < (showto <- srclines[length(srclines)])) {
          dce <- trySrcLines(srcfile, lastshown+1L, showto, NULL)
          leading <- length(dce) # These are all trailing comments
          putSinput(dce)
      }
      
            if(openSinput){
              if (!is.null(options$wrapinput)) {
                cat(options$wrapinput[2], file=chunkout, append=TRUE)
              }
                linesout[thisline + 1L:2L] <- srcline
                thisline <- thisline + 2L
            }

            if(openSchunk){
                cat("\\end{tikzCodeBlock}\n", file=chunkout, append=TRUE)
                linesout[thisline + 1L] <- srcline
                thisline <- thisline + 1L
            }

            if(is.null(options$label) & options$split)
              close(chunkout)

            if(options$split & options$include){
                cat("\\input{", chunkprefix, "}\n", sep="",
                  file=object$output, append=TRUE)
                linesout[thisline + 1L] <- srcline
                thisline <- thisline + 1L
            }

            if(options$fig && options$eval){
                if(options$eps){
                    grDevices::postscript(file=paste(chunkprefix, "eps", sep="."),
                                          width=options$width, height=options$height,
                                          paper="special", horizontal=FALSE)

                    err <- try({SweaveHooks(options, run=TRUE)
                                eval(chunkexps, envir=.GlobalEnv)})
                    grDevices::dev.off()
                    if(inherits(err, "try-error")) stop(err)
                }
                if(options$pdf){
                    grDevices::pdf(file=paste(chunkprefix, "pdf", sep="."),
                                   width=options$width, height=options$height,
                                   version=options$pdf.version,
                                   encoding=options$pdf.encoding)

                    err <- try({SweaveHooks(options, run=TRUE)
                                eval(chunkexps, envir=.GlobalEnv)})
                    grDevices::dev.off()
                    if(inherits(err, "try-error")) stop(err)
                }
                if(options$include) {
                    cat("\\includegraphics{", chunkprefix, "}\n", sep="",
                        file=object$output, append=TRUE)
                    linesout[thisline + 1L] <- srcline
                    thisline <- thisline + 1L
                }
            }
            object$linesout <- c(object$linesout, linesout)
            return(object)
        }
      RweaveLatexRuncode
  }

} else {

  # R 2.13.0
  patchedCodeRunner <- function(evalFunc = RweaveEvalWithOpt)
  {
      ## Return a function suitable as the 'runcode' element
      ## of an Sweave driver.  evalFunc will be used for the
      ## actual evaluation of chunk code.
      ## FIXME: well, actually not for the figures.
      ## If there were just one figure option set, we could eval the chunk
      ## only once.
      function(object, chunk, options) {
          pdf.Swd <- function(name, width, height, ...)
              grDevices::pdf(file = paste(chunkprefix, "pdf", sep = "."),
                             width = width, height = height,
                             version = options$pdf.version,
                             encoding = options$pdf.encoding)
          eps.Swd <- function(name, width, height, ...)
              grDevices::postscript(file = paste(name, "eps", sep = "."),
                                    width = width, height = height,
                                    paper = "special", horizontal = FALSE)
          png.Swd <- function(name, width, height, options, ...)
              grDevices::png(filename = paste(chunkprefix, "png", sep = "."),
                             width = width, height = height,
                             res = options$resolution, units = "in")
          jpeg.Swd <- function(name, width, height, options, ...)
              grDevices::jpeg(filename = paste(chunkprefix, "png", sep = "."),
                              width = width, height = height,
                              res = options$resolution, units = "in")

          if (!(options$engine %in% c("R", "S"))) return(object)

          if (!object$quiet) {
              cat(formatC(options$chunknr, width = 2), ":")
              if (options$echo) cat(" echo")
              if (options$keep.source) cat(" keep.source")
              if (options$eval) {
                  if (options$print) cat(" print")
                  if (options$term) cat(" term")
                  cat("", options$results)
                  if (options$fig) {
                      if (options$eps) cat(" eps")
                      if (options$pdf) cat(" pdf")
                      if (options$png) cat(" png")
                      if (options$jpeg) cat(" jpeg")
                      if (!is.null(options$grdevice)) cat("", options$grdevice)
                  }
              }
              if (!is.null(options$label))
                  cat(" (label=", options$label, ")", sep = "")
              cat("\n")
          }

          chunkprefix <- RweaveChunkPrefix(options)

          if (options$split) {
              ## [x][[1L]] avoids partial matching of x
              chunkout <- object$chunkout[chunkprefix][[1L]]
              if (is.null(chunkout)) {
                  chunkout <- file(paste(chunkprefix, "tex", sep = "."), "w")
                  if (!is.null(options$label))
                      object$chunkout[[chunkprefix]] <- chunkout
              }
          } else chunkout <- object$output

          srcfile <- srcfilecopy(object$filename, chunk)
          SweaveHooks(options, run = TRUE)

          # Pull in non-standard options.
          options <- c( options, options( 'wrapinput', 'wrapoutput' ) )

          ## Note that we edit the error message below, so change both
          ## if you change this line:
          chunkexps <- try(parse(text = chunk, srcfile = srcfile), silent = TRUE)

          if (inherits(chunkexps, "try-error"))
              chunkexps[1L] <- sub(" parse(text = chunk, srcfile = srcfile) : \n ",
                                   "", chunkexps[1L], fixed = TRUE)

          RweaveTryStop(chunkexps, options)

          ## Some worker functions used below...
          putSinput <- function(dce) {
              if (!openSinput) {
                  if (!openSchunk) {
                      cat("\\begin{tikzCodeBlock}[listing style=sweavechunk]\n", file = chunkout)
                      linesout[thisline + 1L] <<- srcline
                      thisline <<- thisline + 1L
                      openSchunk <<- TRUE
                  }
                  if (!is.null(options$wrapinput)) {
                    cat(options$wrapinput[1], file = chunkout)
                  }
                  openSinput <<- TRUE
              }
              cat("\n", paste(getOption("prompt"), dce[1L:leading],
                              sep = "", collapse = "\n"),
                  file = chunkout, sep = "")
              if (length(dce) > leading)
                  cat("\n", paste(getOption("continue"), dce[-(1L:leading)],
                                  sep = "", collapse = "\n"),
                      file = chunkout, sep = "")
              linesout[thisline + seq_along(dce)] <<- srcline
              thisline <<- thisline + length(dce)
          }

          trySrcLines <- function(srcfile, showfrom, showto, ce) {
              lines <- try(suppressWarnings(getSrcLines(srcfile, showfrom, showto)),
                           silent = TRUE)
              if (inherits(lines, "try-error")) {
                  if (is.null(ce)) lines <- character()
                  else lines <- deparse(ce, width.cutoff = 0.75*getOption("width"))
              }
              lines
          }

          echoComments <- function(showto) {
              if (options$echo && !is.na(lastshown) && lastshown < showto) {
                  dce <- trySrcLines(srcfile, lastshown + 1L, showto, NULL)
                  linedirs <- grepl("^#line ", dce)
      dce <- dce[!linedirs]
                  leading <<- length(dce) # These are all trailing comments
                  putSinput(dce)
                  lastshown <<- showto
              }
          }

          openSinput <- FALSE
          openSchunk <- FALSE

          srclines <- attr(chunk, "srclines")
          linesout <- integer()      # maintains concordance
          srcline <- srclines[1L]    # current input line
          thisline <- 0L             # current output line
          lastshown <- 0L            # last line already displayed;

          refline <- NA    # line containing the current named chunk ref
          leading <- 1L    # How many lines get the user prompt

          srcrefs <- attr(chunkexps, "srcref")

          for (nce in seq_along(chunkexps)) {
              ce <- chunkexps[[nce]]
              if (options$keep.source && nce <= length(srcrefs) &&
                  !is.null(srcref <- srcrefs[[nce]])) {
                  showfrom <- srcref[7L]
                  showto <- srcref[8L]

                  dce <- trySrcLines(srcfile, lastshown+1L, showto, ce)
                  leading <- showfrom - lastshown

                  lastshown <- showto
                  srcline <- srcref[3L]

                  linedirs <- grepl("^#line ", dce)
                  dce <- dce[!linedirs]
                  # Need to reduce leading lines if some were just removed
                  leading <- leading - sum(linedirs[seq_len(leading)])

                  while (length(dce) && length(grep("^[[:blank:]]*$", dce[1L]))) {
                      dce <- dce[-1L]
                      leading <- leading - 1L
                  }
              } else {
                  dce <- deparse(ce, width.cutoff = 0.75*getOption("width"))
                  leading <- 1L
              }
              if (object$debug)
                  cat("\nRnw> ", paste(dce, collapse = "\n+  "),"\n")

              if (options$echo && length(dce)) putSinput(dce)

              ## avoid the limitations (and overhead) of output text connections
              if (options$eval) {
                  tmpcon <- file()
                  sink(file = tmpcon)
                  err <- evalFunc(ce, options)
                  cat("\n")           # make sure final line is complete
                  sink()
                  output <- readLines(tmpcon)
                  close(tmpcon)
                  ## delete empty output
                  if (length(output) == 1L && !nzchar(output[1L])) output <- NULL
                  RweaveTryStop(err, options)
              } else output <- NULL

              ## or writeLines(output)
              if (length(output) && object$debug)
                  cat(paste(output, collapse = "\n"))

              if (length(output) && (options$results != "hide")) {
                  if (openSinput) {
                    if(!is.null(options$wrapinput)) {
                      cat(options$wrapinput[2], file = chunkout)
                    }
                      linesout[thisline + 1L:2L] <- srcline
                      thisline <- thisline + 2L
                      openSinput <- FALSE
                  }
                  if (options$results == "verbatim") {
                      if (!openSchunk) {
                          cat("\\begin{tikzCodeBlock}[listing style=sweavechunk]\n", file = chunkout)
                          linesout[thisline + 1L] <- srcline
                          thisline <- thisline + 1L
                          openSchunk <- TRUE
                      }
                      if (!is.null(options$wrapoutput)) {
                        cat(options$wrapoutput[1], file = chunkout)
                      }
                      linesout[thisline + 1L] <- srcline
                      thisline <- thisline + 1L
                  }

                  output <- paste(output, collapse = "\n")
                  if (options$strip.white %in% c("all", "true")) {
                      output <- sub("^[[:space:]]*\n", "", output)
                      output <- sub("\n[[:space:]]*$", "", output)
                      if (options$strip.white == "all")
                          output <- sub("\n[[:space:]]*\n", "\n", output)
                  }
                  cat(output, file = chunkout)
                  count <- sum(strsplit(output, NULL)[[1L]] == "\n")
                  if (count > 0L) {
                      linesout[thisline + 1L:count] <- srcline
                      thisline <- thisline + count
                  }

                  remove(output)

                  if (options$results == "verbatim") {
                    if(!is.null(options$wrapoutput)) {
                      cat(options$wrapoutput[2], file = chunkout)
                    }
                      linesout[thisline + 1L:2L] <- srcline
                      thisline <- thisline + 2L
                  }
              }
          } # end of loop over chunkexps.

          ## Echo remaining comments if necessary
          if (options$keep.source) echoComments(length(srcfile$lines))

          if (openSinput) {
            if(!is.null(options$wrapinput)) {
              cat(options$wrapinput[2], file = chunkout)
            }
              linesout[thisline + 1L:2L] <- srcline
              thisline <- thisline + 2L
          }

          if (openSchunk) {
              cat("\\end{tikzCodeBlock}\n", file = chunkout)
              linesout[thisline + 1L] <- srcline
              thisline <- thisline + 1L
          }

          if (is.null(options$label) && options$split) close(chunkout)

          if (options$split && options$include) {
              cat("\\input{", chunkprefix, "}\n", sep = "",
                  file = object$output)
              linesout[thisline + 1L] <- srcline
              thisline <- thisline + 1L
          }

          if (options$fig && options$eval) {
              devs <- list()
              if (options$pdf) devs <- c(devs, list(pdf.Swd))
              if (options$eps) devs <- c(devs, list(eps.Swd))
              if (options$png) devs <- c(devs, list(png.Swd))
              if (options$jpeg) devs <- c(devs, list(jpeg.Swd))
              if (!is.null(grd <- options$grdevice))
                  devs <- c(devs, list(get(grd, envir = .GlobalEnv)))
              for (dev in devs) {
                  dev(name = chunkprefix, width = options$width,
                      height = options$height, options)
                  err <- tryCatch({
                      SweaveHooks(options, run = TRUE)
                      eval(chunkexps, envir = .GlobalEnv)
                  }, error = function(e) {
                      grDevices::dev.off()
                      stop(conditionMessage(e), call. = FALSE, domain = NA)
                  })
                  grDevices::dev.off()
              }

              if (options$include) {
                  cat("\\includegraphics{", chunkprefix, "}\n", sep = "",
                      file = object$output)
                  linesout[thisline + 1L] <- srcline
                  thisline <- thisline + 1L
              }
          }
          object$linesout <- c(object$linesout, linesout)
          object
      }
  }

} # End big hairy if statement


# We need some extra options for controlling
# Sweave output. Therefore we're going to bust
# out and overwrite some core code with more
# flexible alternatives.

# *** Begin Mother of All R Hacks ****

# First, we find out where Sweave lives.
env <- as.environment('package:utils')

# And we break into it's house.
unlockBinding('makeRweaveLatexCodeRunner',env)

# And rearrange it's furniture....
assignInNamespace('makeRweaveLatexCodeRunner',patchedCodeRunner,ns='utils')

# ...and lock the door on the way out.
lockBinding('makeRweaveLatexCodeRunner',env)

#  But we take something with us when we leave...
newCodeRunner <- utils:::makeRweaveLatexCodeRunner()

# ...to stalk it down on the street...
env <- grep('driver', lapply(sys.frames(),
  function(frame){ ls(envir=frame) }
))

# Fix for R 2.13.0. Starting with this release, R parses all code chunks in the
# Rnw file before running Sweave. The parsing step fails because Sweave is not
# running and the 'driver' object cannot be located. This is a quick hack to
# skip the driver replacement code if 'driver' was not found.
if(length(env)){

  env <- sys.frames()[[env]]

  # ...and give it a facelift.
  evalq( driver$runcode <- newCodeRunner, envir=env)

}

# *** End Mother of All R Hacks ***
