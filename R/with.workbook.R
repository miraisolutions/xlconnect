# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

with.workbook <- function(data, expr, ...) {
	env <- new.env(parent = parent.frame())
	for(name in getDefinedNames(data)) {
		tryCatch(assign(name, readNamedRegion(data, name = name, ...), env = env),
			error = function(e) {
				warning(e)
			}
		)
	}
	eval(substitute(expr), env = env)
}
