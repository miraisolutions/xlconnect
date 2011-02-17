\name{XLC}
\alias{XLC}
\title{
XLConnect Constants
}
\description{
List structure defining several constants used across \pkg{XLConnect}.
}
\format{
  The format is:
  \preformatted{
List of 82
 $ STYLE_ACTION.XLCONNECT     : chr "XLCONNECT"
 $ STYLE_ACTION.NONE          : chr "NONE"
 $ STYLE_ACTION.PREDEFINED    : chr "PREDEFINED"
 $ STYLE_ACTION.NAME_PREFIX   : chr "STYLE_NAME_PREFIX"
 $ BORDER.DASHED              : num 3
 $ BORDER.DASH_DOT            : num 9
 $ BORDER.DASH_DOT_DOT        : num 11
 $ BORDER.DOTTED              : num 7
 $ BORDER.DOUBLE              : num 6
 $ BORDER.HAIR                : num 4
 $ BORDER.MEDIUM              : num 2
 $ BORDER.MEDIUM_DASHED       : num 8
 $ BORDER.MEDIUM_DASH_DOT     : num 10
 $ BORDER.MEDIUM_DASH_DOT_DOT : num 12
 $ BORDER.NONE                : num 0
 $ BORDER.SLANTED_DASH_DOT    : num 13
 $ BORDER.THICK               : num 5
 $ BORDER.THIN                : num 1
 $ COLOR.BLACK                : num 8
 $ COLOR.WHITE                : num 9
 $ COLOR.RED                  : num 10
 $ COLOR.BRIGHT_GREEN         : num 11
 $ COLOR.BLUE                 : num 12
 $ COLOR.YELLOW               : num 13
 $ COLOR.PINK                 : num 14
 $ COLOR.TURQUOISE            : num 15
 $ COLOR.DARK_RED             : num 16
 $ COLOR.GREEN                : num 17
 $ COLOR.DARK_BLUE            : num 18
 $ COLOR.DARK_YELLOW          : num 19
 $ COLOR.VIOLET               : num 20
 $ COLOR.TEAL                 : num 21
 $ COLOR.GREY_25_PERCENT      : num 22
 $ COLOR.GREY_50_PERCENT      : num 23
 $ COLOR.CORNFLOWER_BLUE      : num 24
 $ COLOR.MAROON               : num 25
 $ COLOR.LEMON_CHIFFON        : num 26
 $ COLOR.ORCHID               : num 28
 $ COLOR.CORAL                : num 29
 $ COLOR.ROYAL_BLUE           : num 30
 $ COLOR.LIGHT_CORNFLOWER_BLUE: num 31
 $ COLOR.SKY_BLUE             : num 40
 $ COLOR.LIGHT_TURQUOISE      : num 41
 $ COLOR.LIGHT_GREEN          : num 42
 $ COLOR.LIGHT_YELLOW         : num 43
 $ COLOR.PALE_BLUE            : num 44
 $ COLOR.ROSE                 : num 45
 $ COLOR.LAVENDER             : num 46
 $ COLOR.TAN                  : num 47
 $ COLOR.LIGHT_BLUE           : num 48
 $ COLOR.AQUA                 : num 49
 $ COLOR.LIME                 : num 50
 $ COLOR.GOLD                 : num 51
 $ COLOR.LIGHT_ORANGE         : num 52
 $ COLOR.ORANGE               : num 53
 $ COLOR.BLUE_GREY            : num 54
 $ COLOR.GREY_40_PERCENT      : num 55
 $ COLOR.DARK_TEAL            : num 56
 $ COLOR.SEA_GREEN            : num 57
 $ COLOR.DARK_GREEN           : num 58
 $ COLOR.OLIVE_GREEN          : num 59
 $ COLOR.BROWN                : num 60
 $ COLOR.PLUM                 : num 61
 $ COLOR.INDIGO               : num 62
 $ COLOR.GREY_80_PERCENT      : num 63
 $ COLOR.AUTOMATIC            : num 64
 $ FILL.NO_FILL               : num 0
 $ FILL.SOLID_FOREGROUND      : num 1
 $ FILL.FINE_DOTS             : num 2
 $ FILL.ALT_BARS              : num 3
 $ FILL.SPARSE_DOTS           : num 4
 $ FILL.THICK_HORZ_BANDS      : num 5
 $ FILL.THICK_VERT_BANDS      : num 6
 $ FILL.THICK_BACKWARD_DIAG   : num 7
 $ FILL.THICK_FORWARD_DIAG    : num 8
 $ FILL.BIG_SPOTS             : num 9
 $ FILL.BRICKS                : num 10
 $ FILL.THIN_HORZ_BANDS       : num 11
 $ FILL.THIN_VERT_BANDS       : num 12
 $ FILL.THIN_BACKWARD_DIAG    : num 13
 $ FILL.THIN_FORWARD_DIAG     : num 14
 $ FILL.SQUARES               : num 15
 $ FILL.DIAMONDS              : num 16
 }
}
\author{
Martin Studer\cr
Mirai Solutions GmbH \url{http://www.mirai-solutions.com}
}
\details{
The \code{XLC} list structure defines several constants used throughout \pkg{XLConnect}. The general convention for
enumeration types is to address corresponding constants via \var{XLC$"<ENUM_TYPE>.<VALUE>"} where \var{<ENUM_TYPE>} specifies
the name of the enumeration and \var{<VALUE>} specifies a corresponding enumeration value. An example is \code{XLC$"COLOR.RED"} where
"COLOR" is the enumeration type and "RED" is the corresponding color enumeration value.
}
\keyword{datasets}
\keyword{list}
\keyword{utilities}
