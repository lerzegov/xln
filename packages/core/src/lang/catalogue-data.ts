// The Excel function catalogue, as a table that is easy to edit and to review in a diff.
// It is safety-critical (brief §5, probe F6): a modern function written without its
// stored prefix becomes #NAME? and Excel re-saves it as `_xludf.` permanently. The
// compiler refuses any function that is not listed here, so a gap fails loudly.
//
// One function per line, columns separated by spaces:
//
//   NAME        the English name, upper case, as Excel displays it
//   PREFIX      how the file stores it:  -     bare (Excel 2007 or earlier)
//                                        xlfn  _xlfn.NAME
//                                        xlws  _xlfn._xlws.NAME
//   MIN MAX     arity; `*` means variadic, up to Excel's 255 arguments
//   SINCE       first Excel version with the function: 2007 (or earlier), 2010, 2013,
//               2016, 2019, 2021, 2024, or 365 (subscription only so far)
//   FLAGS       optional, comma-separated:
//                 ?prefix   the stored prefix is not confirmed (from memory, not from a
//                           saved file or Microsoft's [MS-XLSX] list); the compiler warns
//                 internal  written by Excel in the file, not typed by users
//                           (`@` is stored as SINGLE, `x#` as ANCHORARRAY, the trim
//                           references `A1.:.A10` `A1:.A10` `A1.:A10` as _TRO_ALL,
//                           _TRO_TRAILING, _TRO_LEADING: probe F10)
//
// Sources: `xlformula.py` (XLFN, XLWS, LEGACY, ETA sets; prefixes confirmed in saved
// workbooks), probes F6 and T06, and Microsoft's function reference from memory.
// Lines starting with `#` are comments.

export const CATALOGUE_TABLE = `
# ---- Excel 2007 and earlier: stored bare ---------------------------------------------
ABS                 -     1 1   2007
ACCRINT             -     6 8   2007
ACCRINTM            -     4 5   2007
ACOS                -     1 1   2007
ACOSH               -     1 1   2007
ADDRESS             -     2 5   2007
AMORDEGRC           -     6 7   2007
AMORLINC            -     6 7   2007
AND                 -     1 *   2007
AREAS               -     1 1   2007
ASC                 -     1 1   2007
ASIN                -     1 1   2007
ASINH               -     1 1   2007
ATAN                -     1 1   2007
ATAN2               -     2 2   2007
ATANH               -     1 1   2007
AVEDEV              -     1 *   2007
AVERAGE             -     1 *   2007
AVERAGEA            -     1 *   2007
AVERAGEIF           -     2 3   2007
AVERAGEIFS          -     3 *   2007
BAHTTEXT            -     1 1   2007
BESSELI             -     2 2   2007
BESSELJ             -     2 2   2007
BESSELK             -     2 2   2007
BESSELY             -     2 2   2007
BETADIST            -     3 5   2007
BETAINV             -     3 5   2007
BIN2DEC             -     1 1   2007
BIN2HEX             -     1 2   2007
BIN2OCT             -     1 2   2007
BINOMDIST           -     4 4   2007
CALL                -     1 *   2007
CEILING             -     2 2   2007
CELL                -     1 2   2007
CHAR                -     1 1   2007
CHIDIST             -     2 2   2007
CHIINV              -     2 2   2007
CHITEST             -     2 2   2007
CHOOSE              -     2 *   2007
CLEAN               -     1 1   2007
CODE                -     1 1   2007
COLUMN              -     0 1   2007
COLUMNS             -     1 1   2007
COMBIN              -     2 2   2007
COMPLEX             -     2 3   2007
CONCATENATE         -     1 *   2007
CONFIDENCE          -     3 3   2007
CONVERT             -     3 3   2007
CORREL              -     2 2   2007
COS                 -     1 1   2007
COSH                -     1 1   2007
COUNT               -     1 *   2007
COUNTA              -     1 *   2007
COUNTBLANK          -     1 1   2007
COUNTIF             -     2 2   2007
COUNTIFS            -     2 *   2007
COUPDAYBS           -     3 4   2007
COUPDAYS            -     3 4   2007
COUPDAYSNC          -     3 4   2007
COUPNCD             -     3 4   2007
COUPNUM             -     3 4   2007
COUPPCD             -     3 4   2007
COVAR               -     2 2   2007
CRITBINOM           -     3 3   2007
CUBEKPIMEMBER       -     3 4   2007
CUBEMEMBER          -     2 3   2007
CUBEMEMBERPROPERTY  -     3 3   2007
CUBERANKEDMEMBER    -     3 4   2007
CUBESET             -     2 5   2007
CUBESETCOUNT        -     1 1   2007
CUBEVALUE           -     1 *   2007
CUMIPMT             -     6 6   2007
CUMPRINC            -     6 6   2007
DATE                -     3 3   2007
DATEDIF             -     3 3   2007
DATEVALUE           -     1 1   2007
DAVERAGE            -     3 3   2007
DAY                 -     1 1   2007
DAYS360             -     2 3   2007
DB                  -     4 5   2007
DCOUNT              -     3 3   2007
DCOUNTA             -     3 3   2007
DDB                 -     4 5   2007
DEC2BIN             -     1 2   2007
DEC2HEX             -     1 2   2007
DEC2OCT             -     1 2   2007
DEGREES             -     1 1   2007
DELTA               -     1 2   2007
DEVSQ               -     1 *   2007
DGET                -     3 3   2007
DISC                -     4 5   2007
DMAX                -     3 3   2007
DMIN                -     3 3   2007
DOLLAR              -     1 2   2007
DOLLARDE            -     2 2   2007
DOLLARFR            -     2 2   2007
DPRODUCT            -     3 3   2007
DSTDEV              -     3 3   2007
DSTDEVP             -     3 3   2007
DSUM                -     3 3   2007
DURATION            -     5 6   2007
DVAR                -     3 3   2007
DVARP               -     3 3   2007
EDATE               -     2 2   2007
EFFECT              -     2 2   2007
EOMONTH             -     2 2   2007
ERF                 -     1 2   2007
ERFC                -     1 1   2007
ERROR.TYPE          -     1 1   2007
EVEN                -     1 1   2007
EXACT               -     2 2   2007
EXP                 -     1 1   2007
EXPONDIST           -     3 3   2007
FACT                -     1 1   2007
FACTDOUBLE          -     1 1   2007
FALSE               -     0 0   2007
FDIST               -     3 3   2007
FIND                -     2 3   2007
FINDB               -     2 3   2007
FINV                -     3 3   2007
FISHER              -     1 1   2007
FISHERINV           -     1 1   2007
FIXED               -     1 3   2007
FLOOR               -     2 2   2007
FORECAST            -     3 3   2007
FREQUENCY           -     2 2   2007
FTEST               -     2 2   2007
FV                  -     3 5   2007
FVSCHEDULE          -     2 2   2007
GAMMADIST           -     4 4   2007
GAMMAINV            -     3 3   2007
GAMMALN             -     1 1   2007
GCD                 -     1 *   2007
GEOMEAN             -     1 *   2007
GESTEP              -     1 2   2007
GETPIVOTDATA        -     2 *   2007
GROWTH              -     1 4   2007
HARMEAN             -     1 *   2007
HEX2BIN             -     1 2   2007
HEX2DEC             -     1 1   2007
HEX2OCT             -     1 2   2007
HLOOKUP             -     3 4   2007
HOUR                -     1 1   2007
HYPERLINK           -     1 2   2007
HYPGEOMDIST         -     4 4   2007
IF                  -     2 3   2007
IFERROR             -     2 2   2007
IMABS               -     1 1   2007
IMAGINARY           -     1 1   2007
IMARGUMENT          -     1 1   2007
IMCONJUGATE         -     1 1   2007
IMCOS               -     1 1   2007
IMDIV               -     2 2   2007
IMEXP               -     1 1   2007
IMLN                -     1 1   2007
IMLOG10             -     1 1   2007
IMLOG2              -     1 1   2007
IMPOWER             -     2 2   2007
IMPRODUCT           -     1 *   2007
IMREAL              -     1 1   2007
IMSIN               -     1 1   2007
IMSQRT              -     1 1   2007
IMSUB               -     2 2   2007
IMSUM               -     1 *   2007
INDEX               -     2 4   2007
INDIRECT            -     1 2   2007
INFO                -     1 1   2007
INT                 -     1 1   2007
INTERCEPT           -     2 2   2007
INTRATE             -     4 5   2007
IPMT                -     4 6   2007
IRR                 -     1 2   2007
ISBLANK             -     1 1   2007
ISERR               -     1 1   2007
ISERROR             -     1 1   2007
ISEVEN              -     1 1   2007
ISLOGICAL           -     1 1   2007
ISNA                -     1 1   2007
ISNONTEXT           -     1 1   2007
ISNUMBER            -     1 1   2007
ISODD               -     1 1   2007
ISPMT               -     4 4   2007
ISREF               -     1 1   2007
ISTEXT              -     1 1   2007
KURT                -     1 *   2007
LARGE               -     2 2   2007
LCM                 -     1 *   2007
LEFT                -     1 2   2007
LEFTB               -     1 2   2007
LEN                 -     1 1   2007
LENB                -     1 1   2007
LINEST              -     1 4   2007
LN                  -     1 1   2007
LOG                 -     1 2   2007
LOG10               -     1 1   2007
LOGEST              -     1 4   2007
LOGINV              -     3 3   2007
LOGNORMDIST         -     3 3   2007
LOOKUP              -     2 3   2007
LOWER               -     1 1   2007
MATCH               -     2 3   2007
MAX                 -     1 *   2007
MAXA                -     1 *   2007
MDETERM             -     1 1   2007
MDURATION           -     5 6   2007
MEDIAN              -     1 *   2007
MID                 -     3 3   2007
MIDB                -     3 3   2007
MIN                 -     1 *   2007
MINA                -     1 *   2007
MINUTE              -     1 1   2007
MINVERSE            -     1 1   2007
MIRR                -     3 3   2007
MMULT               -     2 2   2007
MOD                 -     2 2   2007
MODE                -     1 *   2007
MONTH               -     1 1   2007
MROUND              -     2 2   2007
MULTINOMIAL         -     1 *   2007
N                   -     1 1   2007
NA                  -     0 0   2007
NEGBINOMDIST        -     3 3   2007
NETWORKDAYS         -     2 3   2007
NOMINAL             -     2 2   2007
NORMDIST            -     4 4   2007
NORMINV             -     3 3   2007
NORMSDIST           -     1 1   2007
NORMSINV            -     1 1   2007
NOT                 -     1 1   2007
NOW                 -     0 0   2007
NPER                -     3 5   2007
NPV                 -     2 *   2007
OCT2BIN             -     1 2   2007
OCT2DEC             -     1 1   2007
OCT2HEX             -     1 2   2007
ODD                 -     1 1   2007
ODDFPRICE           -     8 9   2007
ODDFYIELD           -     8 9   2007
ODDLPRICE           -     7 8   2007
ODDLYIELD           -     7 8   2007
OFFSET              -     3 5   2007
OR                  -     1 *   2007
PEARSON             -     2 2   2007
PERCENTILE          -     2 2   2007
PERCENTRANK         -     2 3   2007
PERMUT              -     2 2   2007
PHONETIC            -     1 1   2007
PI                  -     0 0   2007
PMT                 -     3 5   2007
POISSON             -     3 3   2007
POWER               -     2 2   2007
PPMT                -     4 6   2007
PRICE               -     6 7   2007
PRICEDISC           -     4 5   2007
PRICEMAT            -     5 6   2007
PROB                -     3 4   2007
PRODUCT             -     1 *   2007
PROPER              -     1 1   2007
PV                  -     3 5   2007
QUARTILE            -     2 2   2007
QUOTIENT            -     2 2   2007
RADIANS             -     1 1   2007
RAND                -     0 0   2007
RANDBETWEEN         -     2 2   2007
RANK                -     2 3   2007
RATE                -     3 6   2007
RECEIVED            -     4 5   2007
REGISTER.ID         -     2 3   2007
REPLACE             -     4 4   2007
REPLACEB            -     4 4   2007
REPT                -     2 2   2007
RIGHT               -     1 2   2007
RIGHTB              -     1 2   2007
ROMAN               -     1 2   2007
ROUND               -     2 2   2007
ROUNDDOWN           -     2 2   2007
ROUNDUP             -     2 2   2007
ROW                 -     0 1   2007
ROWS                -     1 1   2007
RSQ                 -     2 2   2007
RTD                 -     3 *   2007
SEARCH              -     2 3   2007
SEARCHB             -     2 3   2007
SECOND              -     1 1   2007
SERIESSUM           -     4 4   2007
SIGN                -     1 1   2007
SIN                 -     1 1   2007
SINH                -     1 1   2007
SKEW                -     1 *   2007
SLN                 -     3 3   2007
SLOPE               -     2 2   2007
SMALL               -     2 2   2007
SQRT                -     1 1   2007
SQRTPI              -     1 1   2007
STANDARDIZE         -     3 3   2007
STDEV               -     1 *   2007
STDEVA              -     1 *   2007
STDEVP              -     1 *   2007
STDEVPA             -     1 *   2007
STEYX               -     2 2   2007
SUBSTITUTE          -     3 4   2007
SUBTOTAL            -     2 *   2007
SUM                 -     1 *   2007
SUMIF               -     2 3   2007
SUMIFS              -     3 *   2007
SUMPRODUCT          -     1 *   2007
SUMSQ               -     1 *   2007
SUMX2MY2            -     2 2   2007
SUMX2PY2            -     2 2   2007
SUMXMY2             -     2 2   2007
SYD                 -     4 4   2007
T                   -     1 1   2007
TAN                 -     1 1   2007
TANH                -     1 1   2007
TBILLEQ             -     3 3   2007
TBILLPRICE          -     3 3   2007
TBILLYIELD          -     3 3   2007
TDIST               -     3 3   2007
TEXT                -     2 2   2007
TIME                -     3 3   2007
TIMEVALUE           -     1 1   2007
TINV                -     2 2   2007
TODAY               -     0 0   2007
TRANSPOSE           -     1 1   2007
TREND               -     1 4   2007
TRIM                -     1 1   2007
TRIMMEAN            -     2 2   2007
TRUE                -     0 0   2007
TRUNC               -     1 2   2007
TTEST               -     4 4   2007
TYPE                -     1 1   2007
UPPER               -     1 1   2007
VALUE               -     1 1   2007
VAR                 -     1 *   2007
VARA                -     1 *   2007
VARP                -     1 *   2007
VARPA               -     1 *   2007
VDB                 -     5 7   2007
VLOOKUP             -     3 4   2007
WEEKDAY             -     1 2   2007
WEEKNUM             -     1 2   2007
WEIBULL             -     4 4   2007
WORKDAY             -     2 3   2007
XIRR                -     2 3   2007
XNPV                -     3 3   2007
YEAR                -     1 1   2007
YEARFRAC            -     2 3   2007
YIELD               -     6 7   2007
YIELDDISC           -     4 5   2007
YIELDMAT            -     5 6   2007
ZTEST               -     2 3   2007

# ---- Excel 2010: _xlfn. ---------------------------------------------------------------
AGGREGATE           xlfn  3 *   2010
BETA.DIST           xlfn  4 6   2010
BETA.INV            xlfn  3 5   2010
BINOM.DIST          xlfn  4 4   2010
BINOM.INV           xlfn  3 3   2010
CEILING.PRECISE     xlfn  1 2   2010
CHISQ.DIST          xlfn  3 3   2010
CHISQ.DIST.RT       xlfn  2 2   2010
CHISQ.INV           xlfn  2 2   2010
CHISQ.INV.RT        xlfn  2 2   2010
CHISQ.TEST          xlfn  2 2   2010
CONFIDENCE.NORM     xlfn  3 3   2010
CONFIDENCE.T        xlfn  3 3   2010
COVARIANCE.P        xlfn  2 2   2010
COVARIANCE.S        xlfn  2 2   2010
ECMA.CEILING        xlfn  2 2   2010  ?prefix
ERF.PRECISE         xlfn  1 1   2010
ERFC.PRECISE        xlfn  1 1   2010
EXPON.DIST          xlfn  3 3   2010
F.DIST              xlfn  4 4   2010
F.DIST.RT           xlfn  3 3   2010
F.INV               xlfn  3 3   2010
F.INV.RT            xlfn  3 3   2010
F.TEST              xlfn  2 2   2010
FLOOR.PRECISE       xlfn  1 2   2010
GAMMA.DIST          xlfn  4 4   2010
GAMMA.INV           xlfn  3 3   2010
GAMMALN.PRECISE     xlfn  1 1   2010
HYPGEOM.DIST        xlfn  5 5   2010
ISO.CEILING         xlfn  1 2   2010  ?prefix
LOGNORM.DIST        xlfn  4 4   2010
LOGNORM.INV         xlfn  3 3   2010
MODE.MULT           xlfn  1 *   2010
MODE.SNGL           xlfn  1 *   2010
NEGBINOM.DIST       xlfn  4 4   2010
NETWORKDAYS.INTL    xlfn  2 4   2010
NORM.DIST           xlfn  4 4   2010
NORM.INV            xlfn  3 3   2010
NORM.S.DIST         xlfn  2 2   2010
NORM.S.INV          xlfn  1 1   2010
PERCENTILE.EXC      xlfn  2 2   2010
PERCENTILE.INC      xlfn  2 2   2010
PERCENTRANK.EXC     xlfn  2 3   2010
PERCENTRANK.INC     xlfn  2 3   2010
POISSON.DIST        xlfn  3 3   2010
QUARTILE.EXC        xlfn  2 2   2010
QUARTILE.INC        xlfn  2 2   2010
RANK.AVG            xlfn  2 3   2010
RANK.EQ             xlfn  2 3   2010
STDEV.P             xlfn  1 *   2010
STDEV.S             xlfn  1 *   2010
T.DIST              xlfn  3 3   2010
T.DIST.2T           xlfn  2 2   2010
T.DIST.RT           xlfn  2 2   2010
T.INV               xlfn  2 2   2010
T.INV.2T            xlfn  2 2   2010
T.TEST              xlfn  4 4   2010
VAR.P               xlfn  1 *   2010
VAR.S               xlfn  1 *   2010
WEIBULL.DIST        xlfn  4 4   2010
WORKDAY.INTL        xlfn  2 4   2010
Z.TEST              xlfn  2 3   2010

# ---- Excel 2013: _xlfn. ---------------------------------------------------------------
ACOT                xlfn  1 1   2013
ACOTH               xlfn  1 1   2013
ARABIC              xlfn  1 1   2013
BASE                xlfn  2 3   2013
BINOM.DIST.RANGE    xlfn  3 4   2013
BITAND              xlfn  2 2   2013
BITLSHIFT           xlfn  2 2   2013
BITOR               xlfn  2 2   2013
BITRSHIFT           xlfn  2 2   2013
BITXOR              xlfn  2 2   2013
CEILING.MATH        xlfn  1 3   2013
COMBINA             xlfn  2 2   2013
COT                 xlfn  1 1   2013
COTH                xlfn  1 1   2013
CSC                 xlfn  1 1   2013
CSCH                xlfn  1 1   2013
DAYS                xlfn  2 2   2013
DBCS                xlfn  1 1   2013  ?prefix
DECIMAL             xlfn  2 2   2013
ENCODEURL           xlfn  1 1   2013
FILTERXML           xlfn  2 2   2013
FLOOR.MATH          xlfn  1 3   2013
FORMULATEXT         xlfn  1 1   2013
GAMMA               xlfn  1 1   2013
GAUSS               xlfn  1 1   2013
IFNA                xlfn  2 2   2013
IMCOSH              xlfn  1 1   2013
IMCOT               xlfn  1 1   2013
IMCSC               xlfn  1 1   2013
IMCSCH              xlfn  1 1   2013
IMSEC               xlfn  1 1   2013
IMSECH              xlfn  1 1   2013
IMSINH              xlfn  1 1   2013
IMTAN               xlfn  1 1   2013
ISFORMULA           xlfn  1 1   2013
ISOWEEKNUM          xlfn  1 1   2013
MUNIT               xlfn  1 1   2013
NUMBERVALUE         xlfn  1 3   2013
PDURATION           xlfn  3 3   2013
PERMUTATIONA        xlfn  2 2   2013
PHI                 xlfn  1 1   2013
RRI                 xlfn  3 3   2013
SEC                 xlfn  1 1   2013
SECH                xlfn  1 1   2013
SHEET               xlfn  0 1   2013
SHEETS              xlfn  0 1   2013
SKEW.P              xlfn  1 *   2013
UNICHAR             xlfn  1 1   2013
UNICODE             xlfn  1 1   2013
WEBSERVICE          xlfn  1 1   2013
XOR                 xlfn  1 *   2013

# ---- Excel 2016: _xlfn. ---------------------------------------------------------------
FORECAST.ETS            xlfn  3 6   2016
FORECAST.ETS.CONFINT    xlfn  3 7   2016
FORECAST.ETS.SEASONALITY xlfn 2 4   2016
FORECAST.ETS.STAT       xlfn  3 6   2016
FORECAST.LINEAR         xlfn  3 3   2016

# ---- Excel 2019: _xlfn. ---------------------------------------------------------------
CONCAT              xlfn  1 *   2019
IFS                 xlfn  2 *   2019
MAXIFS              xlfn  3 *   2019
MINIFS              xlfn  3 *   2019
SWITCH              xlfn  3 *   2019
TEXTJOIN            xlfn  3 *   2019

# ---- Excel 2021 (dynamic arrays): _xlfn., FILTER and SORT _xlfn._xlws. ---------------
ANCHORARRAY         xlfn  1 1   2021  internal
FILTER              xlws  2 3   2021
LET                 xlfn  3 *   2021
RANDARRAY           xlfn  0 5   2021
SEQUENCE            xlfn  1 4   2021
SINGLE              xlfn  1 1   2021  internal
SORT                xlws  1 4   2021
SORTBY              xlfn  2 *   2021
UNIQUE              xlfn  1 3   2021
XLOOKUP             xlfn  3 6   2021
XMATCH              xlfn  2 4   2021

# ---- Excel 2024: _xlfn. ---------------------------------------------------------------
ARRAYTOTEXT         xlfn  1 2   2024
BYCOL               xlfn  2 2   2024
BYROW               xlfn  2 2   2024
CHOOSECOLS          xlfn  2 *   2024
CHOOSEROWS          xlfn  2 *   2024
DROP                xlfn  2 3   2024
EXPAND              xlfn  2 4   2024
HSTACK              xlfn  1 *   2024
ISOMITTED           xlfn  1 1   2024
LAMBDA              xlfn  1 *   2024
MAKEARRAY           xlfn  3 3   2024
MAP                 xlfn  2 *   2024
REDUCE              xlfn  3 3   2024
SCAN                xlfn  3 3   2024
TAKE                xlfn  2 3   2024
TEXTAFTER           xlfn  2 6   2024
TEXTBEFORE          xlfn  2 6   2024
TEXTSPLIT           xlfn  2 6   2024
TOCOL               xlfn  1 3   2024
TOROW               xlfn  1 3   2024
VALUETOTEXT         xlfn  1 2   2024
VSTACK              xlfn  1 *   2024
WRAPCOLS            xlfn  2 3   2024
WRAPROWS            xlfn  2 3   2024

# ---- Microsoft 365 only so far: _xlfn. ------------------------------------------------
COPILOT             xlfn  1 *   365   ?prefix
DETECTLANGUAGE      xlfn  1 1   365   ?prefix
FIELDVALUE          xlfn  2 2   365   ?prefix
GROUPBY             xlfn  3 8   365
IMAGE               xlfn  1 5   365   ?prefix
PERCENTOF           xlfn  2 2   365
PIVOTBY             xlfn  4 11  365
PY                  xlfn  2 2   365   ?prefix
REGEXEXTRACT        xlfn  2 4   365   ?prefix
REGEXREPLACE        xlfn  3 5   365   ?prefix
REGEXTEST           xlfn  2 3   365   ?prefix
STOCKHISTORY        xlfn  2 11  365   ?prefix
TRANSLATE           xlfn  1 3   365   ?prefix
TRIMRANGE           xlfn  1 3   365
_TRO_ALL            xlfn  1 1   365   internal
_TRO_LEADING        xlfn  1 1   365   internal
_TRO_TRAILING       xlfn  1 1   365   internal
`;
