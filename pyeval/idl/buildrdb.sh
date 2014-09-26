# use this on linux with ooo build env
export INPATH=/usr/lib64/libreoffice/sdk
idlc -I $SOLARVER/$INPATH/idl XPyEval.idl
regmerge ../rdb/sample.rdb UCR  XPyEval.urd

#rem use this on windows with ooo build env
#rem guw.pl idlc -I $SOLARVER/$INPATH/idl XPyEval.idl
#rem guw.pl regmerge ../rdb/sample.rdb UCR  XPyEval.urd
