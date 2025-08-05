# ExcelFunctions
Excel LAMBDAs (Custom Functions)

I'll list some of my LAMBDAs that may or may not be useful to other people. Nothing ground breaking.

Refer to [Microsoft's LAMBDA documentation](https://support.microsoft.com/en-au/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67) for a full explanation of LAMBDA functions.

TLDR: Add the functions to the Name Manager (Ribbon > Formulas > Defined Names > Name Manager, or Ctrl + F3).

**Example**
```
Name: IsNull
Refers to: =LAMBDA(VALUE,LEN(VALUE)=0)
```

## CleanFileName(FILENAME)
Strips illegal characters from a filename, as documented in [Microsoft's Naming Files, Paths, and Namespaces](https://learn.microsoft.com/en-us/windows/win32/fileio/naming-a-file).
It expects a file name *only*, this will strip slashes, so concatenate the path later.
```
=LAMBDA(FILENAME,REDUCE(FILENAME,HSTACK(CHAR(SEQUENCE(,30,,1)),{"<",">",":","""","/","\","|","?","*"}),LAMBDA(NAME,CHAR,SUBSTITUTE(NAME,CHAR,""))))
```
Examples:
```
CleanFileName("Bad> *File* Name!")
Return value: Bad File Name!

```

## GetDigits(TEXT)
Strips all non-numeric characters from text
```
=LAMBDA(VALUE,REDUCE(VALUE,CHAR(HSTACK(SEQUENCE(,47,1),SEQUENCE(,127-58,58))),LAMBDA(TXT,CHR,SUBSTITUTE(TXT,CHR,""))))
```
Examples:
```
=GetDigits("Invoice 2025-01-01.xlsx")
Return value: 20250101
```

## GetFirst(ARRAY)
Returns first non-zero length value in ARRAY.
Expects data as ROWS. If number of COLS > 1, it assumes data is COLS and will transpose.
If you need to pass discrete values, use VSTACK().
```
=LAMBDA(ARRAY,LET(ARRAY_ROWS,IF(COLUMNS(ARRAY)>1,TRANSPOSE(ARRAY),ARRAY),INDEX(ARRAY_ROWS,XMATCH(TRUE,BYROW(ARRAY_ROWS,LAMBDA(VALUE,LEN(VALUE)>0))))))
```
Examples:
```
=GetFirst({"","","Hello","World"})
Return value: Hello
=GetFirst(VSTACK(A1,A3,A5,B2,B4))
```

## InArray(ARRAY, VALUE)
Checks if ARRAY contains VALUE.
```
=LAMBDA(VALUE,ARRAY,ISNUMBER(MATCH(VALUE,ARRAY,0)))
```
Examples:
```
=InArray({1,2,3,4,5}, 3)
Return value: TRUE

=InArray({1,2,3,4,5}, 10)
Return value: FALSE
```

## IsNull(VALUE)
Checks if VALUE has zero length.
```
=LAMBDA(VALUE,LEN(VALUE)=0)
```
Examples:
```
=IsNull("Hello")
Return value: FALSE

=IsNull("")
Return value: TRUE

=IsNull(0)
Return value: FALSE
```

## NetBookValue(AcquisitionValue,CapitalisationDate,UsefulLife,NBVDate)
Calculates Net Book Value, assuming straight line depreciation.
```
=LAMBDA(AcquisitionValue,CapitalisationDate,UsefulLife,NBVDate,MAX(0,AcquisitionValue-(NBVDate-CapitalisationDate)/365*(AcquisitionValue/UsefulLife)))
```
