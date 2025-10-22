# ExcelFunctions
Excel LAMBDAs (Custom Functions)

I'll list some of my LAMBDAs that may or may not be useful to other people. Typically trivial, nothing ground breaking.

Refer to Microsoft's LAMBDA [announcement](https://techcommunity.microsoft.com/blog/excelblog/announcing-lambda-helper-functions-lambdas-as-arguments-and-more/2576648) and [documentation](https://support.microsoft.com/en-au/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67) for a full explanation of LAMBDA functions.

TLDR: Add the functions to the Name Manager (Ribbon > Formulas > Defined Names > Name Manager, or Ctrl + F3).

**Example**
```
Name: IsNull
Refers to: =LAMBDA(VALUE,LEN(VALUE)=0)
```

## Between(VALUE,LOWER,UPPER)
Checks if numeric value is between an UPPER and LOWER bound.
```
=LAMBDA(VALUE,LOWER,UPPER,AND(VALUE>=LOWER,VALUE<=UPPER))
```
Examples:
```
=Between(10,5,15)
Return value: TRUE

=Between(10,20,30)
Return value: FALSE
```

## CleanFileName(FILENAME)
Strips illegal characters from a filename, as documented in [Microsoft's Naming Files, Paths, and Namespaces](https://learn.microsoft.com/en-us/windows/win32/fileio/naming-a-file).
It expects a file name *only*, this will strip slashes, so concatenate the path later.
```
=LAMBDA(FILENAME,REDUCE(FILENAME,HSTACK(CHAR(SEQUENCE(,30,,1)),{"<",">",":","""","/","\","|","?","*"}),LAMBDA(NAME,CHAR,SUBSTITUTE(NAME,CHAR,""))))
```
Examples:
```
=CleanFileName("Bad> *File* Name!")
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
If you need to pass discrete values, use VSTACK().
```
=LAMBDA(ARRAY,CHOOSEROWS(TOCOL(ARRAY,1),1))
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

## Pad(Value,Width,[Cut])
Makes a value fixed width by padding with trailing spaces.

If the string length of Value exceeds Width, and optional Cut parameter is TRUE, only the first WIDTH characters will be retuned. Otherwise, an NA error will be thrown.
```
=LAMBDA(Value,Width,[Cut],
    IF(LEN(Value)>Width,
        IF(Cut,
            LEFT(Value,Width),
            NA()
        ),
        Value&REPT(" ",Width-LEN(Value))
    )
)
```
Note:

[TEXTJOIN](https://support.microsoft.com/en-au/office/textjoin-function-357b449a-ec91-49d0-80c3-0e8fc845691c) and [CONCAT](https://support.microsoft.com/en-au/office/concat-function-9b1a9a3f-94ff-41af-9736-694cbd6b4ca2) do not play nice with [SPILL ARRAYS](https://support.microsoft.com/en-au/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531), hence the use of the &amp; (ampersand) concatenation operator in this instance.

Example:
```
(Spaces shown as # for sake of clarity)
=Pad("Hello",10)
Return value: Hello#####
=Pad("Hello",5)
Return value: Hello
=Pad("Hello",2)
Return value: #N/A
=Pad("Hello",2,TRUE)
Return value: He
```

## SimpleDecipher(OffsetStart,OffsetEnd,EncodedText)
Enumerates possible simple offset cipher solutions for a given offset range. OffsetStart and OffsetEnd can be either positive or negative.  If descending order of offset is desired, put the larger value in OffsetStart and smaller in OffsetEnd.
```
=LAMBDA(OffsetStart,OffsetEnd,EncodedText,
    BYROW(
        SEQUENCE(MAX(OffsetStart,OffsetEnd)-MIN(OffsetStart,OffsetEnd)+1,1,OffsetStart,SIGN(OffsetEnd-OffsetStart)),
        LAMBDA(OFFSET,CONCAT(MAP(MID(EncodedText,SEQUENCE(LEN(EncodedText)),1),LAMBDA(CHR,CHAR(CODE(CHR)+OFFSET)))))
    )
)
```

Examples:
```
=SimpleDecipher(-3,3,"FCJJM")
Return value (array):
C@GGJ
DAHHK
EBIIL
FCJJM
GDKKN
HELLO
IFMMP
```
To enumerate offsets in another column you can use:
```
=SEQUENCE(MAX(OffsetStart,OffsetEnd)-MIN(OffsetStart,OffsetEnd)+1,1,OffsetStart,SIGN(OffsetEnd-OffsetStart))
```

## WildCardFilter(Array,SearchArray,SearchValue)
Filters Array by SearchArray where it contains SearchValue within an element, for example "EL" is contained in "HELLO"
```
=LAMBDA(Array,SearchArray,SearchValue,FILTER(Array,BYROW(SearchArray,LAMBDA(NAME,IFERROR(SEARCH(SearchValue,NAME),FALSE)))))
```

Examples:

```
Consider the following table named Contacts
| Name         | Email                   |
| ------------ | ----------------------- |
| Shinji Ikari | Shinji.Ikari@github.com |
| Gendo Ikari  | Gendo.Ikari@github.com  |
| Katsu Don    | Katsu.Don@github.com    |

=WildCardFilter(Contacts,Contacts[Name],"Ikari")
Output:
| Shinji Ikari | Shinji.Ikari@github.com |
| Gendo Ikari  | Gendo.Ikari@github.com  |
```
