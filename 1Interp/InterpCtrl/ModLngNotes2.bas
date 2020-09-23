Attribute VB_Name = "ModLngNotes"
'
'Please Read
'
'*********************************************
'To add a new keyword to the system, you must
'follow the steps below.
'---------------------------------------------
'
'1) Add the keyword to ModLangDef's Default sub
'   with the SetKey method
'
'2) Add an enumeration element to ValidInput in
'   ModLangUDT's General section
'
'3) Add the structure for the keyword use, as a
'   Public Constant in ModLangUDT's General
'   section
'
'4) Increment the constant value of LFNUM found
'   in ModLangUDT's General section
'
'5) Add a new LF(n=n+1) assignment in ModLang's
'   ParseCode Sub cooresponding to the Public
'   Constant you added in step 3.
'
'6) Add a new case to ModLang's DoIt function that
'   corresponds to the new ValidInput enum
'
'7) Add a new Modlue that performs the work of
'   the new keyword that will be called from
'   ModLang's DoIt function
'
'8) Test the new keyword
'
'*********************************************



'*********************************************
'To add a new symbol to the system, you must
'follow the steps below.
'---------------------------------------------
'
'1)Add the symbol to ModLangDef's Default sub
'   with the SetSymbol method
'
'*********************************************



'*********************************************
'Adding a structure string to ModLangUDT's
'   General section.
'---------------------------------------------
'Each entry in the structure string must be
'   a numeric value. All numeric values are
'   followed by a comma, including the last
'   value in the string.
'
'   The numeric values of the structure string
'   must corespond to one of the following:
'
'   A: the value index+1 of the keyword as assigned
'       by the SetKey method
'   B: the value of the sum of keys(the number of
'       keywords) and index+1 of the symbol as
'       assigned by the SetSymbol mehtod
'   or
'   C: 0. The value assigned if the word is neither
'       a system keyword nor a system symbol.
'
'   e.g. if SetKey "IF" and SetKey "THEN" are the
'   only two key words assigned (in the order given),
'   and SetSymbol = "=" is the first symbol assigned,
'   then you would use the following structure string
'   to deifne the normal VB IF x = y THEN structure:
'   Public Const myForThen = "1,0,3,0,2,"
'
'*********************************************



'*********************************************
'Adding a ValidInput enum to ModLangUDT's
'   General section.
'---------------------------------------------
'Insert a new Label at the end of the ValidInput
'   enumerations setting its value equal to
'   one more than the previous enumeration.
'
'   e.g. Add MYLANGIFTHEN as a new enum when
'   only one enum (FIRSTENUM = 1) currenty exists;
'   Insert MYLANGIFTHEN = 2 below FIRSTENUM and
'   above END Enum
'
'*********************************************
