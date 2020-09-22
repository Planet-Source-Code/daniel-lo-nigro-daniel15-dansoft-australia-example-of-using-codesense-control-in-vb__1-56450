Attribute VB_Name = "modFunctionDefinitions"
Option Base 1
Option Explicit

'Type used for AutoComplete
Type ObjectDescription
    strMembers() As String
    intMemberType() As MemberTypes
End Type

'The different member types of an object -
'Const, Enum, Function and Property (variable)
Enum MemberTypes
    memConst
    memEnum
    memFunction
    memProperty
End Enum

'Type used to hold function definitions (the little
'tooltip that pops up)
Type FunctionDescription
    strDef() As String
End Type

'Function List and Function Description variables
'Function list holds key numbers for descriptions (a lookup table),
'eg. if colFuncList("MsgBox") = 4, then
' udtFuncDesc(4).strDef(1) is the first definition
'for MsgBox(key no. 4). Understand?
Global colFuncList As New Collection
Global udtFuncDesc() As FunctionDescription

'Object List and Object info vars
Global colObjList As New Collection
Global udtObjInfo() As ObjectDescription

Dim intOldCount

'InitializeFuncs ---
' Purpose: Initialize variables and add in functions
Sub InitializeFuncs()

'make ObjectInfo array have 3 items
ReDim udtObjInfo(2)
'object #1 is always where standard autocomplete
'functions go (that don't have definitions listed
' below)
AddObject "<default>", "CvbBlack", "CvbRed", "CvbGreen", "CvbYellow", "CvbBlue", "CvbMagenta", "CvbCyan", "CvbWhite", _
          "CvbBinaryCompare", "CvbTextCompare", _
          "CvbSunday", "CvbMonday", "CvbTuesday", "CvbWednesday", "CvbThursday", "CvbFriday", "CvbUseSystem", "CvbUseSystemDayOfWeek", "CvbFirstJan1", "CvbFirstFourDays", "CvbFirstFullWeek", _
          "CvbGeneralDate", "CvbLongDate", "CvbShortDate", "CvbLongTime", "CvbShortTime", _
          "CvbObjectError", _
          "CvbOKOnly", "CvbOKCancel", "CvbAbortRetryIgnore", "CvbYesNoCancel", "CvbYesNo", "CvbRetryCancel", "CvbCritical", "CvbQuestion", "CvbInformation", "CvbExclamation", "CvbDefaultButton1", "CvbDefaultButton2", "CvbDefaultButton3", "CvbDefaultButton4", "CvbApplicationModal", "CvbSystemModal", "CvbOK", "CvbCancel", "CvbAbort", "CvbRetry", "CvbIgnore", "CvbYes", "CvbNo", _
          "CvbCr", "CvbCrLf", "CvbFormFeed", "CvbLf", "CvbNewLine", "CvbNullChar", "CvbNullString", "CvbTab", "CvbVerticalTab", _
          "CvbUseDefault", "CvbTrue", "CvbFalse", _
          "CvbNull", "CvbEmpty", "CvbInteger", "CvbLong", "CvbSingle", "CvbDouble", "CvbCurrency", "CvbDate", "CvbString", "CvbObject", "CvbError", "CvbBoolean", "CvbVariant", "CvbDataObject", "CvbDecimal", "CvbByte", "CvbArray", _
          "Cof_Input", "Cof_Output", "Cof_Append", "CTextBox", "CRadioBox", "CCheckBox", "COKButton", "CCancelButton", "CPicture", "CLabel"

'AddObject "DanProg", "FOpenFile", "FReadFromFile", "FWriteToFile", "FEOF", "FCloseFile"
AddObject "dpDialog", "FaddControl", "Pcancelled", "Pcaption", "FfreeResources", "FgetCheckInput", "FgetOptionInput", "FgetTextInput", "PHeight", "PLength", "PX", "PY"

'add a test object
'AddObject "testobj", "Ptestprop", "Ctestconst", "Ftestfunction", "Otestother", "Etestenum", "FAAAAAAAAA"

'size function description array to have 104 items
ReDim udtFuncDesc(114)
'get old count of autocomplete array
intOldCount = UBound(udtObjInfo(1).strMembers)
'ReDim Preserve udtObjInfo(1).strMembers(intOldCount + UBound(udtFuncDesc))
'ReDim Preserve udtObjInfo(1).intMemberType(intOldCount + UBound(udtFuncDesc))

'add standard vbscript functions
AddFunc "Abs", True, "Number"
AddFunc "Array", True, "Arglist"
AddFunc "Asc", True, "String"
AddFunc "Atn", True, "Number"
AddFunc "CBool", True, "Expression"
AddFunc "CByte", True, "Expression"
AddFunc "CCur", True, "Expression"
AddFunc "CDate", True, "Date"
AddFunc "CDbl", True, "Expression"
AddFunc "Chr", True, "CharCode"
AddFunc "CInt", True, "Expression"
AddFunc "CLng", True, "Expression"
AddFunc "Cos", True, "number"
AddFunc "CreateObject", True, "servername.typename, [location]"
AddFunc "CSng", True, "expression"
AddFunc "CStr", True, "expression"
AddFunc "Date", True, ""
AddFunc "DateAdd", True, "interval, number, date"
AddFunc "DateDiff", True, "interval, date1, date2, [firstdayofweek], [firstdayofyear]"
AddFunc "DatePart", True, "interval, date, [firstdayofweek], [firstdayofyear]"
AddFunc "DateSerial", True, "year, month, day"
AddFunc "DateValue", True, "date"
AddFunc "Day", True, "date"
AddFunc "Eval", True, "expression"
AddFunc "Exp", True, "number"
AddFunc "Filter", True, "InputStrings, Value, [include], [compare]"
AddFunc "FormatCurrency", True, "Expression, [NumDigitsAfterDecimal], [IncludeLeadingDigit], [UseParansForNegativeNumbers], [GroupDigits]"
AddFunc "FormatDateTime", True, "Date, [NamedFormat]"
AddFunc "FormatNumber", True, "Expression, [NumDigitsAfterDecimal], [IncludeLeadingDigit], [UseParansForNegativeNumbers], [GroupDigits]"
AddFunc "FormatPercent", True, "Expression, [NumDigitsAfterDecimal], [IncludeLeadingDigit], [UseParansForNegativeNumbers], [GroupDigits]"
AddFunc "GetLocale", True, ""
AddFunc "GetObject", True, "[pathname], [class]"
AddFunc "GetRef", True, "procname"
AddFunc "Hex", True, "number"
AddFunc "Hour", True, "time"
AddFunc "InputBox", True, "prompt, [title], [default], [xpos], [ypos], [helpfile], [context]"
AddFunc "InStr", True, "string1, string2, [compare]", "start, string1, string2, [compare]"
AddFunc "InStrRev", True, "string1, string2, [start], [compare]"
AddFunc "Int", True, "number"
AddFunc "Fix", True, "number"
AddFunc "IsArray", True, "varname"
AddFunc "IsDate", True, "expression"
AddFunc "IsEmpty", True, "expression"
AddFunc "IsNull", True, "expression"
AddFunc "IsNumeric", True, "expression"
AddFunc "IsObject", True, "expression"
AddFunc "Join", True, "list, [delimiter]"
AddFunc "LBound", True, "arrayname, [dimension]"
AddFunc "LCase", True, "string"
AddFunc "Left", True, "string, length"
AddFunc "Len", True, "string/varname"
AddFunc "LoadPicture", True, "picturename"
AddFunc "Log", True, "number"
AddFunc "LTrim", True, "string"
AddFunc "RTrim", True, "string"
AddFunc "Trim", True, "string"
AddFunc "Mid", True, "string, start, [length]"
AddFunc "Minute", True, "time"
AddFunc "Month", True, "date"
AddFunc "MonthName", True, "month, [abbreviate]"
AddFunc "MsgBox", True, "prompt, [buttons], [title], [helpfile], [context]"
AddFunc "Now", True, ""
AddFunc "Oct", True, "number"
AddFunc "Replace", True, "expression, find, replacewith, [start], [count], [compare]"
AddFunc "RGB", True, "red, green, blue"
AddFunc "Right", True, "string, length"
AddFunc "Rnd", True, "[number]"
AddFunc "Round", True, "expression, [numdecimalplaces]"
AddFunc "ScriptEngine", True, ""
AddFunc "ScriptEngineBuildVersion", True, ""
AddFunc "ScriptEngineMajorVersion", True, ""
AddFunc "ScriptEngineMinorVersion", True, ""
AddFunc "Second", True, "time"
AddFunc "SetLocale", True, "lcid"
AddFunc "Sgn", True, "number"
AddFunc "Sin", True, "number"
AddFunc "Space", True, "number"
AddFunc "Split", True, "expression, [delimiter], [count], [compare]"
AddFunc "Sqr", True, "number"
AddFunc "StrComp", True, "string1, string2, [compare]"
AddFunc "String", True, "number, character"
AddFunc "StrReverse", True, "string"
AddFunc "Tan", True, "number"
AddFunc "Time", True, ""
AddFunc "Timer", True, ""
AddFunc "TimeSerial", True, "hour, minute, second"
AddFunc "TimeValue", True, "time"
AddFunc "TypeName", True, "varname"
AddFunc "UBound", True, "arrayname. [dimension]"
AddFunc "UCase", True, "string"
AddFunc "VarType", True, "varname"
AddFunc "Weekday", True, "date, [firstdayofweek]"
AddFunc "WeekdayName", True, "weekday, [abbreviate], [firstdayofweek]"
AddFunc "Year", True, "date"

'add in our custom commands
AddFunc "OpenFile", True, "path, OpenType [of_Input/of_Output/of_Append], [FileNum]"
AddFunc "ReadFromFile", True, "filenum"
AddFunc "WriteToFile", True, "filenum, WhatToWrite"
AddFunc "EOF", True, "filenum"
AddFunc "CloseFile", True, "[filenum]"

AddFunc "addControl", False, "controlName, controlType, X, Y, Height, Length, [controlCaption], [controlPicture]"
AddFunc "freeResources", False, ""
AddFunc "getCheckInput", False, "controlName"
AddFunc "getTextInput", False, "controlName"
AddFunc "getOptionInput", False, ""



'add test function
'AddFunc "TestFunction", "TestFunction(test1 as String, test2 as Integer, [test3]) as Test", "TestFunction(bla, [blo])", "TestFunction(blar, [blor])"
End Sub

'AddFunc ---
' PURPOSE: Add a function into definition array
' INPUTS:
'  strFuncName - Function Name
'  ParamArray strFuncDefs - Function parameters
' RETURNS: New index number
' EXAMPLE: AddFunc("test", "test1", "foo, bar")
Function AddFunc(strFuncName As String, bolAddToAutoComplete As Boolean, ParamArray strFuncDefs()) As Integer
Dim intNewCount As Integer
Dim intTemp As Integer

'find new id
intNewCount = colFuncList.Count + 1
'add function to lookup table
colFuncList.Add intNewCount, strFuncName

'if we have to add to autocomplete list
If bolAddToAutoComplete = True Then
    intTemp = UBound(udtObjInfo(1).strMembers) + 1
    ReDim Preserve udtObjInfo(1).strMembers(intTemp)
    ReDim Preserve udtObjInfo(1).intMemberType(intTemp)
    udtObjInfo(1).strMembers(intOldCount + intNewCount) = strFuncName
    udtObjInfo(1).intMemberType(intOldCount + intNewCount) = memFunction
End If

'resize definition array to hold the number of
'definitions passed in
ReDim udtFuncDesc(intNewCount).strDef(UBound(strFuncDefs) - LBound(strFuncDefs) + 1)
'loop through all defs...
For intTemp = LBound(strFuncDefs) To UBound(strFuncDefs)
    '...and add them to def array
    udtFuncDesc(intNewCount).strDef(intTemp - LBound(strFuncDefs) + 1) = strFuncName & "(" & strFuncDefs(intTemp) & ")"
Next intTemp

'return new index
AddFunc = intNewCount
End Function

'FuncDefined ---
' PURPOSE: Find out if a function is defined
' INPUTS:
'  strFunc - The Function name
' RETURNS: Boolean stating whether function is defined
' EXAMPLE: bolTemp = FuncDefined("MsgBox")
'
' I won't comment this as anyone should understand
' how it works ;)
Function FuncDefined(strFunc As String) As Boolean
Dim intTemp As Integer
On Error GoTo nofunc
intTemp = colFuncList(strFunc)
FuncDefined = True
Exit Function

nofunc:
FuncDefined = False
End Function

'AddObject (similar to AddFunc)---
' PURPOSE: Add Object into AutoComplete array
' INPUTS:
'  strObjName - Name of object
'  ParamArray strObjMembers - Members of this object,
'   prefixed by P(roperty), F(unction), C(onst), or E(num)
' RETURNS: New index number
' EXAMPLE: AddObject("testobj", "Ftestfunction", "Etestenum")
Function AddObject(strObjName As String, ParamArray strObjMembers()) As Integer
Dim intNewCount As Integer
Dim intTemp As Integer

'Find new index
intNewCount = colObjList.Count + 1
'Add object to lookup table
colObjList.Add intNewCount, strObjName

'resize array of members
ReDim udtObjInfo(intNewCount).strMembers(UBound(strObjMembers) - LBound(strObjMembers) + 1)
'resize array of members' type
ReDim udtObjInfo(intNewCount).intMemberType(UBound(strObjMembers) - LBound(strObjMembers) + 1)
'loop through all the members passed in...
For intTemp = LBound(strObjMembers) To UBound(strObjMembers)
    '...find the member type...
    Select Case Left(strObjMembers(intTemp), 1)
        Case "P": udtObjInfo(intNewCount).intMemberType(intTemp - LBound(strObjMembers) + 1) = memProperty
        Case "F": udtObjInfo(intNewCount).intMemberType(intTemp - LBound(strObjMembers) + 1) = memFunction
        Case "C": udtObjInfo(intNewCount).intMemberType(intTemp - LBound(strObjMembers) + 1) = memConst
        Case "E": udtObjInfo(intNewCount).intMemberType(intTemp - LBound(strObjMembers) + 1) = memEnum
        Case Else: udtObjInfo(intNewCount).intMemberType(intTemp - LBound(strObjMembers) + 1) = memFunction
    End Select
    '...and add member to array
    udtObjInfo(intNewCount).strMembers(intTemp - LBound(strObjMembers) + 1) = Mid$(strObjMembers(intTemp), 2)
'continue loop
Next intTemp

'return new index
AddObject = intNewCount
End Function

'ObjDefined (similar to FuncDefined)---
' PURPOSE: Find out if object defined
' INPUTS:
'  strObject - Object to check
' RETURNS: True if object is defined
' EXAMPLE: bolTemp = ObjDefined("testobj")
'As with FuncDefined, i won't comment this.
Function ObjDefined(strObject As String) As Boolean
Dim strTemp As String
On Error GoTo noobj
strTemp = colObjList(strObject)
ObjDefined = True
Exit Function

noobj:
ObjDefined = False
End Function
'FuncString ---
' returns a string with all the functions in it,
' seperated by a VbLf (linefeed)
Function FuncString() As String
Dim strTemp As String
Dim intTemp As Integer
For intTemp = 1 To UBound(udtObjInfo(1).strMembers)
    If udtObjInfo(1).intMemberType(intTemp) = memFunction Then
        strTemp = strTemp & udtObjInfo(1).strMembers(intTemp) & vbLf
    End If
Next intTemp
strTemp = strTemp & "Call" & vbLf & "Class" & vbLf & "Dim" & vbLf & "Do" & vbLf & "Loop" & vbLf & "While" & vbLf & "Until" & vbLf & "Erase" & vbLf & "ExecuteGlobal" & vbLf & "Exit" & vbLf & "Next" & vbLf & "If" & vbLf & "Then" & vbLf & "Else" & vbLf & "On Error" & vbLf & "Option Explicit" & vbLf & "Private" & vbLf & "Property" & vbLf & "Get" & vbLf & "Let" & vbLf & "Set" & vbLf & "Public" & vbLf & "Randomize" & vbLf & "ReDim" & vbLf & "Select" & vbLf & "Wend" & vbLf & "With" & vbLf & "Case" & vbLf & "DoEvents" & vbLf & "Const" & vbLf & "Erach" & vbLf & "ElseIf" & vbLf & "Each" & vbLf & "ElseIf" & vbLf & "Format" & vbLf & "GoSub" & vbLf & "GoTo" & vbLf & "In" & vbLf & "Input" & vbLf & "Module" & vbLf & "New" & vbLf & "Open" & vbLf & "Preserve" & vbLf & "Print" & vbLf & "Put" & vbLf & "Read" & vbLf & "Resume" & vbLf & "Select" & vbLf & "Shared" & vbLf & "Static" & vbLf & "Stop" & vbLf & "To" & vbLf & "True"
FuncString = strTemp
End Function
