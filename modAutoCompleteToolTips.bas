Attribute VB_Name = "modAutoCompleteToolTips"
Option Base 1
Option Explicit

Global selRange As CodeSenseCtl.IRange
'globals of CodeSense Control
Global CSGlobals As New CodeSenseCtl.Globals
'current word (for tooltip)
Global strCurrentWord As String
'current word function key in udtFuncDesc (see modFunctionDefinitions)
Global intCurrentWordItem As Integer

'CodeList ---
' TRIGGERED: When AutoComplete key shortcut is pressed
' PURPOSE: Shows list
' INPUTS:
'  Control - CodeSense control that caused this
'  ListCtrl - List control assigned to us by
'             the CodeSense control.
' RETURNS: True (to show list)

Function CodeList(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ImageList As ImageList) As Boolean
Dim intObjItem As Integer
Dim intTemp As Integer

'set up list properties
'ListCtrl.BackColor = Control.GetColor(cmClrWindow)
ListCtrl.Font.Name = "Tahoma"
ListCtrl.Font.Size = 10
'image list for the little pictures to the side
'on the list
ListCtrl.ImageList = ImageList

'if current word is an object (ie. person typed
' a dot (.) after object name...
If ObjDefined(Control.CurrentWord) Then
    '...look in autocomplete list for that object
    intObjItem = colObjList(Control.CurrentWord)
'otherwise
Else
    'use normal autocomplete list
    intObjItem = 1
End If

'go through all autocomplete items...
For intTemp = 1 To UBound(udtObjInfo(intObjItem).strMembers)
    '...and add them to list
    ListCtrl.AddItem udtObjInfo(intObjItem).strMembers(intTemp), udtObjInfo(intObjItem).intMemberType(intTemp)
Next intTemp

'show list
CodeList = True
End Function

'CodeListSelMade ---
' TRIGGERED: When selection is made in list
' PURPOSE: To add item chosen to text
' INPUTS:
'  Control - CodeSense control that triggered this
'  ListCtrl - the list containing AutoComplete items
' RETURNS: False (to kill list)
Function CodeListSelMade(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
Dim strItem As String
Dim range As New CodeSenseCtl.range

'get current word in text
strItem = ListCtrl.GetItemText(ListCtrl.SelectedItem)
'if current word is start of word chosen in box
'(ie. user entered 'Msg', went into AutoComplete,
' and chose 'MsgBox, it would replace 'Msg' with
' 'MsgBox' instead of inserting it. If it inserted
' it, the word would become 'MsgMsgbox')...
If LCase$(Control.CurrentWord) = LCase$(Left$(strItem, Control.CurrentWordLength)) Then
    '...shorten text (if user typed 'Msg', went into
    ' autocomplete and chose MsgBox, this shortens
    ' the item to 'Box' ['Msg' + 'Box' = 'MsgBox'])
    strItem = Mid$(strItem, Control.CurrentWordLength + 1)
End If

'replace selection with this item
Control.ReplaceSel (strItem)

'get cursor position
Set range = Control.GetSel(True)
range.StartColNo = range.StartColNo + Len(strItem)
range.EndColNo = range.StartColNo
range.EndLineNo = range.StartLineNo
'set cursor position to just after word
Control.SetSel range, True

'kill list
CodeListSelMade = False
End Function


'CodeTip ---
' TRIGGERED: When ToolTip should be shown
' PURPOSE: To check if it really should be shown
' INPUTS:
'  Control - You should know by now ;)
' RETURNS: Type of tip to show (refer to documentation
'          on CodeSense control)
Function CodeTip(ByVal Control As CodeSenseCtl.ICodeSense) As CodeSenseCtl.cmToolTipType
Dim token As CodeSenseCtl.cmTokenType
'get current token type
token = Control.CurrentToken
'if current token is text or keyword...
If ((token = cmTokenTypeText) Or (token = cmTokenTypeKeyword)) Then
    '...save current word
    strCurrentWord = Control.CurrentWord
    'if current word is defined...
    If FuncDefined(strCurrentWord) Then
        '...get the index of it...
        intCurrentWordItem = colFuncList(strCurrentWord)
        '...and tell codesource control to show
        'tip.
        CodeTip = cmToolTipTypeMultiFunc
    'if word not defined...
    Else
        '...show no tip
        CodeTip = cmToolTipTypeNone
    End If
'if in a comment...
Else
    '...show no tip
    CodeTip = cmToolTipTypeNone
End If
End Function

'OK, from now on I am not putting the 'Control'
'variable in the input description, as it is
'getting annoying to type it in everytime!
'And yes, instead of cutting and pasting the
'header on every function, i type it out again
'and again and again and again and again ;)


'CodeTipInitialize ---
' TRIGGERED: When ToolTip is initializing
' PURPOSE: To initialize the tooltio
' INPUTS:
'  ToolTipCtrl - the tooltip created by the control
' RETURNS: Nothing
Sub CodeTipInitialize(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ToolTipCtrl As CodeSenseCtl.ICodeTip)
Dim tip As CodeSenseCtl.CodeTipMultiFunc
'get the control
Set tip = ToolTipCtrl

'set current argument to the first one (this is zero-based)
tip.Argument = 0

'save position
Set selRange = Control.GetSel(True)
selRange.EndColNo = selRange.EndColNo '+ 1

'get definition count for the function
tip.FunctionCount = UBound(udtFuncDesc(intCurrentWordItem).strDef) - 1
'set current definition to the first one (again, 0-based)
tip.CurrentFunction = 0
'set tip to the first one
tip.TipText = udtFuncDesc(intCurrentWordItem).strDef(1)

'set font
tip.Font.Name = "Arial"
tip.Font.Size = 10
tip.Font.Italic = True
tip.Font.Bold = False
End Sub

'CodeTipUpdate ---
' TRIGGERED: When tip should be updated
' PURPOSE: To update tip
' INPUTS:
'  ToolTipCtrl - current ToolTip
' RETURNS: Nothing
Sub CodeTipUpdate(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ToolTipCtrl As CodeSenseCtl.ICodeTip)
Dim iTrim As Integer, j As Integer
Dim bolInQuote As Boolean
Dim tip As CodeSenseCtl.CodeTipMultiFunc
'get tip
Set tip = ToolTipCtrl

Dim range As CodeSenseCtl.IRange
'get current cursor position
Set range = Control.GetSel(True)
'if user has moved up/down or before the statement...
If (range.EndLineNo <> selRange.EndLineNo) Or _
   (range.EndColNo < selRange.EndColNo) Then
   '...what do you think?
    tip.Destroy
Else
    Dim iArg, I As Integer
    Dim strLine As String

    iArg = 0
    'set i to the line number
    I = selRange.EndLineNo
    'get the line
    strLine = Control.GetLine(I)
    'set iTrim to length of line + 1
    iTrim = Len(strLine) + 1
    'if cursor isn't at end of line...
    If (range.EndColNo < iTrim) Then
        '...then iTrim = current cursor pos
        iTrim = range.EndColNo
    End If
    'get current line, up to iTrim
    strLine = Left(strLine, iTrim)
    bolInQuote = False
    j = 0
    'go through every character in line
    While ((Len(strLine) <> 0) And (j <= Len(strLine)) And (iArg <> -1))
        'check if quote encountered
        If (Mid(strLine, j + 1, 1) = """") Then
            bolInQuote = Not bolInQuote
        'if character is comma...
        ElseIf (Mid(strLine, j + 1, 1) = ",") And bolInQuote = False Then
            '...add 1 to argument count
            iArg = iArg + 1
        'if character is end bracket...
        ElseIf (Mid(strLine, j + 1, 1) = ")") And bolInQuote = False Then
            '...signal to destroy tip
            iArg = -1
        'if character is quote...
        ElseIf (Mid(strLine, j + 1, 1) = "'") And bolInQuote = False Then
            '...set iArg to -1 to destroy tip (since
            'user is starting a comment
            iArg = -1
        End If
        'add one to character count
        j = j + 1
    Wend
    'if tip should be destroyed...
    If (iArg = -1) Then
        '...destroy it, ...
        tip.Destroy
    '...otherwise
    Else
        'set number of current argument
        tip.Argument = iArg
        'set tiptext to current function description
        tip.TipText = udtFuncDesc(intCurrentWordItem).strDef(tip.CurrentFunction + 1)
    End If
End If
End Sub

'KeyPress ---
' TRIGGERED: When a key is pressed
' PURPOSE: To see if AutoComplete should be activated
' INPUTS:
'  KeyAscii - The ASCII code of the key
'  Shift - The KeyMask (eg. shift, alt or ctrl)
' RETURNS: Nothing
Function KeyPress(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyAscii As Long, ByVal Shift As Long) As Boolean
    Select Case KeyAscii
        'if key is starting bracket or space...
        Case (Asc("(")), (Asc(" "))
            '...show AutoComplete
            Control.ExecuteCmd (cmCmdCodeTip)
        'if key is dot...
        Case (Asc("."))
            '...and current word is a defined object...
            If ObjDefined(Control.CurrentWord) Then
                '...show autocomplete
                Control.ExecuteCmd cmCmdCodeList
            End If
    End Select
End Function
