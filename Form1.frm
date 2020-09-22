VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "CMCS21.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlPopup 
      Left            =   2760
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0360
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A20
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CodeSenseCtl.CodeSense cs 
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0D80
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   4575
   End
   Begin VB.Menu test 
      Caption         =   "&Test"
      Begin VB.Menu undo 
         Caption         =   "undoo"
      End
   End
   Begin VB.Menu mnuLoading 
      Caption         =   "<<<<PLEASE WAIT, LOADING>>>>"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Everything below is a wrapper to our functions, in
'the modAutoCompleteToolTip module
Private Function cs_CodeList(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
cs_CodeList = CodeList(Control, ListCtrl, imlPopup)
End Function

Private Function cs_CodeListSelMade(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
cs_CodeListSelMade = CodeListSelMade(Control, ListCtrl)
End Function

'cs_CodeListSelWord ---
' Tells CodeSense control to allow itself to choose
' word most like the current word
Private Function cs_CodeListSelWord(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal lItem As Long) As Boolean
cs_CodeListSelWord = True
End Function

Private Function cs_CodeTip(ByVal Control As CodeSenseCtl.ICodeSense) As CodeSenseCtl.cmToolTipType
cs_CodeTip = CodeTip(Control)
End Function

Private Sub cs_CodeTipInitialize(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ToolTipCtrl As CodeSenseCtl.ICodeTip)
CodeTipInitialize Control, ToolTipCtrl
End Sub

Private Sub cs_CodeTipUpdate(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ToolTipCtrl As CodeSenseCtl.ICodeTip)
CodeTipUpdate Control, ToolTipCtrl
End Sub

Private Function cs_KeyPress(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyAscii As Long, ByVal Shift As Long) As Boolean
cs_KeyPress = KeyPress(Control, KeyAscii, Shift)
End Function
'end wrappers

'use this to print (into the immediate window) any
'keys that are pressed
'Private Function cs_KeyDown(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyCode As Long, ByVal Shift As Long) As Boolean
'Debug.Print KeyCode, GetKeyMaskName(CByte(Shift)); GetVirtKeyName(KeyCode)
'End Function




'Form_Load ---
' Called when program starts
Private Sub Form_Load()
Me.Show
'disable text box
cs.Enabled = False
Dim objTempLang As New Language
'show loading message
mnuLoading.Caption = "Setting up defaults..."
'set default text
cs.Text = "'Starting procedure" & vbCrLf & "Sub Main()" & vbCrLf & vbCrLf & "End Sub"
'set colors for line numbers on side
cs.SetColor cmClrLineNumberBk, RGB(0, 0, 128)
cs.SetColor cmClrLineNumber, vbWhite
'make the properties of this CodeSense control
'the same for all of them in this program
cs.GlobalProps = True

mnuLoading.Caption = "Initializing AutoComplete / ToolTips..."
'initialize AutoComplete and ToolTip-thingy
Call InitializeFuncs

mnuLoading.Caption = "Initializing Syntax highlighting..."
'get the definition for Basic
Set objTempLang = CSGlobals.GetLanguageDef("Basic")
'get rid of all languages
CSGlobals.UnregisterAllLanguages
'fix up basic definition
objTempLang.ScopeKeywords1 = objTempLang.ScopeKeywords1 & vbLf & "Sub"
objTempLang.ScopeKeywords2 = objTempLang.ScopeKeywords2 & vbLf & "End Sub"
objTempLang.Keywords = FuncString
'register language (DanProgrammer VBScript)...
CSGlobals.RegisterLanguage "DanProgrammer VBScript", objTempLang
'...and set it as the language in use
cs.Language = "DanProgrammer VBScript"

mnuLoading.Caption = ""
mnuLoading.Visible = False
cs.Enabled = True
End Sub

Private Sub Form_Resize()
cs.Width = Me.Width - 225
cs.Height = Me.Height - 795
End Sub

'mnuTest_Click ---
' Set undo caption to 'Undo' and show the keypress on the right
Private Sub test_Click()
undo.Caption = "&Undo" & GetKeyPressName(CSGlobals.GetHotKeyForCmd(cmCmdUndo, 0))
End Sub

Private Sub undo_Click()
MsgBox "mwahahahahah"
End Sub
