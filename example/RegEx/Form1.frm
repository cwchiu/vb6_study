VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RegEx Demo"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtRegEx 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mobjRegEx As clsRegEx

Private Sub Form_Load()
    Set mobjRegEx = New clsRegEx
    txtRegEx.Text = GetSetting("REGEX", "FORM", "Expression", "[Ee]x.{2}pp?[^0-9][abcde]")
    txtData.Text = GetSetting("REGEX", "FORM", "Text", "Example")
    lblInfo.Caption = "Type a regular expression in the Left Hand box, and some text in the Right Hand box. If the text satisfies the expression then it will be displayed in Green. Text that does not satisfy the expression will be shown in Red."
    CheckForMatch
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting "REGEX", "FORM", "Expression", txtRegEx.Text
    SaveSetting "REGEX", "FORM", "Text", txtData.Text
    Set mobjRegEx = Nothing
End Sub

Private Sub txtData_Change()
    CheckForMatch
End Sub

Private Sub txtRegEx_Change()
' Ignore errors here - the DLL will raise them for invalid expressions which will occur as you type
' Eg. between typing an opening bracket and a closing one.
    On Error Resume Next
    mobjRegEx.Expression = txtRegEx.Text
    CheckForMatch
End Sub

Private Sub CheckForMatch()
    txtData.ForeColor = IIf(mobjRegEx.Match(txtData.Text), &H109910, vbRed)
End Sub

