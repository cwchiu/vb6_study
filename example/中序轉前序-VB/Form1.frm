VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   6000
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "postFix"
      Height          =   495
      Left            =   1710
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtExp 
      Height          =   495
      Left            =   210
      TabIndex        =   1
      Text            =   "1+2*3/4-5"
      Top             =   300
      Width           =   5265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "preFix"
      Height          =   495
      Left            =   210
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MAX_SIZE As Integer = 20

Dim arr_oprand() As String
Dim arr_oprator() As String

Dim stk_oprand As New clsStack
Dim stk_oprator As New clsStack


Private Sub Command1_Click()
   MsgBox infix2prefix(txtExp.Text)
End Sub

Private Sub Command2_Click()
   MsgBox infix2postfix(txtExp.Text)
End Sub

Private Sub Form_Load()
   Call stk_oprand.create(arr_oprand(), MAX_SIZE)
   Call stk_oprator.create(arr_oprator(), MAX_SIZE)
End Sub

' 中序轉後序
Function infix2postfix(s As String) As String
   Dim i As Integer
   Dim ch As String
   Dim ch_stack As String
   For i = 1 To Len(s)
      ch = Mid(s, i, 1)
      If IsNumeric(ch) Then
         infix2postfix = infix2postfix & ch
      ElseIf ch = "(" Then
         Call stk_oprator.push(arr_oprator(), ch)
      ElseIf ch = ")" Then
         Do Until (stk_oprator.isEmpty())
            ch_stack = stk_oprator.stacktop(arr_oprator())
            If ch_stack = "empty" Then
               Exit Do
            Else
               If ch_stack <> "(" Then
                  ch_stack = stk_oprator.pop(arr_oprator())
                  infix2postfix = infix2postfix & ch_stack
               Else
                  ch_stack = stk_oprator.pop(arr_oprator())
                  Exit Do
               End If
            End If
         Loop
      Else
         Do Until (stk_oprator.isEmpty())
            ch_stack = stk_oprator.stacktop(arr_oprator())
            If ch_stack = "empty" Then
               Exit Do
            Else
               If priority(ch) <= priority(ch_stack) Then
                  ch_stack = stk_oprator.pop(arr_oprator())
                  infix2postfix = infix2postfix & ch_stack
               Else
                  Exit Do
               End If
            End If
         Loop
         Call stk_oprator.push(arr_oprator(), ch)
      End If
   Next i
   
   Do Until stk_oprator.isEmpty
      infix2postfix = infix2postfix & stk_oprator.pop(arr_oprator())
   Loop
End Function

' ' 中序轉前序
Function infix2prefix(s As String) As String
   Dim i As Integer
   Dim ch As String
   Dim ch_stack As String
   For i = Len(s) To 1 Step -1
      ch = Mid(s, i, 1)
      If IsNumeric(ch) Then
         Call stk_oprand.push(arr_oprand(), ch)
      Else
         Do Until (stk_oprator.isEmpty())
            ch_stack = stk_oprator.stacktop(arr_oprator())
            If ch_stack = "empty" Then
               Exit Do
            Else
               If priority(ch) < priority(ch_stack) Then
                  ch_stack = stk_oprator.pop(arr_oprator())
                  Call stk_oprand.push(arr_oprand(), ch_stack)
               Else
                  Exit Do
               End If
            End If
         Loop
         Call stk_oprator.push(arr_oprator(), ch)
      End If
   Next i
   
   Do Until stk_oprator.isEmpty
       Call stk_oprand.push(arr_oprand(), stk_oprator.pop(arr_oprator()))
   Loop
   
   Do Until stk_oprand.isEmpty
      infix2prefix = infix2prefix & stk_oprand.pop(arr_oprand())
   Loop
End Function

' 取得優先權值
Function priority(x As String) As Integer
   Select Case Left(x, 1)
      Case "+"
         priority = 1
      Case "-"
         priority = 1
      Case "*"
         priority = 2
      Case "/"
         priority = 2
      Case ")"
         priority = 99
      Case "("
         priority = 0
   End Select
 End Function

Private Sub txtExp_Validate(Cancel As Boolean)
   Dim i As Integer
   Dim ch As String
   Dim s As String
   For i = 1 To Len(txtExp.Text)
      ch = Mid(txtExp.Text, i, 1)
      If IsNumeric(ch) Or _
         ch = "+" Or ch = "-" Or ch = "*" Or ch = "/" _
         Or ch = "(" Or ch = ")" Then s = s & ch
   Next i
   txtExp.Text = s
End Sub
