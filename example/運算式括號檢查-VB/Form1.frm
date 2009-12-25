VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1740
      TabIndex        =   1
      Top             =   1350
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   660
      TabIndex        =   0
      Text            =   "(1+2*(3/4)+2"
      Top             =   315
      Width           =   3720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --------------------------------------------------
' 最後更新：2001/01//15
' 設計者：Arick
' 功能：運算式括號檢查
' 網址：http://home.pchome.com.tw/world/sisimi
' 信箱：sisimi1@sinamail.com
' --------------------------------------------------
Option Explicit

Private Sub Command1_Click()
Dim i As Integer, d As String
Dim t As Integer
For i = 1 To Len(Text1.Text)
   d = Mid(Text1.Text, i, 1)
   If d = "(" Then t = t + 1
   If d = ")" Then t = t - 1
   If t < 0 Then
      MsgBox "括號錯誤"
      Exit For
   End If
Next i
If t <> 0 Then MsgBox "括號錯誤"
End Sub

