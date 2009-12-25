VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Stack 的基本運作"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3555
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2205
   ScaleWidth      =   3555
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command4 
      Caption         =   "isEmpty"
      Height          =   315
      Left            =   2130
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "TOP"
      Height          =   315
      Left            =   2130
      TabIndex        =   3
      Top             =   1050
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "POP"
      Height          =   315
      Left            =   2130
      TabIndex        =   2
      Top             =   660
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  '平面
      Height          =   2010
      Left            =   45
      TabIndex        =   1
      Top             =   105
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PUSH"
      Height          =   315
      Left            =   2130
      TabIndex        =   0
      Top             =   255
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --------------------------------------------------
' 最後更新：2001/10//09
' 設計者：Arick
' 功能：模擬堆疊運作
' 網址：http://home.pchome.com.tw/world/sisimi
' 信箱：sisimi1@sinamail.com
' --------------------------------------------------
Option Explicit

Dim tmp() As String
Dim t As Integer

Private Sub Command1_Click()
   Call stk_push(tmp, random(100, 1), t)
   Call x
   End Sub

Private Sub Command2_Click()
   Dim sRet As String
   sRet = stk_pop(tmp, t)
   If sRet = "empty" Then
      MsgBox "目前堆疊中沒有任何資料", , "POP"
   Else
'      MsgBox sRet, , "POP"
   End If
   Call x
End Sub

Private Sub Command3_Click()
   Dim sRet As String
   sRet = stk_top(tmp, t)
   If sRet = "empty" Then
      MsgBox "目前堆疊中沒有任何資料", , "TOP"
   Else
      MsgBox sRet, , "TOP"
   End If
End Sub

Private Sub Command4_Click()
   If (stk_isEmpty) Then
      MsgBox "目前堆疊是空的", , "isEmpty"
   Else
      MsgBox "目前堆疊有資料", , "isEmpty"
   End If
End Sub

Private Sub Form_Load()
   Call stk_create(tmp)
   Call listClear
End Sub


Private Sub x()
   Dim i As Integer
   Call listClear
   For i = 0 To t - 1
      If i + 1 = t Then
         List1.List(MAX_STACK_SIZE - i) = tmp(i + 1) & vbTab & "<-- Top"
      Else
         List1.List(MAX_STACK_SIZE - i) = tmp(i + 1)
      End If
   Next i
End Sub

Private Function random(max As Integer, min As Integer) As Integer
   Randomize
   random = Int(Rnd * (max - min + 1)) + min
End Function

Private Sub listClear()
   Dim i As Integer
   List1.Clear
   For i = 0 To 9
      List1.AddItem ""
   Next i
End Sub

