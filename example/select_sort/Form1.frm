VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "������ܱƧ�"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   8130
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4590
      TabIndex        =   1
      Top             =   2085
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   5280
      Left            =   225
      TabIndex        =   0
      Top             =   315
      Width           =   4065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
' ----- ������ܱƧ� ----------
' �b�����ƦCa(i)~a(n-1)����X�̤p��,�Pa(i)�洫�A
' �q�����ƦCa(i)~a(n-1)�}�l�A���Ƴo�@�ާ@���쳡���ƦCa(n-1)����


Dim i As Integer, j As Integer
Dim a(5) As Integer
Dim t As Integer, s As Integer
Dim min As Integer
a(0) = 80
a(1) = 41
a(2) = 35
a(3) = 90
a(4) = 40
a(5) = 20
For i = 0 To 5
   min = a(i)
   s = i
   For j = i + 1 To 4
      If a(j) < min Then
         min = a(j)
         s = j
      End If
   Next j
   t = a(i)
   a(i) = a(s)
   a(s) = t
Next i

For i = 0 To 5
   List1.AddItem a(i)
Next i
End Sub


