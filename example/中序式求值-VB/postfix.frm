VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "���Ǧ��D�� "
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      Caption         =   "�w�����D"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1680
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "�H�U�����չL���B�⦡�A�i�H�������o�B�⦡"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim s As String
    s = infix2postfix(Trim(Text1.Text) & ETX)
    MsgBox (postfix(s & Chr(0)))
End Sub

Private Sub Command2_Click()
  Dim str As String
  str = "1. �u��B�z�Ӧ�Ʀr" + vbCrLf + _
        "2. �A���B�z������"
  MsgBox str, , "�w�����D"
End Sub

Private Sub Form_Load()
    Call init
    Label1.Caption = "�H�U�����չL���B�⦡�A�i�H�������o�B�⦡"
    List1.AddItem "1+2*3"                   ' 7
    List1.AddItem "2*3/4"                   ' 1.5
    List1.AddItem "2*3+4"                   ' 10
    List1.AddItem "(1+2)&4"                 ' 16
    List1.AddItem "1-2*3+4"                 ' -1
    List1.AddItem "1*2+3*4"                 ' 14
    List1.AddItem "6/2-3+4*2"               ' 8
    List1.AddItem "1*(2+3)*4"
    List1.AddItem "(1+2)*(3+4)"
    List1.AddItem "1+(2*3)+4"
    List1.AddItem "(1+2)*3+4"
    List1.AddItem "1+2*3+4"
    List1.AddItem "1+2*(3+4)"
    List1.AddItem "((4/2)-2)+(3*3)-(4*2)"   ' 1
    List1.AddItem "(4/(2-2+3))*(3-4)*2"     ' -2.6666
End Sub

Private Sub List1_DblClick()
    Text1.Text = List1.Text
End Sub
