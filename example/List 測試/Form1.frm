VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim p(10) As ListNode
    Dim i As Long
    For i = 0 To 10
        Set p(i) = New ListNode
        p(i).Data = i
        If i <> 0 Then Set p(i - 1).Link = p(i)
    Next i
    
    ' (4)-> (5)-> (6)
    ' Remove Node
    Set p(4).Link = p(5).Link
    
    ' Add Node
    Dim tn As ListNode
    Set tn = p(1).Link
    Set p(1).Link = p(5)
    Set p(5).Link = tn
    
    Dim tmp As ListNode
    Set tmp = p(0).Link
    Do Until tmp Is Nothing
        Debug.Print tmp.Data
        Set tmp = tmp.Link
    Loop
End Sub
