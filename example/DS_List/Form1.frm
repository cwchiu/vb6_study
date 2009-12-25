VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   5565
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim myList As CList
    Set myList = New CList
    
    Dim newItem As CListItem
    Set newItem = New CListItem
    newItem.Data = "1"
    myList.AddNew newItem

    Set newItem = New CListItem
    newItem.Data = "3"
    myList.AddNew newItem


    Set newItem = New CListItem
    newItem.Data = "5"
    Dim frontNode As CListItem
    Set frontNode = New CListItem
    frontNode.Data = "1"
    myList.InsertNode frontNode, newItem
    
    Dim node As CListItem
    Set node = New CListItem
    node.Data = "3"
    myList.Delete node
    
    Set newItem = New CListItem
    newItem.Data = "2"
    myList.AddNew newItem

    myList.track
End Sub
