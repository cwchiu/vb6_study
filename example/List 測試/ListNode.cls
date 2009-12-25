VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mData As String
Private mLink As ListNode

Public Property Let Data(newValue As String)
    mData = newValue
End Property

Public Property Get Data() As String
    Data = mData
End Property

Public Property Set Link(newValue As ListNode)
    Set mLink = newValue
End Property

Public Property Get Link() As ListNode
    Set Link = mLink
End Property

