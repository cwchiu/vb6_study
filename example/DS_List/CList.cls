VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private col As Collection

' -- 新增一個節點在最後
Public Sub AddNew(ByRef newItem As CListItem)
    Set newItem.NextItem = Nothing
    
    Dim lastNode As CListItem
    Set lastNode = getLastNode()
    If Not lastNode Is Nothing Then
        Set lastNode.NextItem = newItem
    End If
    col.Add newItem
End Sub

' -- 插入一個節點在指定節點之後
Public Sub InsertNode(ByRef frontNode As CListItem, ByRef newItem As CListItem)
    Dim node As CListItem
    Set node = findNode(frontNode)
    
    If Not node Is Nothing Then
        Set newItem.NextItem = node.NextItem
        Set node.NextItem = newItem
    End If
    col.Add newItem
End Sub

' -- 刪除指定節點
Public Sub Delete(ByRef node As CListItem)
    If col Is Nothing Or col.Count = 0 Then
        Exit Sub
    End If
    
    Dim preNode As CListItem
    Set node = findNode(node, preNode)
    If Not preNode Is Nothing Then
        Set preNode.NextItem = node.NextItem
    End If
    
    Dim i As Long
    For i = 1 To col.Count
        If col(i).Data = node.Data Then
            col.Remove i
            Exit For
        End If
    Next
End Sub

Private Function findNode(ByRef node As CListItem, Optional ByRef preNode As CListItem = Nothing) As CListItem
    Dim tmpNode As CListItem
    If col Is Nothing Or col.Count = 0 Then
        Set findNode = Nothing
        Exit Function
    End If
    
    Set tmpNode = col.Item(1)
    Do While (Not tmpNode.NextItem Is Nothing)
        If tmpNode.Data = node.Data Then
            Exit Do
        End If
        Set preNode = tmpNode
        Set tmpNode = tmpNode.NextItem
    Loop
    Set findNode = tmpNode
End Function

Public Sub track()
    Dim node As CListItem
    If col Is Nothing Then
        Set getLastNode = Nothing
        Exit Sub
    End If
    
    Set node = col.Item(1)
    Debug.Print node.Data
    
    Do While (Not node.NextItem Is Nothing)
        Set node = node.NextItem
        Debug.Print node.Data
    Loop
End Sub

' 取得最後一個節點
Private Function getLastNode() As CListItem
    Dim node As CListItem
    If col Is Nothing Or col.Count = 0 Then
        Set getLastNode = Nothing
        Exit Function
    End If
    
    Set node = col.Item(1)
    Do While (Not node.NextItem Is Nothing)
        Set node = node.NextItem
    Loop
    Set getLastNode = node
End Function

Private Sub Class_Initialize()
    Set col = New Collection
End Sub

Private Sub Class_Terminate()
    Set col = Nothing
End Sub
