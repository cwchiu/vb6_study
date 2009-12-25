Attribute VB_Name = "Module1"
Option Explicit
Public Const MAX_STACK_SIZE = 10
'Dim STACK() As String
Dim ptr_top As Integer

'|   |           overflow
'|   |           MAX
' ~~~
'|   |
'| D | <-- top    4
'| C |            3
'| B |            2
'| A |            1
'+---+           underflow
'
Public Sub stk_create(stack() As String)
   ptr_top = 0 ' 表示沒有任何資料
   ReDim stack(MAX_STACK_SIZE)
End Sub

Public Function stk_top(stack() As String, lpTop As Integer) As String
   If ptr_top = 0 Then
      stk_top = "empty"
   Else
      stk_top = stack(ptr_top)
      lpTop = ptr_top
   End If
End Function

Public Sub stk_push(stack() As String, item As String, lpTop As Integer)
   If ptr_top >= MAX_STACK_SIZE Then
      MsgBox "stack overflow"
   Else
      ptr_top = ptr_top + 1
      stack(ptr_top) = item
      lpTop = ptr_top
   End If
End Sub

Public Function stk_pop(stack() As String, lpTop As Integer) As String
   If ptr_top = 0 Then
      stk_pop = "empty"
   Else
      stk_pop = stack(ptr_top)
      ptr_top = ptr_top - 1
      lpTop = ptr_top
   End If
End Function

Public Function stk_isEmpty() As Boolean
   stk_isEmpty = False
   If ptr_top = 0 Then stk_isEmpty = True
End Function


