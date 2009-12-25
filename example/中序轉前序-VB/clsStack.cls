VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------
' 種類：類別
' 功能：堆疊
' 作者：Arick
' 日期：2001/11/1
' 範例：
' 1. Dim stk_exp As New clsStack  ' 建立一個新的類別實體
' 2. Call stk_exp.create(stk1(), 5) ' 建立堆疊
' 3. Call stk_exp.push(stk1(), "A") ' 推資料到堆疊中
' 4. MsgBox stk_exp.stacktop(stk1()) ' 從堆疊中取出資料並顯示
'
' 已知限制：
' 1. 只能處理'單一數字'和'四則運算(+-*/)
' 2. 尚未加入確認符號功能
'-----------------------------------------------------------
Private MAX_STACK_SIZE As Integer
Private ptr_top As Integer

'|   |           overflow
'|   |           MAX
' ~~~
'|   |
'| D | <-- top    3
'| C |            2
'| B |            1
'| A |            0
'+---+           underflow
'
'-----------------------------------------------------------
' 功能：建立堆疊
' 傳入：stack() - 基底陣列
'       MAX_STACK_SIZE - 堆疊上限
' 回傳：目前指標
'-----------------------------------------------------------
Public Function create(stack() As String, ByVal MAX_SIZE As Integer) As Integer
   ptr_top = -1 ' 表示沒有任何資料
   MAX_STACK_SIZE = MAX_SIZE
   ReDim stack(MAX_STACK_SIZE)
End Function

'-----------------------------------------------------------
' 功能：回傳堆疊頂端元素
' 傳入：stack() - 基底陣列
'       lpTop   - 回傳目前堆疊指標
' 回傳：堆疊頂端元素
'-----------------------------------------------------------
Public Function stacktop(stack() As String, Optional lpTop As Integer) As String
   If isEmpty() Then
      stacktop = "empty"
   Else
      stacktop = stack(ptr_top)
      lpTop = ptr_top
   End If
End Function

'-----------------------------------------------------------
' 功能：push()
' 傳入：stack() - 基底陣列
'       item - 推入元素
'       lpTop   - 回傳目前堆疊指標
' 回傳：無
'-----------------------------------------------------------
Public Sub push(stack() As String, item As String, Optional lpTop As Integer)
   If (ptr_top >= MAX_STACK_SIZE - 1) Then
      MsgBox "stack overflow"
   Else
      ptr_top = ptr_top + 1
      stack(ptr_top) = item
      lpTop = ptr_top
   End If
End Sub

'-----------------------------------------------------------
' 功能：pop()
' 傳入：stack() - 基底陣列
'       lpTop   - 回傳目前堆疊指標
' 回傳：堆疊元素/empty(堆疊已空識別字)
'-----------------------------------------------------------
Public Function pop(stack() As String, Optional lpTop As Integer) As String
   If isEmpty() Then
      pop = "empty"
   Else
      pop = stack(ptr_top)
      ptr_top = ptr_top - 1
      lpTop = ptr_top
   End If
End Function

'-----------------------------------------------------------
' 功能：判斷堆疊是否為空
' 傳入：無
' 回傳：布林值
'-----------------------------------------------------------
Public Function isEmpty() As Boolean
   isEmpty = False
   If ptr_top = -1 Then isEmpty = True
End Function
