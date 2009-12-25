Attribute VB_Name = "Module1"
Public Enum precedence
  lparen    ' (
  rparen    ' )
  plus      ' +
  minus     ' -
  times     ' /
  divide    ' *
  pow       ' &
  eos       ' chr(0)
  operand   ' �Ʀr
End Enum

Public Const MAX_STACK_SIZE = 100

Public Type stack
  top As Integer
  items(MAX_STACK_SIZE) As String
End Type


Public ETX As String
Public ps As stack
Public icp(7) As Integer   ' ��J�u���v
Public isp(7) As Integer   ' ���|�u���v

Sub init()
    ETX = Chr(0)
    
    isp(0) = 0      ' (
    isp(1) = 19     ' )
    isp(2) = 12     ' +
    isp(3) = 12     ' -
    isp(4) = 13     ' *
    isp(5) = 13     ' /
    isp(6) = 14     ' &
    isp(7) = 0      ' eos
    
    icp(0) = 20
    icp(1) = 19
    icp(2) = 12
    icp(3) = 12
    icp(4) = 13
    icp(5) = 13
    icp(6) = 14
    icp(7) = 0
End Sub

' �N��Ʊ��J���|
Sub push(data As String)
    If isFull Then
       MsgBox "Stack Overflow"
       Exit Sub
    Else
       ps.top = ps.top + 1
       ps.items(ps.top) = data
    End If
End Sub

' �P�_���|�O�_�w��
Function isFull() As Boolean
    isFull = (ps.top >= MAX_STACK_SIZE)
End Function

' �P�_���|�O�_����
Function isEmpty() As Boolean
    isEmpty = (ps.top = 0)
End Function

' �q���|�����X���
Function pop() As String
  If isEmpty Then
    MsgBox "Stack Underflow"
    Exit Function
  Else
    pop = ps.items(ps.top)
    ps.top = ps.top - 1
  End If
End Function

' �B�⦡�B�z
Function eval(op1 As String, op2 As String, c As String) As Single
  Select Case c
    Case "+"
      eval = Val(op1) + Val(op2)
    Case "-"
      eval = Val(op1) - Val(op2)
    Case "*"
      eval = Val(op1) * Val(op2)
    Case "/"
      eval = Val(op1) / Val(op2)
    Case "&"
      eval = Val(op1) ^ Val(op2)
  End Select
End Function


Function postfix(sym As String) As String
    Dim c As String, op1 As String, op2 As String
    Dim pos As Integer
    ' �����r��  chr(0)
    ' sym = "123*+"   --> 7
    ' sym = "123*+4+" --> 11
    ' sym = "24&"     --> 16
    pos = 1
    Do
      c = Mid(sym, pos, 1)
      If c = ETX Then Exit Do
      If IsNumeric(c) Then
         Call push(c)
      Else
         op2 = pop
         op1 = pop
         Call push(eval(op1, op2, c))
      End If
      pos = pos + 1
    Loop
    postfix = pop
End Function

' �u���v����
Function get_token(c As String) As precedence
    Select Case c
        Case "("
            get_token = lparen
        Case ")"
            get_token = rparen
        Case "+"
            get_token = plus
        Case "-"
            get_token = minus
        Case "/"
            get_token = divide
        Case "*"
            get_token = times
        Case "%"
            get_token = modx
        Case "&"
        Case ETX
            get_token = eos
        Case Else
            get_token = operand
    End Select
End Function

' ���ǹB�⦡���ǹB�⦡
Function infix2postfix(sym As String) As String
    Dim pos As Integer
    Dim c As String
    Dim token As precedence
    Dim op As Integer
    
    pos = 1
    Do
      c = Mid(sym, pos, 1)
      token = get_token(c)
      If c = ETX Then Exit Do ' �O�_�������r��
      If IsNumeric(c) Then    ' �Ʀr�B�z
         infix2postfix = infix2postfix & c
      ElseIf c = ")" Then ' �k�A���B�z
         Do
           c = pop
           If c = "(" Or ps.top <= 0 Then Exit Do
           infix2postfix = infix2postfix & c
         Loop
      Else
         If ps.top > 0 Then ' �O�_�b���|����
            Do ' �Y���|�u���v>�ثe�u���v
               op = get_token(ps.items(ps.top))
               If isp(op) < icp(token) Then Exit Do
               infix2postfix = infix2postfix & pop
               If ps.top <= 0 Then Exit Do
            Loop
         End If
         push c ' ���J���|
      End If
      pos = pos + 1
    Loop
    
    Do ' �N���|��ƨ̧Ǩ��X
       c = pop
       infix2postfix = infix2postfix & c
       If ps.top = 0 Then Exit Do
    Loop
End Function
