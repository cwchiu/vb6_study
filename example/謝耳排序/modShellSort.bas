Attribute VB_Name = "modShellSort"
'專案名稱：Shell 排序
'作者：Jack & Gisa (Visual Software)
'網址：http://www.hello.com.tw/~vjack
'信箱：vjack@hello.com.tw

Public Const MAX_MIN = True
Public Const MIN_MAX = False

Public Sub Main()
    Dim Num(100) As Integer
    For i = 0 To 100
        Num(i) = Int(Rnd * 21)
        Debug.Print Num(i);
    Next i
    ShellSort Num(), 100, MIN_MAX
    Debug.Print ""
    For j = 0 To 100
        Debug.Print Num(j);
    Next j
End Sub

Public Sub ShellSort(ByRef Values() As Integer, ByVal NumElements As Integer, ByVal Order As Integer)
    Dim Temp As Integer
    Dim i As Integer
    Dim Gap As Integer
    Dim ExchangeOccurred As Boolean
    Gap = NumElements / 2
    Do
        Do
            ExchangeOccurred = False
            For i = 0 To NumElements - Gap
                If ((Not Order) And (Values(i) > Values(i + Gap))) Or (Order And (Values(i) < Values(Gap + i))) Then
                    Temp = Values(i)
                    Values(i) = Values(i + Gap)
                    Values(i + Gap) = Temp
                    ExchangeOccurred = True
                Else
                End If
            Next i
        Loop While (ExchangeOccurred)
        Gap = Gap / 2
    Loop While (Gap <> 0)
End Sub

