VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  '沒有框線
   Caption         =   "Tails..."
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   437
   ScaleMode       =   3  '像素
   ScaleWidth      =   545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Dots
    x As Integer
    Y As Integer
    A As Double
    R As Double
    W As Boolean
End Type

Dim Worm(100) As Dots
Dim World As Dots
Dim counter As Integer
Dim Quit As Boolean
Dim R As Integer
Dim G As Integer
Dim B As Integer

Private Type Col
    R As Integer
    G As Integer
    B As Integer
End Type

Dim Color(1024) As Col

Dim MyC As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Quit = True
End Sub

Private Sub Form_Load()
    MakeColors
    Me.Show
    DoEvents
    
    ' 視窗中心
    For i = 0 To 100
        Worm(i).x = (Me.ScaleWidth / 2)
        Worm(i).Y = (Me.ScaleHeight / 2)
    Next i

    Do While Quit = False
        'the colors
        MyC = MyC - 1
        If MyC < 0 Then MyC = 1024
        
        'world calcs...
        World.A = World.A + 0.03
        
        If World.W = False Then
            World.R = World.R + 0.2
        Else
            World.R = World.R - 0.1
        End If
        
        If World.R > (Me.ScaleWidth / 3) Then
            World.W = True
        ElseIf World.R < 10 Then
            World.W = False
        End If
        
        If World.A > (3.1415 * 2) Then
            World.A = World.A - (3.1415 * 2)
        End If
        
        World.x = (Me.ScaleWidth / 2) + Cos(World.A) * World.R
        World.Y = (Me.ScaleHeight / 2) + Sin(World.A) * World.R
        
        'worm calcs...
        
        'If counter > 10 Then
        For i = 100 To 1 Step -1
            Worm(i).x = Worm(i - 1).x
            Worm(i).Y = Worm(i - 1).Y
        Next i
    
            
        Worm(0).A = Worm(0).A - 0.08
        If Worm(0).A < 0 Then
            Worm(0).A = Worm(0).A + (3.1415 * 2)
        End If
        
        If Worm(0).W = False Then
            Worm(0).R = Worm(0).R + 0.1
        Else
            Worm(0).R = Worm(0).R - 0.2
        End If
        
        If Worm(0).R > (Me.ScaleWidth / 4) Then
            Worm(0).W = True
        ElseIf Worm(0).R < 20 Then
            Worm(0).W = False
        End If
        
        Worm(0).x = World.x + Cos(Worm(0).A) * Worm(0).R
        Worm(0).Y = World.Y + Cos(Worm(0).A) * Worm(0).R
        
        Me.Cls
        
        
        For i = 99 To 0 Step -1
            x = 1024 - (MyC + (i * 2))
            If x < 0 Then x = x + 1024
            
            R = Color(x).R - i / 2
            G = Color(x).G - i / 2
            B = Color(x).B - i / 2
            
            If R < 0 Then R = 0
            If G < 0 Then G = 0
            If B < 0 Then B = 0
            
            Me.Line (Worm(i).x, Worm(i).Y)-(Worm(i + 1).x, Worm(i + 1).Y), RGB(R, G, B)
            
            For u = -4 To 4
                If u <> 0 Then
                    Me.Line (Worm(i).x + i / u, Worm(i).Y + i / u)-(Worm(i + 1).x + i / u, Worm(i + 1).Y + i / u), RGB(R, G, B)
                End If
            Next u
        Next
    
        DoEvents
    Loop
    
    End
End Sub


Private Sub MakeColors()
    Dim x As Integer
    
    R = 255
     
    ' R = 255
    ' G = 0-255
    ' B = 0
    For i = 0 To 255 Step 1
        G = i
        x = x + 1
        Color(x).R = R
        Color(x).G = G
        Color(x).B = B
    Next i
    
    ' R = 255 - 0
    ' G = 255
    ' B = 0 - 255
    G = 255
    For i = 0 To 255 Step 1
        R = 255 - i
        B = i
        x = x + 1
        Color(x).R = R
        Color(x).G = G
        Color(x).B = B
    Next i
    
    ' R = 0
    ' G = 255 - 0
    ' B = 255
    R = 0
    For i = 0 To 255 Step 1
        G = 255 - i
        x = x + 1
        Color(x).R = R
        Color(x).G = G
        Color(x).B = B
    Next i
    
    
    ' R = 0 - 255
    ' G = 0
    ' B = 255 - 0
    G = 0
    For i = 0 To 255 Step 1
        R = i
        B = 255 - i
        x = x + 1
        Color(x).R = R
        Color(x).G = G
        Color(x).B = B
    Next i
End Sub
