VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSendMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Mail Sender"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmSendMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   3120
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Attach files"
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   6375
      Begin VB.ListBox lstAttachments 
         Height          =   840
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4695
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   4920
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddFile 
         Caption         =   "&Add..."
         Height          =   375
         Left            =   4920
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2400
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send message"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New Message"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtMessage 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Text            =   "frmSendMail.frx":030A
      Top             =   1560
      Width           =   6375
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Text            =   "txtSubject"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtRecipient 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Text            =   "txtRecipient"
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtSender 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "txtSender"
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "txtHost"
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      Height          =   195
      Left            =   1245
      TabIndex        =   3
      Top             =   1200
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Recipient e-mail address:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   840
      Width           =   1770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Your e-mail address:"
      Height          =   195
      Left            =   405
      TabIndex        =   1
      Top             =   480
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SMTP Host:"
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum SMTP_State
    MAIL_CONNECT
    MAIL_HELO
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
End Enum

Private m_State As SMTP_State
Private m_strEncodedFiles As String
'

Private Sub cmdAddFile_Click()
    
    With ComDialog
        .ShowOpen
        If Len(.FileName) > 0 Then
            lstAttachments.AddItem .FileName
        End If
    End With

End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdNew_Click()

    txtRecipient = ""
    txtSubject = ""
    txtMessage = ""
    
End Sub

Private Sub cmdRemove_Click()

    On Error Resume Next
    
    lstAttachments.RemoveItem lstAttachments.ListIndex

End Sub

Private Sub cmdSend_Click()
    '
    Dim i As Integer
    '
    'prepare attachments
    '
    For i = 0 To lstAttachments.ListCount - 1
        lstAttachments.ListIndex = i
        m_strEncodedFiles = m_strEncodedFiles & _
                         UUEncodeFile(lstAttachments.Text) & vbCrLf
    Next i
    '
    Winsock1.Connect Trim$(txtHost), 25
    m_State = MAIL_CONNECT
    '
End Sub

Private Sub Form_Load()
    '
    'clear all textboxes
    '
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_colAttachments = Nothing
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    '
    'Retrive data from winsock buffer
    '
    Winsock1.GetData strServerResponse
    '
    Debug.Print strServerResponse
    '
    'Get server response code (first three symbols)
    '
    strResponseCode = Left(strServerResponse, 3)
    '
    'Only these three codes tell us that previous
    'command accepted successfully and we can go on
    '
    If strResponseCode = "250" Or _
       strResponseCode = "220" Or _
       strResponseCode = "354" Then
       
        Select Case m_State
            Case MAIL_CONNECT
                'Change current state of the session
                m_State = MAIL_HELO
                '
                'Remove blank spaces
                strDataToSend = Trim$(txtSender)
                '
                'Retrieve mailbox name from e-mail address
                strDataToSend = Left$(strDataToSend, _
                                InStr(1, strDataToSend, "@") - 1)
                'Send HELO command to the server
                Winsock1.SendData "HELO " & strDataToSend & vbCrLf
                '
                Debug.Print "HELO " & strDataToSend
                '
            Case MAIL_HELO
                '
                'Change current state of the session
                m_State = MAIL_FROM
                '
                'Send MAIL FROM command to the server
                Winsock1.SendData "MAIL FROM:" & Trim$(txtSender) & vbCrLf
                '
                Debug.Print "MAIL FROM:" & Trim$(txtSender)
                '
            Case MAIL_FROM
                '
                'Change current state of the session
                m_State = MAIL_RCPTTO
                '
                'Send RCPT TO command to the server
                Winsock1.SendData "RCPT TO:" & Trim$(txtRecipient) & vbCrLf
                '
                Debug.Print "RCPT TO:" & Trim$(txtRecipient)
                '
            Case MAIL_RCPTTO
                '
                'Change current state of the session
                m_State = MAIL_DATA
                '
                'Send DATA command to the server
                Winsock1.SendData "DATA" & vbCrLf
                '
                Debug.Print "DATA"
                '
            Case MAIL_DATA
                '
                'Change current state of the session
                m_State = MAIL_DOT
                '
                'So now we are sending a message body
                'Each line of text must be completed with
                'linefeed symbol (Chr$(10) or vbLf) not with vbCrLf
                '
                'Send Subject line
                Winsock1.SendData "Subject:" & txtSubject & vbLf & vbCrLf
                '
                Debug.Print "Subject:" & txtSubject
                '
                Dim varLines    As Variant
                Dim varLine     As Variant
                Dim strMessage  As String
                '
                'Add atacchments
                strMessage = txtMessage & vbCrLf & vbCrLf & m_strEncodedFiles
                'clear memory
                m_strEncodedFiles = ""
                'Parse message to get lines (for VB6 only)
                varLines = Split(strMessage, vbCrLf)
                'clear memory
                strMessage = ""
                '
                'Send each line of the message
                For Each varLine In varLines
                    Winsock1.SendData CStr(varLine) & vbLf
                    '
                    Debug.Print CStr(varLine)
                Next
                '
                'Send a dot symbol to inform server
                'that sending of message comleted
                Winsock1.SendData "." & vbCrLf
                '
                Debug.Print "."
                '
            Case MAIL_DOT
                'Change current state of the session
                m_State = MAIL_QUIT
                '
                'Send QUIT command to the server
                Winsock1.SendData "QUIT" & vbCrLf
                '
                Debug.Print "QUIT"
            Case MAIL_QUIT
                '
                'Close connection
                Winsock1.Close
                '
        End Select
       
    Else
        '
        'If we are here server replied with
        'unacceptable respose code therefore we need
        'close connection and inform user about problem
        '
        Winsock1.Close
        '
        If Not m_State = MAIL_QUIT Then
            MsgBox "SMTP Error: " & strServerResponse, _
                    vbInformation, "SMTP Error"
        Else
            MsgBox "Message sent successfuly.", vbInformation
        End If
        '
    End If
    
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    MsgBox "Winsock Error number " & Number & vbCrLf & _
            Description, vbExclamation, "Winsock Error"

End Sub
