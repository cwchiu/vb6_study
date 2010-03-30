Attribute VB_Name = "M_UUCode"
Public Function UUDecodeToFile(strUUCodeData As String, strFilePath As String)

    Dim vDataLine   As Variant
    Dim vDataLines  As Variant
    Dim strDataLine As String
    Dim intSymbols  As Integer
    Dim intFile     As Integer
    Dim strTemp     As String
    
    If Left$(strUUCodeData, 6) = "begin " Then
        strUUCodeData = Mid$(strUUCodeData, InStr(1, strUUCodeData, vbLf) + 1)
    End If
    
    If Right$(strUUCodeData, 4) = "end" + vbLf Then
        strUUCodeData = Left$(strUUCodeData, Len(strUUCodeData) - 7)
    End If
    
    intFile = FreeFile
    Open strFilePath For Binary As intFile
    
        vDataLines = Split(strUUCodeData, vbLf)
        
        For Each vDataLine In vDataLines
                strDataLine = CStr(vDataLine)
                intSymbols = Asc(Left$(strDataLine, 1))
                strDataLine = Mid$(strDataLine, 2, intSymbols)
                For i = 1 To Len(strDataLine) Step 4
                    strTemp = strTemp + Chr((Asc(Mid(strDataLine, i, 1)) - 32) * 4 + _
                              (Asc(Mid(strDataLine, i + 1, 1)) - 32) \ 16)
                    strTemp = strTemp + Chr((Asc(Mid(strDataLine, i + 1, 1)) Mod 16) * 16 + _
                              (Asc(Mid(strDataLine, i + 2, 1)) - 32) \ 4)
                    strTemp = strTemp + Chr((Asc(Mid(strDataLine, i + 2, 1)) Mod 4) * 64 + _
                              Asc(Mid(strDataLine, i + 3, 1)) - 32)
                Next i
                Put intFile, , strTemp
                strTemp = ""
        Next
    
    Close intFile
    
End Function


Public Function UUEncodeFile(strFilePath As String) As String

    Dim intFile         As Integer      'file handler
    Dim intTempFile     As Integer      'temp file
    Dim lFileSize       As Long         'size of the file
    Dim strFileName     As String       'name of the file
    Dim strFileData     As String       'file data chunk
    Dim lEncodedLines   As Long         'number of encoded lines
    Dim strTempLine     As String       'temporary string
    Dim i               As Long         'loop counter
    Dim j               As Integer      'loop counter
    
    Dim strResult       As String
    '
    'Get file name
    strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
    '
    'Insert first marker: "begin 664 ..."
    strResult = "begin 664 " + strFileName + vbLf
    '
    'Get file size
    lFileSize = FileLen(strFilePath)
    lEncodedLines = lFileSize / 45 + 1
    '
    'Prepare buffer to retrieve data form
    'the file by 45 symbols chunks
    strFileData = Space(45)
    '
    intFile = FreeFile
    '
    
    Open strFilePath For Binary As intFile
        For i = 1 To lEncodedLines
            If i = lEncodedLines Then
                strFileData = Space(lFileSize Mod 45)
            End If
            Get intFile, , strFileData
            strTempLine = Chr(Len(strFileData) + 32)
            If i = lEncodedLines And (Len(strFileData) Mod 3) Then
                strFileData = strFileData + Space(3 - (Len(strFileData) Mod 3))
            End If
            
            For j = 1 To Len(strFileData) Step 3
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j, 1)) \ 4 + 32)
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j, 1)) Mod 4) * 16 _
                               + Asc(Mid(strFileData, j + 1, 1)) \ 16 + 32)
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j + 1, 1)) Mod 16) * 4 _
                               + Asc(Mid(strFileData, j + 2, 1)) \ 64 + 32)
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j + 2, 1)) Mod 64 + 32)
            Next j
            strResult = strResult + strTempLine + vbLf
            strTempLine = ""
        Next i
    Close intFile
    
    strResult = strResult & "'" & vbLf + "end" + vbLf
    UUEncodeFile = strResult
    
End Function
