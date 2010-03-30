Attribute VB_Name = "MUUEncode"

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
    lEncodedLines = lFileSize \ 45 + 1
    '
    'Prepare buffer to retrieve data from
    'the file by 45 symbols chunks
    strFileData = Space(45)
    '
    intFile = FreeFile
    '
    Open strFilePath For Binary As intFile
        For i = 1 To lEncodedLines
            'Read file data by 45-bytes cnunks
            '
            If i = lEncodedLines Then
                'Last line of encoded data often is not
                'equal to 45, therefore we need to change
                'size of the buffer
                strFileData = Space(lFileSize Mod 45)
            End If
            'Retrieve data chunk from file to the buffer
            Get intFile, , strFileData
            'Add first symbol to encoded string that informs
            'about quantity of symbols in encoded string.
            'More often "M" symbol is used.
            strTempLine = Chr(Len(strFileData) + 32)
            '
            If i = lEncodedLines And (Len(strFileData) Mod 3) Then
                'If the last line is processed and length of
                'source data is not a number divisible by 3, add one or two
                'blankspace symbols
                strFileData = strFileData + Space(3 - (Len(strFileData) Mod 3))
            End If
            
            For j = 1 To Len(strFileData) Step 3
                'Breake each 3 (8-bits) bytes to 4 (6-bits) bytes
                '
                '1 byte
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j, 1)) \ 4 + 32)
                '2 byte
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j, 1)) Mod 4) * 16 _
                               + Asc(Mid(strFileData, j + 1, 1)) \ 16 + 32)
                '3 byte
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j + 1, 1)) Mod 16) * 4 _
                               + Asc(Mid(strFileData, j + 2, 1)) \ 64 + 32)
                '4 byte
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j + 2, 1)) Mod 64 + 32)
            Next j
            'replace " " with "`"
            strTempLine = Replace(strTempLine, " ", "`")
            'add encoded line to result buffer
            strResult = strResult + strTempLine + vbLf
            'reset line buffer
            strTempLine = ""
        Next i
    Close intFile

    'add the end marker
    strResult = strResult & "`" & vbLf + "end" + vbLf
    'asign return value
    UUEncodeFile = strResult
    
End Function
