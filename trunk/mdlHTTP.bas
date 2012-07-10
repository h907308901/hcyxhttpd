Attribute VB_Name = "mdlHTTP"
Option Explicit

Public Type HTTP_REQUEST 'ParseRequest returning type
    sFile As String 'the file name to get (maybe directory name if allowed)
End Type
'===Example===
'GET / HTTP/1.1 => sFile
'Host: 127.0.0.1 => (not implemented)
'User-Agent: Mozilla/5.0 (Windows NT 5.1; rv:11.0) Gecko/20100101 Firefox/11.0 => (not implemented)
'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8 => (not implemented)
'Accept-Language: zh-cn,zh;q=0.8,en-us;q=0.5,en;q=0.3 => (not implemented)
'Accept -Encoding: gzip , deflate => (not implemented)
'Connection: keep -alive => (not implemented)
'=============

'===DO NOT TRY TO CHANGE THEM IN YOUR WAY===
Public g_lPort As Long 'the HTTP port
Public g_lMaxConnections As Long 'the max connections from clients
Public g_lConnections As Long 'the number of current connections
Public g_sRootDirectory As String 'the HTTP root directong on local (without "\" at end)
'===DO NOT TRY TO CHANGE THEM IN YOUR WAY===
Public Const HTTP_TRANSFER_BYTES_PER_TIME = &H1000 '0 'the number of transfer bytes per time in TransferFile
Public Const HTTP_DEFAULT_404_PAGE = "<title>404 Not Found</title><font size=30>404 Not Found</font>" 'if the user defined 404 page is not found, use the default page
Public Const HTTP_INDEX_PAGE = "index.html|index.htm" 'TODO: read settings from config

'Name: HttpOpen
'Function: Start HTTP service
'Parameters:
'   sRootDirectory(String): specify the root directory (if end up "\", HttpOpen will delete it automatically)
'   lMaxConnections(Long): specify the max connections
'   lPort(Long)(Optional): specify the HTTP port (80 default)
'Return: None
Sub HttpOpen(sRootDirectory As String, ByVal lMaxConnections As Long, Optional ByVal lPort As Long = 80)
    Dim i As Long
    With frmSocket
        'initialize primary socket to receive requests
        .wsk(0).LocalPort = lPort
        .wsk(0).Listen
        'initialize enough sockets to accept requests
        For i = 1 To lMaxConnections
            Load .wsk(i)
            .wsk(i).LocalPort = 0
        Next
    End With
    g_lPort = lPort
    g_lMaxConnections = lMaxConnections
    If Right$(sRootDirectory, 1) = "\" Then g_sRootDirectory = Left$(sRootDirectory, Len(sRootDirectory) - 1) Else g_sRootDirectory = sRootDirectory
End Sub

'Name: HttpClose
'Function: Stop HTTP service
'Parameters: None
'Return: None
Sub HttpClose()
    Dim i As Long
    With frmSocket
        .wsk(0).Close 'close primary socket
        'unload other sockets in order to release memory
        For i = 1 To .wsk.ubound
            Unload .wsk(i)
        Next
    End With
End Sub

'Name: ParseRequest
'Function: Parse the request string from client
'Parameters:
'   sRequest(String): the request string from client
'Return: (HTTP_REQUEST)a HTTP_REQUEST struct
Function ParseRequest(sRequest As String) As HTTP_REQUEST
    Dim i As Long, j As Long
    'parse GET
    i = InStr(sRequest, "GET ") 'get start position
    i = i + 4 'skip 4 chars "GET "
    j = InStr(i, sRequest, " ") 'get end position
    j = j - i 'length
    ParseRequest.sFile = Mid$(sRequest, i, j)
End Function

'Name: SlashConvert
'Function: Convert all "/" to "\"
'Parameters:
'   s(String): the string to convert
'Return: (String)the converted string
Function SlashConvert(s As String) As String
    Dim i As Long, s2 As String
    For i = 1 To Len(s)
        s2 = Mid$(s, i, 1)
        If s2 = "/" Then
            SlashConvert = SlashConvert & "\"
        Else
            SlashConvert = SlashConvert & s2
        End If
    Next
End Function

'Name: TransferFile
'Function: Transfer a file to a client
'Parameters:
'   sRequestedFile(String): the requested file by client
'   socket(Winsock): the socket to send data
'   bNotFound(Boolean)(Optional): internal, used by TransferFile
'Return: (Long)HTTP code
Function TransferFile(ByVal sRequestedFile As String, ByVal socket As Winsock, Optional ByVal bNotFound As Boolean) As Long
    Dim s As String
    Dim s2 As String
    Dim sFileList As String
    Dim i As Long, sPages() As String, bIndex As Boolean
    s = g_sRootDirectory & SlashConvert(sRequestedFile)
    If PathFileExists(s) <> 0 And PathIsDirectory(s) = 0 Then 'file exists
        SendFile s, socket
        TransferFile = 200
    Else
        If PathIsDirectory(s) <> 0 Then
'TODO:      read settings from config
            If Right$(s, 1) <> "\" Then
                s = s & "\"
                sRequestedFile = sRequestedFile & "/"
            End If
            sPages = Split(HTTP_INDEX_PAGE, "|")
            For i = 0 To UBound(sPages)
                If PathFileExists(s & sPages(i)) <> 0 And PathIsDirectory(s & sPages(i)) = 0 Then
                    bIndex = True
                    Exit For
                End If
            Next
            If bIndex Then
                TransferFile = TransferFile(sRequestedFile & sPages(i), socket)
            Else
                'generate file list
                sFileList = "<title>" & sRequestedFile & "</title><font size=30>Index of " & sRequestedFile & "</font><br><br>"
                s2 = Dir$(s & "*", vbNormal Or vbDirectory Or vbReadOnly)
                Do While s2 <> vbNullString
                    sFileList = sFileList & "<a href=" & sRequestedFile & s2 & ">" & s2 & "</a><br>"
                    s2 = Dir$()
                Loop
                SendHead 200, Len(sFileList), socket
                socket.SendData sFileList
                TransferFile = 200
            End If
        Else
            'not exist
            If bNotFound Then
                SendHead 404, Len(HTTP_DEFAULT_404_PAGE), socket
                socket.SendData HTTP_DEFAULT_404_PAGE
            Else
                TransferFile "/404.html", socket, True 'TODO: read settings from config
                TransferFile = 404
            End If
        End If
    End If
End Function

'Name: SendHead
'Function: Send a http reply head
'Parameters:
'   lCode(Long): HTTP status code, such as 200(ok), 403(forbidden), 404(not found), etc.
'   lPageLen(Long): the length of page
'   socket(Winsock): the socket to send head
'Return: None
Sub SendHead(ByVal lCode As Long, ByVal lPageLen As Long, ByVal socket As Winsock)
    socket.SendData "HTTP/1.1 " & CStr(lCode) & vbCrLf & "Content-Length:" & CStr(lPageLen) & vbCrLf & vbCrLf
End Sub

'Name: SendFile
'Function: Send a file to a client (called by TransferFile)
'Parameters:
'   sFile(String): the file to send, must exist
'   socket(Winsock): the socket to send data
'
'Return: None
Sub SendFile(sFile As String, ByVal socket As Winsock)
    Dim f As Long, l As Long
    Dim head As String
    Dim buff(HTTP_TRANSFER_BYTES_PER_TIME) As Byte
    f = FreeFile
    Open sFile For Binary As #f
    l = LOF(f) 'length of file
    SendHead 200, l, socket 'send head
    Do
        Get #f, , buff 'read data
        socket.SendData buff 'send data
        If l < HTTP_TRANSFER_BYTES_PER_TIME Then Exit Do 'send complete
        l = l - HTTP_TRANSFER_BYTES_PER_TIME
    Loop
    Close #f
End Sub
