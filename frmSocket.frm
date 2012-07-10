VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmSocket 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   2505
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock wsk 
      Index           =   0
      Left            =   600
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Name: FindFreeSock
'Function: Find a free socket to accept a client
'Parameters: None
'Return: (Long) The index of a free socket, if failed return -1
Private Function FindFreeSock() As Long
    Dim i As Long
    For i = 1 To wsk.UBound
        If wsk(i).State = sckClosed Then
            FindFreeSock = i
            Exit Function
        End If
    Next
    FindFreeSock = -1
End Function

Private Sub wsk_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim i As Long
    If Index = 0 Then 'a new client connected
        i = FindFreeSock 'find a free socket
        If i = -1 Then 'no free socket
            'add a new socket
            i = wsk.UBound + 1
            Load wsk(i)
        End If
        'frmMain.ListAdd "socket" & CStr(Index) & ":" & wsk(Index).RemoteHostIP & " connected" 'debug
        wsk(i).Accept requestID 'accept the connection
        If g_lConnections >= g_lMaxConnections Then 'over the max connections
'TODO:      reject (too many users)
        Else
'TODO:      reject if blocked
            g_lConnections = g_lConnections + 1 'increase the connections
        End If
    End If
End Sub

Private Sub wsk_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim buff As String '() As Byte
    Dim c As Long
    Dim r As HTTP_REQUEST
    'ReDim buff(bytesTotal - 1) As Byte
    buff = String(bytesTotal, vbNullChar) 'fill the buffer
    wsk(Index).GetData buff 'get data
    'Form1.lst.AddItem CStr(wsk(Index).RemoteHostIP) & " data:" & buff 'debug
    r = ParseRequest(buff) 'parse the request
    'frmMain.ListAdd "socket" & CStr(Index) & ":" & wsk(Index).RemoteHostIP & " GET " & r.sFile 'debug
    c = TransferFile(r.sFile, wsk(Index))
    'frmMain.ListAdd "socket" & CStr(Index) & ":" & wsk(Index).RemoteHostIP & " HTTP Code " & c 'debug
End Sub

Private Sub wsk_SendComplete(Index As Integer)
    'frmMain.ListAdd "socket" & CStr(Index) & ":" & wsk(Index).RemoteHostIP & " closed"  'debug
    wsk(Index).Close 'close socket to be free
    g_lConnections = g_lConnections - 1 'decrease the connections
End Sub
