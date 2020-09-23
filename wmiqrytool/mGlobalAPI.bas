Attribute VB_Name = "mGlobalAPI"
Option Explicit

Const LVM_FIRST As Long = &H1000
Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Const LVSCW_AUTOSIZE As Long = -1
Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long


Private Const WSA_NoName = "Unknown"
Private Const WSADescription_Len As Long = 256
Private Const WSASYS_Status_Len As Long = 128
Private Const WS_VERSION_REQD As Long = &H101
Private Const IP_SUCCESS As Long = 0
Private Const SOCKET_ERROR As Long = -1
Private Const AF_INET As Long = 2

Private Type WSADATA
  wVersion As Integer
  wHighVersion As Integer
  szDescription(0 To WSADescription_Len) As Byte
  szSystemStatus(0 To WSASYS_Status_Len) As Byte
  iMaxSockets As Integer
  imaxudp As Integer
  lpszvenderinfo As Long
End Type

Public Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long

Private Declare Function WSAStartup Lib "wsock32" _
  (ByVal VersionReq As Long, _
   WSADataReturn As WSADATA) As Long
  
Private Declare Function WSACleanup Lib "wsock32" () As Long

Private Declare Function inet_addr Lib "wsock32" _
  (ByVal s As String) As Long

Private Declare Function gethostbyaddr Lib "wsock32" _
  (haddr As Long, _
   ByVal hnlen As Long, _
   ByVal addrtype As Long) As Long

Private Declare Sub CopyMemory Lib "KERNEL32" _
   Alias "RtlMoveMemory" _
  (xDest As Any, _
   xSource As Any, _
   ByVal nbytes As Long)
   
Private Declare Function lstrlen Lib "KERNEL32" _
   Alias "lstrlenA" _
  (lpString As Any) As Long

Function CarDot(txt As String) As String
    Dim i As Integer
    i = InStr(txt, ".")
    If i = 0 Then i = 255
    CarDot = Left$(txt, i - 1)
End Function

Function CdrDot(txt As String) As String
    Dim i As Integer
    i = InStr(txt, ".")
    If i = 0 Then
        CdrDot = ""
    Else
        CdrDot = mID$(txt, i + 1)
    End If
End Function

Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   
   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
    
End Function


Public Function GetLocalHostName() As String

    Dim sName$

    sName = String(256, 0)

    If gethostname(sName, 256) Then

        sName = WSA_NoName

    Else

        If InStr(sName, Chr(0)) Then

            sName = Left(sName, InStr(sName, Chr(0)) - 1)

        End If

    End If

    GetLocalHostName = sName

End Function


Function IP2Hex(IPAddress As String) As String
'Converts in IP address in dot decimal notation
'  to a hexidecimal string.  Useful for sorting IP Addresses.
    Dim bytes(4) As Integer
    Dim buffer As String
    Dim i As Integer
    
    i = InStr("/", IPAddress)
    If i > 0 Then IPAddress = Left(IPAddress, i - 1)
    
    bytes(1) = Val(CarDot(IPAddress))
    buffer = CdrDot(IPAddress)
    bytes(2) = Val(CarDot(buffer))
    buffer = CdrDot(buffer)
    bytes(3) = Val(CarDot(buffer))
    bytes(4) = Val(CdrDot(buffer))
    buffer = ""
    For i = 1 To 4
        
'        If bytes(i) < 16 Then buffer = buffer & "0"
                
        buffer = buffer & Right("00" & Hex(bytes(i)), 2)
    Next i
    IP2Hex = buffer
    
End Function

Function Hex2IP(HexIP As String) As String
    Dim bytes(4) As Integer
    Dim i As Integer
    Dim StrLen As Integer
    StrLen = Len(HexIP)
    'pad with zeros
    For i = StrLen To 7
        HexIP = "0" & HexIP
    Next i
      
    For i = 0 To 3
        bytes(i + 1) = Val("&H" & mID$(HexIP, 2 * i + 1, 2))
    Next i
    
    Hex2IP = bytes(1) & "." & bytes(2) & "." & bytes(3) & "." & bytes(4)
    
End Function

Function IP2Int(IPAddress As String) As Long
    Dim HexIP As String
    HexIP = IP2Hex(IPAddress)
    IP2Int = Val("&h" & HexIP)
End Function

Function Int2IP(intIP As Long) As String
    
    Int2IP = Hex2IP(Hex(intIP))

End Function

Function Valid_IP(IP As String) As Boolean
    Dim l_IP As Long
    On Error Resume Next
    l_IP = IP2Int(IP)
    If Trim(IP) <> Int2IP(l_IP) Then
        Valid_IP = False
    Else
        Valid_IP = True
    End If
End Function

Public Function GetHostNameFromIP(ByVal sAddress As String) As String

   Dim ptrHosent As Long
   Dim hAddress As Long
   Dim nbytes As Long
   
   If SocketsInitialize() Then

     'convert string address to long
      hAddress = inet_addr(sAddress)
      
      If hAddress <> SOCKET_ERROR Then
         
        'obtain a pointer to the HOSTENT structure
        'that contains the name and address
        'corresponding to the given network address.
            DoEvents
         ptrHosent = gethostbyaddr(hAddress, 4, AF_INET)
            DoEvents
         If ptrHosent <> 0 Then
         
           'convert address and
           'get resolved hostname
            CopyMemory ptrHosent, ByVal ptrHosent, 4
            nbytes = lstrlen(ByVal ptrHosent)
         
            If nbytes > 0 Then
               sAddress = Space$(nbytes)
               CopyMemory ByVal sAddress, ByVal ptrHosent, nbytes
               
               GetHostNameFromIP = sAddress
            End If
         
         'Else: MsgBox "Call to gethostbyaddr failed."
         Else
            GetHostNameFromIP = sAddress
         End If 'If ptrHosent
      
      SocketsCleanup
      
      Else: MsgBox "String passed is an invalid IP."
      End If 'If hAddress
   
   Else: MsgBox "Sockets failed to initialize."
   End If  'If SocketsInitialize
      
End Function



Public Sub SocketsCleanup()
   
   If WSACleanup() <> 0 Then
       MsgBox "Windows Sockets error occurred in Cleanup.", vbExclamation
   End If
    
End Sub






Public Sub AutosizeColumns(ListViewControl As ListView)
' This sub resizes the columns in a ListView: based on a sample code by Barcode (Andy D.)
Dim vColumn As Variant
Dim iColumn As Byte

    ' Resize each column in the ListView
    For Each vColumn In ListViewControl.ColumnHeaders
        ' Lock the ListView area
        LockWindowUpdate ListViewControl.hwnd
        ' Autosize the column
        SendMessage ListViewControl.hwnd, LVM_SETCOLUMNWIDTH, iColumn, LVSCW_AUTOSIZE_USEHEADER
        ' Release the ListView area
        LockWindowUpdate 0
        ' Update the column number
        iColumn = iColumn + 1
    Next
End Sub

