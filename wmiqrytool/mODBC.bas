Attribute VB_Name = "mODBC"
Option Explicit

Private Const SQL_SUCCESS     As Long = 0        ' ODBC Success
Private Const SQL_ERROR       As Long = -1       ' ODBC Error
Private Const SQL_FETCH_NEXT  As Long = 1        ' ODBC Move Next
 
Private Declare Function SQLDataSources Lib "ODBC32.DLL" _
                  (ByVal henv As Long, ByVal fDirection _
                  As Integer, ByVal szDSN As String, _
                  ByVal cbDSNMax As Integer, pcbDSN As Integer, _
                  ByVal szDescription As String, ByVal cbDescriptionMax _
                  As Integer, pcbDescription As Integer) As Integer

Private Declare Function SQLAllocEnv Lib "ODBC32.DLL" _
                  (env As Long) As Integer
'
 
Public Function GetDSNs() As Variant

    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    ':::                                                      :::
    ':::  This routine does the actual work                   :::
    ':::                                                      :::
    '::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    '
    Dim intRetCode  As Integer    ' the return code
    Dim strDSNItem  As String     ' the dsn name
    Dim strDRVItem  As String     ' the driver name
    Dim strDSN      As String     ' the formatted dsn name
    Dim intDSNLen   As Integer    ' the length of the dsn name
    Dim intDRVLen   As Integer    ' the length of the driver name
    Dim henv        As Long       ' handle to the environment
    Dim strTemp     As String     ' Tempspace
    Dim strDSNTemp  As String     ' Tempspace
 
    On Error Resume Next
 
    If (SQLAllocEnv(henv) <> SQL_ERROR) Then
        Do
 
            strDSNItem = Space$(1024)
            strDRVItem = Space$(1024)
 
            intRetCode = SQLDataSources(henv, SQL_FETCH_NEXT, strDSNItem, _
                         Len(strDSNItem), intDSNLen, strDRVItem, _
                         Len(strDRVItem), intDRVLen)
 
            strDSN = Left$(strDSNItem, intDSNLen)
 
            If (Len(strDSN) > 0) And (strDSN <> Space$(intDSNLen)) Then
               strDSNTemp = strDSN & "|"
               ' Check for dupes...
               If InStr(strTemp, strDSNTemp) = 0 Then
                  strTemp = strTemp & strDSNTemp
               End If
            End If
 
        Loop Until intRetCode <> SQL_SUCCESS
    End If
    
    GetDSNs = Split(strTemp, "|")
    
End Function


