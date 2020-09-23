VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32_Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************
' Created By: Gary Noble - gwnoble@msn.com 13-05-2003
'*******************************************************************
Option Explicit

Dim objWMIService As Object
Dim colItems As Variant
Dim objItem As Object
Dim Locator As Object

Private mSystem                             As LibWIN32WMI
Private mvar_Win32__Account_SID             As Win32_Account_SID

Public Property Set SID(ByVal vData As Win32_Account)
    
    Set mvar_Win32__Account_SID = vData

End Property

Public Property Get SID() As Win32_Account_SID
  If (mvar_Win32__Account_SID Is Nothing) Then Set mvar_Win32__Account_SID = New Win32_Account_SID
    
  Set mvar_Win32__Account_SID.System = Me.System ' /* Always maintain a reference to the correct parent. */
  Set SID = mvar_Win32__Account_SID
  
End Property

Public Property Get System() As LibWIN32WMI
  Set System = mSystem
End Property

Friend Property Set System(ByRef value As LibWIN32WMI)
  Set mSystem = value
End Property

Public Function Fetch()
On Error GoTo ErrorHandler:

Dim Property As Object

Const SubSystem = "Win32_Account"

Dim Locator As Object
Set Locator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = Locator.ConnectServer(System.ComputerName, "root\cimv2", System.UserName, System.UserPassword)
objWMIService.Security_.ImpersonationLevel = 3


'Set objWMIService = GetObject("winmgmts:\\" & System.ComputerName & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from " & SubSystem, , 48)

Dim i As Long
For Each objItem In colItems
    i = i + 1
    mSystem.ProgMax = i
    DoEvents
Next

Set colItems = objWMIService.ExecQuery("Select * from " & SubSystem, , 48)
For Each objItem In colItems
    For Each Property In objItem.Properties_
        
        DoEvents
        Dim PropVal As String
        If IsNull(Property.value) Then
            PropVal = ""
        ElseIf IsArray(Property.value) Then
            PropVal = Join(Property.value, ",")
        Else
            PropVal = Property.value
        End If
  
        mSystem.AddressReturnData objItem.Path_.RelPath, System.ComputerID, System.ComputerName, SubSystem, Property.Name, PropVal

        Next
        mSystem.UpdateProgress
Next

Set objWMIService = Nothing
Set colItems = Nothing

ErrorHandler:
Select Case Err.Number
    Case "462" 'Server Doesn't Exist
        MsgBox Err.Description
        Exit Function
    Case "70" 'permission denied
        MsgBox Err.Description
        Exit Function
    Case Else
        Resume Next
End Select
End Function

Private Sub Class_Terminate()
  If Not (mvar_Win32__Account_SID Is Nothing) Then Set mvar_Win32__Account_SID = Nothing

End Sub

