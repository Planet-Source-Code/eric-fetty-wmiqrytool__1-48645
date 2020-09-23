VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32_Network"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*******************************************************************
' Created By: Gary Noble - gwnoble@msn.com 13-05-2003
' Name: Newtwork Wrapper
' Purpose: Interface To The Network WMI Calls
'          Use This Class To Interface All Network Call
'*******************************************************************
Option Explicit

Private mvar_Win32__NetworkAdapter As Win32_NetworkAdapter
Private mvar_Win32__NetworkAdapterConfig As Win32_NetworkAdapterConfig
Private mvar_Win32__NetworkClient As Win32_NetworkClient
Private mvar_Win32__NetworkConnection As Win32_NetworkConnection
Private mvar_Win32__NetworkLoginProfile As Win32_NetworkLoginProfile
Private mvar_Win32__NetworkProtocol As Win32_NetworkProtocol


Private mSystem                             As LibWIN32WMI

Public Property Get System() As LibWIN32WMI
  Set System = mSystem
End Property

Friend Property Set System(ByRef value As LibWIN32WMI)
  Set mSystem = value
End Property

Public Property Set Adapter(ByVal vData As Win32_NetworkAdapter)
    Set mvar_Win32__NetworkAdapter = vData
End Property

Public Property Get Adapter() As Win32_NetworkAdapter
  If (mvar_Win32__NetworkAdapter Is Nothing) Then Set mvar_Win32__NetworkAdapter = New Win32_NetworkAdapter
  Set mvar_Win32__NetworkAdapter.System = mSystem                 ' /* Always maintain a reference to the correct parent. */
  Set Adapter = mvar_Win32__NetworkAdapter
End Property

Public Property Set AdapterConfig(ByVal vData As Win32_NetworkAdapterConfig)
    Set mvar_Win32__NetworkAdapterConfig = vData
End Property

Public Property Get AdapterConfig() As Win32_NetworkAdapterConfig
  If (mvar_Win32__NetworkAdapterConfig Is Nothing) Then Set mvar_Win32__NetworkAdapterConfig = New Win32_NetworkAdapterConfig
  Set mvar_Win32__NetworkAdapterConfig.System = mSystem                  ' /* Always maintain a reference to the correct parent. */
  Set AdapterConfig = mvar_Win32__NetworkAdapterConfig
End Property


Public Property Set Client(ByVal vData As Win32_NetworkClient)
    Set mvar_Win32__NetworkClient = vData
End Property

Public Property Get Client() As Win32_NetworkClient
  If (mvar_Win32__NetworkClient Is Nothing) Then Set mvar_Win32__NetworkClient = New Win32_NetworkClient
  Set mvar_Win32__NetworkClient.System = mSystem                  ' /* Always maintain a reference to the correct parent. */
  Set Client = mvar_Win32__NetworkClient
End Property


Public Property Set Connection(ByVal vData As Win32_NetworkConnection)
    Set mvar_Win32__NetworkConnection = vData
End Property

Public Property Get Connection() As Win32_NetworkConnection
  If (mvar_Win32__NetworkConnection Is Nothing) Then Set mvar_Win32__NetworkConnection = New Win32_NetworkConnection
  Set mvar_Win32__NetworkConnection.System = mSystem                  ' /* Always maintain a reference to the correct parent. */
  Set Connection = mvar_Win32__NetworkConnection
End Property

Public Property Set LoginProfile(ByVal vData As Win32_NetworkLoginProfile)
    Set mvar_Win32__NetworkLoginProfile = vData
End Property

Public Property Get LoginProfile() As Win32_NetworkLoginProfile
  If (mvar_Win32__NetworkLoginProfile Is Nothing) Then Set mvar_Win32__NetworkLoginProfile = New Win32_NetworkLoginProfile
  Set mvar_Win32__NetworkLoginProfile.System = mSystem                  ' /* Always maintain a reference to the correct parent. */
  Set LoginProfile = mvar_Win32__NetworkLoginProfile
End Property


Public Property Set Protocol(ByVal vData As Win32_NetworkProtocol)
    Set mvar_Win32__NetworkProtocol = vData
End Property

Public Property Get Protocol() As Win32_NetworkProtocol
  If (mvar_Win32__NetworkProtocol Is Nothing) Then Set mvar_Win32__NetworkProtocol = New Win32_NetworkProtocol
  Set mvar_Win32__NetworkProtocol.System = mSystem                  ' /* Always maintain a reference to the correct parent. */
  Set Protocol = mvar_Win32__NetworkProtocol
End Property


Private Sub Class_Terminate()
  If Not (mvar_Win32__NetworkAdapter Is Nothing) Then Set mvar_Win32__NetworkAdapter = Nothing
  If Not (mvar_Win32__NetworkAdapterConfig Is Nothing) Then Set mvar_Win32__NetworkAdapterConfig = Nothing
  If Not (mvar_Win32__NetworkClient Is Nothing) Then Set mvar_Win32__NetworkClient = Nothing
  If Not (mvar_Win32__NetworkConnection Is Nothing) Then Set mvar_Win32__NetworkConnection = Nothing
  If Not (mvar_Win32__NetworkLoginProfile Is Nothing) Then Set mvar_Win32__NetworkLoginProfile = Nothing
  If Not (mvar_Win32__NetworkProtocol Is Nothing) Then Set mvar_Win32__NetworkProtocol = Nothing
End Sub

