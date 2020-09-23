VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Win32_Event_Log"
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

Private mvar_Win32__LogEvents As Win32_NTLogEvent
Private mvar_Win32__LogFile As Win32_NTEventlogFile

Private mSystem                             As LibWIN32WMI

Public Property Get System() As LibWIN32WMI
  Set System = mSystem
End Property

Friend Property Set System(ByRef value As LibWIN32WMI)
  Set mSystem = value
End Property


Public Property Set File(ByVal vData As Win32_NTEventlogFile)
    Set mvar_Win32__LogFile = vData
End Property

Public Property Get File() As Win32_NTEventlogFile
  If (mvar_Win32__LogFile Is Nothing) Then Set mvar_Win32__LogFile = New Win32_NTEventlogFile
  Set mvar_Win32__LogFile.System = mSystem                  ' /* Always maintain a reference to the correct parent. */
  Set File = mvar_Win32__LogFile
End Property


Public Property Set Events(ByVal vData As Win32_NetworkProtocol)
    Set mvar_Win32__LogEvents = vData
End Property

Public Property Get Events() As Win32_NTLogEvent
  If (mvar_Win32__LogEvents Is Nothing) Then Set mvar_Win32__LogEvents = New Win32_NTLogEvent
    Set mvar_Win32__LogEvents.System = mSystem                  ' /* Always maintain a reference to the correct parent. */
  Set Events = mvar_Win32__LogEvents
End Property


Private Sub Class_Terminate()
  If Not (mvar_Win32__LogEvents Is Nothing) Then Set mvar_Win32__LogEvents = Nothing
  If Not (mvar_Win32__LogFile Is Nothing) Then Set mvar_Win32__LogFile = Nothing
End Sub
