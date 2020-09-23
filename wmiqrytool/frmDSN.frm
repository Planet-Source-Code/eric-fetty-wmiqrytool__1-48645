VERSION 5.00
Begin VB.Form frmDSN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose the DSN to store results in"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit DSNs"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox cmbDSN 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "frmDSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aDSNs As Variant
Private Declare Function SQLManageDataSources Lib "ODBCCP32.DLL" (ByVal hwnd As Long) As Long   'Show ODBC Manager



Private Sub cmbDSN_Click()
If cmbDSN = "" Then
    cmdOK.Enabled = False
Else
    cmdOK.Enabled = True
End If

End Sub

Private Sub cmdEdit_Click()
SQLManageDataSources (Me.hwnd)
Form_Load
End Sub

Private Sub cmdOK_Click()
Call SaveSetting(App.Title, DEF_REGISTRY_SETTINGS, "DefaultDSN", cmbDSN, HKEY_LOCAL_MACHINE)
  
Me.Hide
End Sub

Private Sub Form_Load()
cmdOK.Enabled = False
cmbDSN.Clear
aDSNs = mODBC.GetDSNs()
For i = LBound(aDSNs) To UBound(aDSNs)
    cmbDSN.AddItem aDSNs(i), i
Next i

Dim DefaultDSN As String
DefaultDSN = GetSetting(App.Title, DEF_REGISTRY_SETTINGS, "DefaultDSN", "", HKEY_LOCAL_MACHINE)
    
For i = 0 To cmbDSN.ListCount - 1
    cmbDSN.ListIndex = i
    If cmbDSN = DefaultDSN Then Exit For
Next i

If cmbDSN <> DefaultDSN And cmbDSN.ListIndex = cmbDSN.ListCount Then cmbDSN.ListIndex = 0
End Sub
