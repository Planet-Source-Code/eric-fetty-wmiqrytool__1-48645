VERSION 5.00
Begin VB.Form frmImprtDB 
   Caption         =   "Import Database"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4110
   ScaleWidth      =   8700
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCrit 
      Height          =   315
      Left            =   4440
      TabIndex        =   13
      Top             =   2760
      Width           =   3855
   End
   Begin VB.ComboBox cmbCrit 
      Height          =   315
      ItemData        =   "frmImprtDB.frx":0000
      Left            =   3600
      List            =   "frmImprtDB.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin VB.ComboBox cmbFilter 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CheckBox chkFilter 
      Caption         =   "Filter"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6600
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Default         =   -1  'True
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1800
      Width           =   3255
   End
   Begin VB.ComboBox cmbPrimKey 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1800
      Width           =   3255
   End
   Begin VB.ComboBox cmbTable 
      Height          =   315
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.ComboBox cmbDSN 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   8415
      Begin VB.Label lblWhere 
         Caption         =   "Where"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label lblData 
      Caption         =   "Select the node field to use:"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label lblPrimKey 
      Caption         =   "Select the ID field to use:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label lblTable 
      Caption         =   "Select the table to use:"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblDSN 
      Caption         =   "Select the DSN for the Database to use:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "frmImprtDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aDSNs As Variant
Public mCon As ADODB.Connection
Public mRS As ADODB.Recordset

Private Sub chkFilter_Click()
lblWhere.Enabled = chkFilter
cmbCrit.Enabled = chkFilter
cmbFilter.Enabled = chkFilter
txtCrit.Enabled = chkFilter
End Sub

Private Sub cmbDSN_Click()
Dim newTableName As String

On Error GoTo ErrorHandler:

cmbTable.Clear
cmbPrimKey.Clear
cmbData.Clear
cmbFilter.Clear

cmbDSN.BackColor = &H80000005
cmbTable.BackColor = &HC0FFC0
cmbPrimKey.BackColor = &H80000005
cmbData.BackColor = &H80000005
cmbDSN.Enabled = False
Set mCon = New ADODB.Connection

mCon.ConnectionString = "DSN=" & cmbDSN
mCon.Open
Set mRS = mCon.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))

Do While Not mRS.EOF
    If UCase(Left(mRS!Table_name, 4)) <> "MSYS" Then
        If UCase(Left(mRS!Table_name, 11)) <> "SWITCHBOARD" Then
            newTableName = mRS!Table_name
            cmbTable.AddItem newTableName
        End If
    End If
    mRS.MoveNext
Loop
cmbTable.Enabled = True

Exit Sub

ErrorHandler:

MsgBox Err.Description
Form_Load


End Sub

Private Sub cmbTable_Click()

On Error GoTo ErrorHandler:

cmbPrimKey.Clear
cmbData.Clear
cmbFilter.Clear

cmbDSN.BackColor = &H80000005
cmbTable.BackColor = &H80000005
cmbPrimKey.BackColor = &HC0FFC0
cmbData.BackColor = &H80000005

cmbTable.Enabled = False
Set mRS = New ADODB.Recordset
mRS.Open "SELECT * FROM " & cmbTable, mCon, adOpenDynamic, adLockReadOnly
For Each Field In mRS.Fields
    cmbPrimKey.AddItem Field.Name
    cmbData.AddItem Field.Name
    cmbFilter.AddItem Field.Name
Next
cmbPrimKey.Enabled = True

Exit Sub

ErrorHandler:

MsgBox Err.Description
Form_Load

End Sub

Private Sub cmbPrimKey_Click()

On Error GoTo ErrorHandler:

cmbDSN.BackColor = &H80000005
cmbTable.BackColor = &H80000005
cmbPrimKey.BackColor = &H80000005
cmbData.BackColor = &HC0FFC0

cmbPrimKey.Enabled = False
cmbData.Enabled = True

Exit Sub

ErrorHandler:

MsgBox Err.Description
Form_Load


End Sub

Private Sub cmbData_Click()

On Error GoTo ErrorHandler:

cmbDSN.BackColor = &H80000005
cmbTable.BackColor = &H80000005
cmbPrimKey.BackColor = &H80000005
cmbData.BackColor = &H80000005

cmbData.Enabled = False
cmdContinue.Enabled = True

Exit Sub

ErrorHandler:

MsgBox Err.Description
Form_Load

End Sub

Private Sub cmdCancel_Click()
cmdContinue.Enabled = False
cmbTable.Enabled = False
cmbPrimKey.Enabled = False
cmbData.Enabled = False
cmbDSN.Enabled = True
cmbDSN.BackColor = &HC0FFC0
cmbTable.BackColor = &H80000005
cmbPrimKey.BackColor = &H80000005
cmbData.BackColor = &H80000005
chkFilter = False
Me.Hide
End Sub

Private Sub cmdContinue_Click()
Form1.blnFromDB = True
Form1.txtComputer.Text = "Pulling Computers From Database"
Form1.txtComputer.Enabled = False
Set mRS = New ADODB.Recordset
If chkFilter Then
    If cmbCrit = "Like" Then
    mRS.Open "SELECT DISTINCT [" & cmbPrimKey & "], [" & cmbData & "] FROM " & cmbTable & " WHERE [" & cmbFilter & "] " & cmbCrit & " '%" & txtCrit & "%'", mCon, adOpenDynamic, adLockReadOnly
    Else
    mRS.Open "SELECT DISTINCT [" & cmbPrimKey & "], [" & cmbData & "] FROM " & cmbTable & " WHERE [" & cmbFilter & "] " & cmbCrit & " '" & txtCrit & "'", mCon, adOpenDynamic, adLockReadOnly
    End If
Else
    mRS.Open "SELECT DISTINCT [" & cmbPrimKey & "], [" & cmbData & "] FROM " & cmbTable, mCon, adOpenDynamic, adLockReadOnly
End If

Dim x As Integer
x = 0
Do Until mRS.EOF
x = x + 1
mRS.MoveNext
Loop

MsgBox "You have Selected " & x & " items to be polled", vbInformation

cmdContinue.Enabled = False
cmbTable.Enabled = False
cmbPrimKey.Enabled = False
cmbData.Enabled = False
cmbDSN.Enabled = True
cmbDSN.BackColor = &HC0FFC0
cmbTable.BackColor = &H80000005
cmbPrimKey.BackColor = &H80000005
cmbData.BackColor = &H80000005
chkFilter = False
Me.Hide
End Sub

Private Sub Form_Load()

Dim i As Integer

cmbCrit.Clear
cmbCrit.AddItem "="
cmbCrit.AddItem "Like"
cmbCrit.AddItem ">"
cmbCrit.AddItem "<"
cmbCrit.AddItem "<>"


chkFilter.Value = False
chkFilter_Click

cmdContinue.Enabled = False
cmbTable.Enabled = False
cmbPrimKey.Enabled = False
cmbData.Enabled = False
cmbDSN.Enabled = True
cmbDSN.BackColor = &HC0FFC0
cmbTable.BackColor = &H80000005
cmbPrimKey.BackColor = &H80000005
cmbData.BackColor = &H80000005
cmbDSN.Clear
cmbTable.Clear
cmbPrimKey.Clear
cmbData.Clear
cmbFilter.Clear

aDSNs = mODBC.GetDSNs()
For i = LBound(aDSNs) To UBound(aDSNs)
    cmbDSN.AddItem aDSNs(i), i
Next i

End Sub
