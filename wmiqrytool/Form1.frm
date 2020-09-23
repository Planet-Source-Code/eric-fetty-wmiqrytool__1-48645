VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WMI Query Tool"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8325
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select Data to Poll..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtComputer 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Text            =   "."
      Top             =   210
      Width           =   5895
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   6000
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Computer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "SubSystem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Detail"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblNotes 
      Alignment       =   2  'Center
      ForeColor       =   &H80000011&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   8055
   End
   Begin VB.Label Label2 
      Caption         =   "Computer to Poll:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Items Returned"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu OpenCSV 
         Caption         =   "Import CSV"
      End
      Begin VB.Menu ImportDB 
         Caption         =   "Import Database"
      End
      Begin VB.Menu EXIT 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu User 
      Caption         =   "User"
      Begin VB.Menu ChangeUser 
         Caption         =   "Change User"
      End
   End
   Begin VB.Menu Database 
      Caption         =   "Database"
      Begin VB.Menu EditDSN 
         Caption         =   "Edit DSN"
      End
      Begin VB.Menu ViewDatabase 
         Caption         =   "View Database"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mCon As ADODB.Connection
Dim aPollers() As Variant
Dim aNames() As Variant
Dim mRS As ADODB.Recordset
Public WithEvents m_WMIWrapper  As LibWIN32WMI
Attribute m_WMIWrapper.VB_VarHelpID = -1
Dim mType As Variant
Dim xItem As MSComctlLib.ListItem
Dim filename As String
Dim FH As Integer
Dim blnFromFile As Boolean
Public blnFromDB As Boolean
Private Declare Function SQLManageDataSources Lib "ODBCCP32.DLL" (ByVal hwnd As Long) As Long   'Show ODBC Manager



Private Sub About_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub ChangeUser_Click()

frmLogin.Show vbModal, Me
If m_WMIWrapper.UserName = "" Then
lblNotes = ""
Else
lblNotes = "User changed to: " & m_WMIWrapper.UserName
End If

End Sub

Private Sub cmdSelect_Click()
frmSelect.Show vbModal, Me
End Sub

Private Sub cmdStart_Click()
 Dim x As Integer
    
    txtComputer.Enabled = False
    File.Enabled = False
    User.Enabled = False
    Database.Enabled = False
    Help.Enabled = False
    cmdStart.Enabled = False
    cmdSelect.Enabled = False
    
    
    If blnFromDB = True Then
    PollDB
    txtComputer.Text = "Enter Computer to Scan"
    txtComputer.Enabled = True
    blnFromDB = False
    File.Enabled = True
    User.Enabled = True
    Database.Enabled = True
    Help.Enabled = True
    cmdStart.Enabled = True
    cmdSelect.Enabled = True
    Exit Sub
    End If
    
    If FH <> 0 Then
    If blnFromFile = True And (Not EOF(FH)) Then
        Dim Compname As String
        Line Input #FH, Compname
        txtComputer.Text = Compname
    End If
    Else
    Me.ListView1.ListItems.Clear
    End If
    
        
    If txtComputer.Text = "." Then
    txtComputer.Text = mGlobalAPI.GetLocalHostName()
    ElseIf mGlobalAPI.Valid_IP(txtComputer.Text) Then
        txtComputer.Text = mGlobalAPI.GetHostNameFromIP(txtComputer.Text)
    End If
    txtComputer.Text = UCase(txtComputer.Text)
    m_WMIWrapper.ComputerName = txtComputer.Text

    Screen.MousePointer = vbHourglass
      
    DoEvents
    
    LockWindowUpdate Me.ListView1.hwnd
      
    StartPolling
    
    Dim c As LibWIN32WMI
    
    
    
    AutosizeColumns Me.ListView1
    
   LockWindowUpdate ByVal 0&
    
    'combo1.Enabled = True
    File.Enabled = True
    User.Enabled = True
    Database.Enabled = True
    Help.Enabled = True
    cmdStart.Enabled = True
    cmdSelect.Enabled = True
    
    Screen.MousePointer = vbDefault
    If FH <> 0 Then
        If EOF(FH) And blnFromFile = True Then
            blnFromFile = False
            Close #FH
            FH = 0
            txtComputer.Text = "Enter Computer to Scan"
            txtComputer.Enabled = True
        ElseIf Not EOF(FH) And blnFromFile = True Then
            cmdStart_Click
        ElseIf blnFromFile = False Then
            'do nothing
        End If
    Else
        txtComputer.Enabled = True
        txtComputer.Text = "Enter Computer to Scan"
        
    End If
    
    
    
        
    
End Sub



Private Sub EditDSN_Click()
    SQLManageDataSources (Me.hwnd)
    Form_Load
End Sub

Private Sub EXIT_Click()
Unload Me
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler:
 Dim dbUser, dbPass As String
frmload:
cmdStart.Enabled = False
cmdSelect.Enabled = True
blnFromFile = False
blnFromDB = False
Set mCon = New ADODB.Connection

frmDSN.Show vbModal, Me
mCon.Open "DSN=" & frmDSN.cmbDSN, dbUser & "", dbPass & ""
txtComputer.Text = "Enter Computer to Scan"
Set m_WMIWrapper = New LibWIN32WMI
    
    Set m_WMIWrapper.ProgressBar = Me.ProgressBar1
    Dim y As Integer
    y = 0
    frmSelect.lstData.Clear
    For Each mType In m_WMIWrapper.WMICalls
        frmSelect.lstData.AddItem mType
        BuildArray mType, y
        ReDim Preserve aNames(y)
        aNames(y) = mType
        
        y = y + 1
    Next
    
Exit Sub
    
ErrorHandler:
Select Case Err.Number
    Case "-2147467259"
        MsgBox "PLEASE CREATE WMI DSN FIRST:" & vbCrLf & vbCrLf & "Create a System DSN with the title of ""WMI""," & vbCrLf & "point it to the database you wish to store your query results in.", vbCritical, "DSN Does Not Exist"
        SQLManageDataSources (Me.hwnd)
        Resume frmload:
    Case "-2147217843"
        frmLogin1.Show vbModal, Me
        dbUser = frmLogin1.txtUserName
        dbPass = frmLogin1.txtPassword
        Resume
    Case Else
    MsgBox Err.Number & vbCrLf & Err.Description
    Unload Me
End Select


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set m_WMIWrapper = Nothing
End Sub


Private Sub ImportDB_Click()
frmImprtDB.Show vbModal, Me
End Sub

Private Sub m_WMIWrapper_ItemDetail(ID As String, ComputerID As String, Computer As String, SubSystem As String, Header As String, Detail As String, PercentComplete As Long)
On Error GoTo ErrorHandler:
Start:
Dim DateTime As String
If ComputerID = Empty Then ComputerID = Computer
Set mRS = New ADODB.Recordset
mRS.Open "SELECT ID, COMPUTERID, DATETIMESTAMP, [" & Header & "] FROM " & SubSystem & " WHERE ID = '" & ID & "' AND COMPUTERID = '" & ComputerID & "'", mCon, adOpenDynamic, adLockPessimistic
DateTime = Now()
If mRS.EOF And mRS.BOF Then 'nothing with that ID, need to add new record
    mRS.AddNew
    mRS("ID") = ID
    mRS("COMPUTERID") = ComputerID
    mRS("DATETIMESTAMP") = DateTime
    mRS(Header) = Detail
    mRS.Update
Else 'Just update
    mRS("DATETIMESTAMP") = DateTime
    mRS(Header) = Detail
    mRS.Update
End If


Set xItem = Me.ListView1.ListItems.Add
    xItem.Text = Computer
    xItem.SubItems(1) = SubSystem
    xItem.SubItems(2) = Header
    xItem.SubItems(3) = Detail
    Label4.Caption = PercentComplete
    
mRS.Close
Set mRS = Nothing
Exit Sub

ErrorHandler:

Select Case Err.Number

    Case "-2147217865" 'Table doesn't exist
        If InStr(1, mCon.ConnectionString, "Access") > 0 Then 'This is an access database
            mCon.Execute "CREATE TABLE " & SubSystem & "(ID varchar(255), COMPUTERID varchar(255), DATETIMESTAMP DATETIME, [" & Header & "] varchar(255))"
        Else
            mCon.Execute "CREATE TABLE " & SubSystem & "(ID varchar(255), COMPUTERID varchar(255), DATETIMESTAMP DATETIME, [" & Header & "] varchar(255))"
        End If
    Case "3705" 'Recordset Already Open
        mRS.Close
    Case "-2147217904" 'Column Doesn't Exist access
        Dim Z As Integer
        Set mRS = mCon.Execute("SELECT * FROM " & SubSystem)
        For Z = 0 To mRS.Fields.Count - 1
        If mRS(Z).Name = "DATETIMESTAMP" Then Exit For
        Next Z
        If (Z = mRS.Fields.Count) And (mRS(Z - 1).Name <> "DATETIMESTAMP") Then
            Set mRS = New ADODB.Recordset
            mCon.Execute "ALTER TABLE " & SubSystem & " ADD DATETIMESTAMP DATETIME"
        Else
            Set mRS = New ADODB.Recordset
            If Len(Detail) <= 255 Then
            mCon.Execute "ALTER TABLE " & SubSystem & " ADD [" & Header & "] varchar(" & Len(Detail) + 1 & ")"
            Else
            mCon.Execute "ALTER TABLE " & SubSystem & " ADD [" & Header & "] memo"
            End If
        End If
    Case "-2147217900" 'Column Doesn't Exist sql
        If InStr(1, Err.Description, "DATETIMESTAMP") > 0 Then
            mCon.Execute "ALTER TABLE " & SubSystem & " ADD DATETIMESTAMP DATETIME"
        ElseIf InStr(1, Err.Description, "COMPUTERID") > 0 Then
            mCon.Execute "ALTER TABLE " & SubSystem & " ADD COMPUTERID varchar(255)"
        Else
            mCon.Execute "ALTER TABLE " & SubSystem & " ADD [" & Header & "] varchar(" & Len(Detail) + 1 & ")"
        End If
    Case "-2147217887" 'Column not big enough
        Set mRS = Nothing
     
        If InStr(1, mCon.ConnectionString, "Access") > 0 Then 'if MSAccess handle large fields with a memo field type
                    If Len(Detail) > 255 Then
                        mCon.Execute "ALTER TABLE " & SubSystem & " ALTER [" & Header & "] memo"
                    Else 'otherwise use carchar datatype
                        mCon.Execute "ALTER TABLE " & SubSystem & " ALTER [" & Header & "] varchar(" & Len(Detail) + 1 & ")"
                    End If
                    
        Else
        mCon.Execute "ALTER TABLE " & SubSystem & " ALTER COLUMN [" & Header & "] varchar(" & Len(Detail) + 1 & ")"
        End If
        Resume Start:
    Case Else
        MsgBox "New Error:" & vbCrLf & Err.Number & vbCrLf & Err.Description


End Select

Resume
End Sub

Private Sub OpenCSV_Click()
CommonDialog1.ShowOpen
filename = CommonDialog1.filename
If filename <> "" Then
    blnFromFile = True
    FH = FreeFile()
    Open filename For Input As #FH
    Me.ListView1.ListItems.Clear
    txtComputer.Enabled = False
    txtComputer.Text = "Using " & filename
End If
End Sub
Private Sub BuildArray(mType As Variant, y As Integer)
ReDim Preserve aPollers(y)
    Select Case y
        
        Case Is = 0
            Set aPollers(y) = m_WMIWrapper.wWin32__1394Controller
        Case Is = 1
            Set aPollers(y) = m_WMIWrapper.wWin32__BaseBoard
        Case Is = 2
            Set aPollers(y) = m_WMIWrapper.wWin32__Account
        Case Is = 3
            Set aPollers(y) = m_WMIWrapper.wWin32__Account.SID
        Case Is = 4
            Set aPollers(y) = m_WMIWrapper.wWin32__SoftwareFeature
        Case Is = 5
            Set aPollers(y) = m_WMIWrapper.wWin32__ApplicationService
        Case Is = 6
            Set aPollers(y) = m_WMIWrapper.wWin32__BaseService
        Case Is = 7
            Set aPollers(y) = m_WMIWrapper.wWin32__Battery
        Case Is = 8
            Set aPollers(y) = m_WMIWrapper.wWin32__Binary
        Case Is = 9
            Set aPollers(y) = m_WMIWrapper.wWin32__BindImageAction
        Case Is = 10
            Set aPollers(y) = m_WMIWrapper.wWin32__Bios
        Case Is = 11
            Set aPollers(y) = m_WMIWrapper.wWin32__BootConfig
        Case Is = 12
            Set aPollers(y) = m_WMIWrapper.wWin32__Bus
        Case Is = 13
            Set aPollers(y) = m_WMIWrapper.wWin32__ComputerSystem
        Case Is = 14
            Set aPollers(y) = m_WMIWrapper.wWin32__ComputerSystemProduct
        Case Is = 15
            Set aPollers(y) = m_WMIWrapper.wWin32__DiskDrive
        Case Is = 16
            Set aPollers(y) = m_WMIWrapper.wWin32__DiskPartition
        Case Is = 17
            Set aPollers(y) = m_WMIWrapper.wWin32__Network.Adapter
        Case Is = 18
            Set aPollers(y) = m_WMIWrapper.wWin32__Network.AdapterConfig
        Case Is = 19
            Set aPollers(y) = m_WMIWrapper.wWin32__Printer
        Case Is = 20
            Set aPollers(y) = m_WMIWrapper.wWin32__Printer.Configurations
        Case Is = 21
            Set aPollers(y) = m_WMIWrapper.wWin32__Printer.CurrentJobs
        Case Is = 22
            Set aPollers(y) = m_WMIWrapper.wWin32__Network.Client
        Case Is = 23
            Set aPollers(y) = m_WMIWrapper.wWin32__Network.Connection
        Case Is = 24
            Set aPollers(y) = m_WMIWrapper.wWin32__Network.LoginProfile
        Case Is = 25
            Set aPollers(y) = m_WMIWrapper.wWin32__Network.Protocol
        Case Is = 26
            Set aPollers(y) = m_WMIWrapper.wWin32__Eventlog.File
        Case Is = 27
            Set aPollers(y) = m_WMIWrapper.wWin32__Eventlog.Events
        Case Is = 28
            Set aPollers(y) = m_WMIWrapper.wWin32__OperatingSystem
        Case Is = 29
            Set aPollers(y) = m_WMIWrapper.wWin32__Process
        Case Is = 30
            Set aPollers(y) = m_WMIWrapper.wWin32__Processor
        Case Is = 31
            Set aPollers(y) = m_WMIWrapper.wWin32__PhysicalMemory
        Case Is = 32
            Set aPollers(y) = m_WMIWrapper.wWin32__SystemSlot
        Case Is = 33
            Set aPollers(y) = m_WMIWrapper.wWin32__Pagefile
        Case Is = 34
            Set aPollers(y) = m_WMIWrapper.wWin32__LogicalDisk
        Case Is = 35
            Set aPollers(y) = m_WMIWrapper.wWin32__QuickFixEngineering
        Case Is = 36
            Set aPollers(y) = m_WMIWrapper.wWin32__Share
        Case Is = 37
            Set aPollers(y) = m_WMIWrapper.wWin32__StartupCommand
   
    DoEvents
    End Select
End Sub

Private Sub ViewDatabase_Click()
frmDB.Show vbModal, Me  ' True  ', Me
End Sub

Private Sub PollDB()
Dim x As Integer
Dim mData As String
Dim mID As String
frmImprtDB.mRS.MoveFirst
Do Until frmImprtDB.mRS.EOF
        mData = frmImprtDB.cmbData
        mID = frmImprtDB.cmbPrimKey
        m_WMIWrapper.ComputerName = frmImprtDB.mRS(mData)
        m_WMIWrapper.ComputerID = frmImprtDB.mRS(mID)
        txtComputer.Text = m_WMIWrapper.ComputerName
      
      
    If txtComputer.Text = "." Then
        txtComputer.Text = mGlobalAPI.GetLocalHostName()
    ElseIf mGlobalAPI.Valid_IP(txtComputer.Text) Then
        txtComputer.Text = mGlobalAPI.GetHostNameFromIP(txtComputer.Text)
    End If
    
    txtComputer.Text = UCase(txtComputer.Text)
    m_WMIWrapper.ComputerName = txtComputer.Text
      
    StartPolling
    
    frmImprtDB.mRS.MoveNext
    
Loop


End Sub

Private Sub StartPolling()
Dim i As Integer
Dim y As Integer
For i = 0 To frmSelect.lstData.ListCount - 1
If frmSelect.lstData.Selected(i) Then
    For y = LBound(aNames) To UBound(aNames)
    frmSelect.lstData.ListIndex = i
    If aNames(y) = frmSelect.lstData Then
        aPollers(y).fetch
        AutosizeColumns Me.ListView1
        DoEvents
    End If
    Next
End If
DoEvents

Next
End Sub
