VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LibWIN32WMI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'***************************************************************************************
' Created By: Gary Noble - gwnoble@msn.com - 13-05-2003
' Purpose: Interface Wrapper For WMI Calls
'          Only Reference This Class In Your Project To
'          Gain Access To All The WMI Calls
'
'***************************************************************************************
' The Class Should Be Used With WithEvents To Retrieve Call Data
'***************************************************************************************
'
' Example:
'   Top Level: Private WithEvents mWMIWrapper As LibWIN32WMI
'   FormLoad Proc: Set mWMIWrapper = New LibWIN32WMI
'   FormQueryUnload Proc: If Not mWMIWrapper Is Nothing Then Set mWMIWrapper = Nothing
'
'   All Returned Data Will Be Passed Through The ItemDetail Event
'
'***************************************************************************************
' Enjoy! If You like This Code And Find It Usefull Don't Forget To Vote.
' This Code Is Also Freeware - Use It At Your Own Risk!
' If You Do Use It A Mention Somewhere Would Be Nice.
'
'***************************************************************************************
' Also If You Need Any Thing Added Just Email Me And I'll See What I Can Do!
'***************************************************************************************

Option Explicit

' Events
Public Event ItemDetail(ID As String, ComputerID As String, Computer As String, SubSystem As String, Header As String, Detail As String, PercentComplete As Long)

' WMI Classes
Private mvar_Win32__w1394Controller As win32_1394Controller
Private mvar_Win32__BaseBoard As Win32_BaseBoard
Private mvar_Win32__Account As Win32_Account
Private mvar_Win32__SoftwareFeature As Win32_SoftwareFeature
Private mvar_Win32__ApplicationService As Win32_ApplicationService
Private mvar_Win32__BaseService As Win32_BaseService
Private mvar_Win32__Battery As Win32_Battery
Private mvar_Win32__Binary As Win32_Binary
Private mvar_Win32__BindImageAction As Win32_BindImageAction
Private mvar_Win32__Bios As Win32_Bios
Private mvar_Win32__BootConfig As Win32_BootConfig
Private mvar_Win32__Bus As Win32_Bus
Private mvar_Win32__ComputerSystem As Win32_ComputerSystem
Private mvar_Win32__ComputerSystemProduct As Win32_ComputerSystemProduct
Private mvar_Win32__DiskDrive As Win32_DiskDrive
Private mvar_Win32__DiskPartition As Win32_DiskPartition
Private mvar_Win32__Network As Win32_Network
Private mvar_Win32__Printer As Win32_Printer
Private mvar_Win32__EventLog As Win32_Event_Log
Private mvar_Win32__OperatingSystem As Win32_OperatingSystem
Private mvar_Win32__Process As Win32_Process
Private mvar_Win32__Processor As Win32_Processor
Private mvar_Win32__PhysicalMemory As Win32_PhysicalMemory
Private mvar_Win32__SystemSlot As Win32_SystemSlot

Private mvar_Win32__Pagefile As Win32_Pagefile
Private mvar_Win32__LogicalDisk As Win32_LogicalDisk
Private mvar_Win32__QuickFixEngineering As Win32_QuickFixEngineering
Private mvar_Win32__Share As Win32_Share
Private mvar_Win32__StartupCommand As Win32_StartupCommand
' Used For user to Gather All CallTypes
Private CallTypes As Collection

' Progress
Private mvar_ProgMax As Long
Private mvar_CurrentRec As Long
Private mvarProgressBar As Object

' Computer Name
Private mvar_ComputerName As String
Private mvar_ComputerID As String
Private mdef_ComputerName As String

'User Credentials
Private mvar_UserName As String
Private mvar_UserPassword As String

Public Property Set ProgressBar(ByVal vData As Object)
    Set mvarProgressBar = vData
End Property


Public Property Get ProgressBar() As Object
    Set ProgressBar = mvarProgressBar
End Property

Public Property Let ComputerName(ByVal vData As String)
    
    If Len(Trim(vData)) > 0 Then
        mvar_ComputerName = vData
    Else
        mvar_ComputerName = mdef_ComputerName
    End If
    
End Property

Public Property Get ComputerName() As String
    ComputerName = mvar_ComputerName
End Property

Public Property Let UserName(ByVal vData As String)
    
    If Len(Trim(vData)) > 0 Then
        mvar_UserName = vData
    Else
        mvar_UserName = Empty
    End If
    
End Property

Public Property Get UserName() As String
    UserName = mvar_UserName
End Property

Public Property Let UserPassword(ByVal vData As String)
    
    If Len(Trim(vData)) > 0 Then
        mvar_UserPassword = vData
    Else
        mvar_UserPassword = Empty
    End If
    
End Property

Public Property Get UserPassword() As String
    UserPassword = mvar_UserPassword
End Property


Public Property Set wWin32__Account(ByVal vData As Win32_Account)
    Set mvar_Win32__Account = vData
End Property

Public Property Get ComputerID() As String
    ComputerID = mvar_ComputerID
End Property

Public Property Let ComputerID(ByVal vData As String)
    
    If Len(Trim(vData)) > 0 Then
        mvar_ComputerID = vData
    Else
        mvar_ComputerID = Empty
    End If
    
End Property

Public Property Get wWin32__Account() As Win32_Account
  If (mvar_Win32__Account Is Nothing) Then Set mvar_Win32__Account = New Win32_Account
  Set mvar_Win32__Account.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__Account = mvar_Win32__Account
End Property

Public Property Set wWin32__1394Controller(ByVal vData As win32_1394Controller)
    Set mvar_Win32__w1394Controller = vData
End Property

Public Property Get wWin32__1394Controller() As win32_1394Controller
  If (mvar_Win32__w1394Controller Is Nothing) Then Set mvar_Win32__w1394Controller = New win32_1394Controller
  Set mvar_Win32__w1394Controller.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__1394Controller = mvar_Win32__w1394Controller
End Property

Public Property Set wWin32__BaseBoard(ByVal vData As Win32_BaseBoard)
    Set mvar_Win32__BaseBoard = vData
End Property

Public Property Get wWin32__BaseBoard() As Win32_BaseBoard
  If (mvar_Win32__BaseBoard Is Nothing) Then Set mvar_Win32__BaseBoard = New Win32_BaseBoard
  Set mvar_Win32__BaseBoard.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__BaseBoard = mvar_Win32__BaseBoard
End Property

Public Property Set wWin32__SoftwareFeature(ByVal vData As Win32_SoftwareFeature)
    Set mvar_Win32__SoftwareFeature = vData
End Property

Public Property Get wWin32__SoftwareFeature() As Win32_SoftwareFeature
  If (mvar_Win32__SoftwareFeature Is Nothing) Then Set mvar_Win32__SoftwareFeature = New Win32_SoftwareFeature
  Set mvar_Win32__SoftwareFeature.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__SoftwareFeature = mvar_Win32__SoftwareFeature
  
End Property

Public Property Set wWin32__ApplicationService(ByVal vData As Win32_ApplicationService)
    Set mvar_Win32__ApplicationService = vData
End Property

Public Property Get wWin32__ApplicationService() As Win32_ApplicationService
  If (mvar_Win32__ApplicationService Is Nothing) Then Set mvar_Win32__ApplicationService = New Win32_ApplicationService
  Set mvar_Win32__ApplicationService.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__ApplicationService = mvar_Win32__ApplicationService
End Property

Public Property Set wWin32__BaseService(ByVal vData As Win32_BaseService)
    Set mvar_Win32__BaseService = vData
End Property

Public Property Get wWin32__BaseService() As Win32_BaseService
  If (mvar_Win32__BaseService Is Nothing) Then Set mvar_Win32__BaseService = New Win32_BaseService
  Set mvar_Win32__BaseService.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__BaseService = mvar_Win32__BaseService
End Property


Public Property Set wWin32__Battery(ByVal vData As Win32_Battery)
    Set mvar_Win32__Battery = vData
End Property

Public Property Get wWin32__Battery() As Win32_Battery
  If (mvar_Win32__Battery Is Nothing) Then Set mvar_Win32__Battery = New Win32_Battery
  Set mvar_Win32__Battery.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__Battery = mvar_Win32__Battery
End Property

Public Property Set wWin32__Binary(ByVal vData As Win32_Binary)
    Set mvar_Win32__Binary = vData
End Property

Public Property Get wWin32__Binary() As Win32_Binary
  If (mvar_Win32__Binary Is Nothing) Then Set mvar_Win32__Binary = New Win32_Binary
  Set mvar_Win32__Binary.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__Binary = mvar_Win32__Binary
End Property


Public Property Set wWin32__BindImageAction(ByVal vData As Win32_BindImageAction)
    Set mvar_Win32__BindImageAction = vData
End Property

Public Property Get wWin32__BindImageAction() As Win32_BindImageAction
  If (mvar_Win32__BindImageAction Is Nothing) Then Set mvar_Win32__BindImageAction = New Win32_BindImageAction
  Set mvar_Win32__BindImageAction.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__BindImageAction = mvar_Win32__BindImageAction
End Property



Public Property Set wWin32__Bios(ByVal vData As Win32_Bios)
    Set mvar_Win32__Bios = vData
End Property

Public Property Get wWin32__Bios() As Win32_Bios
  If (mvar_Win32__Bios Is Nothing) Then Set mvar_Win32__Bios = New Win32_Bios
  Set mvar_Win32__Bios.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__Bios = mvar_Win32__Bios
End Property


Public Property Set wWin32__BootConfig(ByVal vData As Win32_BootConfig)
    Set mvar_Win32__BootConfig = vData
End Property

Public Property Get wWin32__BootConfig() As Win32_BootConfig
  If (mvar_Win32__BootConfig Is Nothing) Then Set mvar_Win32__BootConfig = New Win32_BootConfig
  Set mvar_Win32__BootConfig.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__BootConfig = mvar_Win32__BootConfig
End Property


Public Property Set wWin32__Bus(ByVal vData As Win32_Bus)
    Set mvar_Win32__Bus = vData
End Property

Public Property Get wWin32__Bus() As Win32_Bus
  If (mvar_Win32__Bus Is Nothing) Then Set mvar_Win32__Bus = New Win32_Bus
  Set mvar_Win32__Bus.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__Bus = mvar_Win32__Bus
End Property


Public Property Set wWin32__ComputerSystem(ByVal vData As Win32_ComputerSystem)
    Set mvar_Win32__ComputerSystem = vData
End Property

Public Property Get wWin32__ComputerSystem() As Win32_ComputerSystem
  If (mvar_Win32__ComputerSystem Is Nothing) Then Set mvar_Win32__ComputerSystem = New Win32_ComputerSystem
  Set mvar_Win32__ComputerSystem.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__ComputerSystem = mvar_Win32__ComputerSystem
End Property

Public Property Set wWin32__ComputerSystemProduct(ByVal vData As Win32_ComputerSystemProduct)
    Set mvar_Win32__ComputerSystemProduct = vData
End Property

Public Property Get wWin32__ComputerSystemProduct() As Win32_ComputerSystemProduct
  If (mvar_Win32__ComputerSystemProduct Is Nothing) Then Set mvar_Win32__ComputerSystemProduct = New Win32_ComputerSystemProduct
  Set mvar_Win32__ComputerSystemProduct.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__ComputerSystemProduct = mvar_Win32__ComputerSystemProduct
End Property

Public Property Set wWin32__DiskDrive(ByVal vData As Win32_DiskDrive)
    Set mvar_Win32__DiskDrive = vData
End Property

Public Property Get wWin32__DiskDrive() As Win32_DiskDrive
  If (mvar_Win32__DiskDrive Is Nothing) Then Set mvar_Win32__DiskDrive = New Win32_DiskDrive
  Set mvar_Win32__DiskDrive.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__DiskDrive = mvar_Win32__DiskDrive
End Property


Public Property Set wWin32__DiskPartition(ByVal vData As Win32_DiskPartition)
    Set mvar_Win32__DiskPartition = vData
End Property

Public Property Get wWin32__DiskPartition() As Win32_DiskPartition
  If (mvar_Win32__DiskPartition Is Nothing) Then Set mvar_Win32__DiskPartition = New Win32_DiskPartition
  Set mvar_Win32__DiskPartition.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__DiskPartition = mvar_Win32__DiskPartition
End Property


Public Property Set wWin32__Network(ByVal vData As Win32_Network)
    Set mvar_Win32__Network = vData
End Property

Public Property Get wWin32__Network() As Win32_Network
  If (mvar_Win32__Network Is Nothing) Then Set mvar_Win32__Network = New Win32_Network
  Set mvar_Win32__Network.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__Network = mvar_Win32__Network
End Property


Public Property Set wWin32__Printer(ByVal vData As Win32_Printer)
    Set mvar_Win32__Printer = vData
End Property

Public Property Get wWin32__Printer() As Win32_Printer
  If (mvar_Win32__Printer Is Nothing) Then Set mvar_Win32__Printer = New Win32_Printer
  Set mvar_Win32__Printer.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__Printer = mvar_Win32__Printer
End Property

Public Property Set wWin32__Eventlog(ByVal vData As Win32_Event_Log)
    Set mvar_Win32__EventLog = vData
End Property

Public Property Get wWin32__Eventlog() As Win32_Event_Log
  If (mvar_Win32__EventLog Is Nothing) Then Set mvar_Win32__EventLog = New Win32_Event_Log
  Set mvar_Win32__EventLog.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__Eventlog = mvar_Win32__EventLog
End Property

Public Property Set wWin32__OperatingSystem(ByVal vData As Win32_OperatingSystem)
    Set mvar_Win32__OperatingSystem = vData
End Property

Public Property Get wWin32__OperatingSystem() As Win32_OperatingSystem
  If (mvar_Win32__OperatingSystem Is Nothing) Then Set mvar_Win32__OperatingSystem = New Win32_OperatingSystem
  Set mvar_Win32__OperatingSystem.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__OperatingSystem = mvar_Win32__OperatingSystem
End Property

Public Property Set wWin32__Process(ByVal vData As Win32_Process)
    Set mvar_Win32__Process = vData
End Property

Public Property Get wWin32__Process() As Win32_Process
  If (mvar_Win32__Process Is Nothing) Then Set mvar_Win32__Process = New Win32_Process
  Set mvar_Win32__Process.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__Process = mvar_Win32__Process
End Property

Public Property Set wWin32__Processor(ByVal vData As Win32_Processor)
    Set mvar_Win32__Processor = vData
End Property

Public Property Get wWin32__Processor() As Win32_Processor
  If (mvar_Win32__Processor Is Nothing) Then Set mvar_Win32__Processor = New Win32_Processor
  Set mvar_Win32__Processor.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__Processor = mvar_Win32__Processor
End Property

Public Property Set wWin32__PhysicalMemory(ByVal vData As Win32_PhysicalMemory)
    Set mvar_Win32__PhysicalMemory = vData
End Property

Public Property Get wWin32__PhysicalMemory() As Win32_PhysicalMemory
  If (mvar_Win32__PhysicalMemory Is Nothing) Then Set mvar_Win32__PhysicalMemory = New Win32_PhysicalMemory
  Set mvar_Win32__PhysicalMemory.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__PhysicalMemory = mvar_Win32__PhysicalMemory
End Property


Public Property Set wWin32__SystemSlot(ByVal vData As Win32_SystemSlot)
    Set mvar_Win32__SystemSlot = vData
End Property

Public Property Get wWin32__SystemSlot() As Win32_SystemSlot
  If (mvar_Win32__SystemSlot Is Nothing) Then Set mvar_Win32__SystemSlot = New Win32_SystemSlot
  Set mvar_Win32__SystemSlot.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__SystemSlot = mvar_Win32__SystemSlot
End Property


Public Property Set wWin32__Pagefile(ByVal vData As Win32_Pagefile)
    Set mvar_Win32__Pagefile = vData
End Property

Public Property Get wWin32__Pagefile() As Win32_Pagefile
  If (mvar_Win32__Pagefile Is Nothing) Then Set mvar_Win32__Pagefile = New Win32_Pagefile
  Set mvar_Win32__Pagefile.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__Pagefile = mvar_Win32__Pagefile
End Property

Public Property Set wWin32__LogicalDisk(ByVal vData As Win32_LogicalDisk)
    Set mvar_Win32__LogicalDisk = vData
End Property

Public Property Get wWin32__LogicalDisk() As Win32_LogicalDisk
  If (mvar_Win32__LogicalDisk Is Nothing) Then Set mvar_Win32__LogicalDisk = New Win32_LogicalDisk
  Set mvar_Win32__LogicalDisk.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__LogicalDisk = mvar_Win32__LogicalDisk
End Property

Public Property Set wWin32__QuickFixEngineering(ByVal vData As Win32_QuickFixEngineering)
    Set mvar_Win32__QuickFixEngineering = vData
End Property

Public Property Get wWin32__QuickFixEngineering() As Win32_QuickFixEngineering
  If (mvar_Win32__QuickFixEngineering Is Nothing) Then Set mvar_Win32__QuickFixEngineering = New Win32_QuickFixEngineering
  Set mvar_Win32__QuickFixEngineering.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__QuickFixEngineering = mvar_Win32__QuickFixEngineering
End Property

Public Property Set wWin32__Share(ByVal vData As Win32_Share)
    Set mvar_Win32__Share = vData
End Property

Public Property Get wWin32__Share() As Win32_Share
  If (mvar_Win32__Share Is Nothing) Then Set mvar_Win32__Share = New Win32_Share
  Set mvar_Win32__Share.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__Share = mvar_Win32__Share
End Property

Public Property Set wWin32__StartupCommand(ByVal vData As Win32_StartupCommand)
    Set mvar_Win32__StartupCommand = vData
End Property

Public Property Get wWin32__StartupCommand() As Win32_StartupCommand
  If (mvar_Win32__StartupCommand Is Nothing) Then Set mvar_Win32__StartupCommand = New Win32_StartupCommand
  Set mvar_Win32__StartupCommand.System = Me                  ' /* Always maintain a reference to the correct parent. */
  Set wWin32__StartupCommand = mvar_Win32__StartupCommand
End Property

Private Sub Class_Initialize()
    
    Set CallTypes = New Collection
    
    ' Set The Default Computer Name
    ' A '.' Represents This Computer
    mdef_ComputerName = "."
    Me.ComputerName = mdef_ComputerName
    
End Sub

Private Sub Class_Terminate()
  
  Set CallTypes = Nothing
  If Not (mvar_Win32__BaseBoard Is Nothing) Then Set mvar_Win32__BaseBoard = Nothing
  If Not (mvar_Win32__w1394Controller Is Nothing) Then Set mvar_Win32__w1394Controller = Nothing
  If Not (mvar_Win32__Account Is Nothing) Then Set mvar_Win32__Account = Nothing
  If Not (mvar_Win32__SoftwareFeature Is Nothing) Then Set mvar_Win32__SoftwareFeature = Nothing
  If Not (mvar_Win32__ApplicationService Is Nothing) Then Set mvar_Win32__ApplicationService = Nothing
  If Not (mvar_Win32__BaseService Is Nothing) Then Set mvar_Win32__BaseService = Nothing
  If Not (mvar_Win32__Battery Is Nothing) Then Set mvar_Win32__Battery = Nothing
  If Not (mvar_Win32__Binary Is Nothing) Then Set mvar_Win32__Binary = Nothing
  If Not (mvar_Win32__BindImageAction Is Nothing) Then Set mvar_Win32__BindImageAction = Nothing
  If Not (mvar_Win32__Bios Is Nothing) Then Set mvar_Win32__Bios = Nothing
  If Not (mvar_Win32__BootConfig Is Nothing) Then Set mvar_Win32__BootConfig = Nothing
  If Not (mvar_Win32__Bus Is Nothing) Then Set mvar_Win32__Bus = Nothing
  If Not (mvar_Win32__ComputerSystem Is Nothing) Then Set mvar_Win32__ComputerSystem = Nothing
  If Not (mvar_Win32__ComputerSystemProduct Is Nothing) Then Set mvar_Win32__ComputerSystemProduct = Nothing
  If Not (mvar_Win32__DiskDrive Is Nothing) Then Set mvar_Win32__DiskDrive = Nothing
  If Not (mvar_Win32__DiskPartition Is Nothing) Then Set mvar_Win32__DiskPartition = Nothing
  If Not (mvar_Win32__Network Is Nothing) Then Set mvar_Win32__Network = Nothing
  If Not (mvar_Win32__Printer Is Nothing) Then Set mvar_Win32__Printer = Nothing
  If Not (mvar_Win32__EventLog Is Nothing) Then Set mvar_Win32__EventLog = Nothing
  If Not (mvar_Win32__OperatingSystem Is Nothing) Then Set mvar_Win32__OperatingSystem = Nothing
  If Not (mvar_Win32__Process Is Nothing) Then Set mvar_Win32__Process = Nothing
  If Not (mvar_Win32__Processor Is Nothing) Then Set mvar_Win32__Processor = Nothing
  If Not (mvar_Win32__PhysicalMemory Is Nothing) Then Set mvar_Win32__PhysicalMemory = Nothing
  If Not (mvar_Win32__SystemSlot Is Nothing) Then Set mvar_Win32__SystemSlot = Nothing
  If Not (mvar_Win32__Pagefile Is Nothing) Then Set mvar_Win32__Pagefile = Nothing
  If Not (mvar_Win32__LogicalDisk Is Nothing) Then Set mvar_Win32__LogicalDisk = Nothing
  If Not (mvar_Win32__QuickFixEngineering Is Nothing) Then Set mvar_Win32__QuickFixEngineering = Nothing
  If Not (mvar_Win32__Share Is Nothing) Then Set mvar_Win32__Share = Nothing
  If Not (mvar_Win32__StartupCommand Is Nothing) Then Set mvar_Win32__StartupCommand = Nothing
 
  End Sub

'***************************************************************************************
' Return the Available Calls
'***************************************************************************************
Public Property Get WMICalls() As Collection

  Set CallTypes = Nothing
  Set CallTypes = New Collection
  
        CallTypes.Add "1394Controller"
        CallTypes.Add "BaseBoard"
        CallTypes.Add "Account"
        CallTypes.Add "Account SID"
        CallTypes.Add "Installed Software"
        CallTypes.Add "Application Service"
        CallTypes.Add "Base Service"
        CallTypes.Add "Battery"
        CallTypes.Add "Binary"
        CallTypes.Add "Image Action"
        CallTypes.Add "Bios"
        CallTypes.Add "Boot Config"
        CallTypes.Add "Bus"
        CallTypes.Add "Computer System"
        CallTypes.Add "Computer System Product"
        CallTypes.Add "Disk Drives"
        CallTypes.Add "Disk Partition"
        CallTypes.Add "Network Adapter"
        CallTypes.Add "Network Adapter Config"
        CallTypes.Add "Printers"
        CallTypes.Add "Printer Config"
        CallTypes.Add "Printer Jobs"
        CallTypes.Add "Network Client"
        CallTypes.Add "Network Connection"
        CallTypes.Add "Network Login Profile"
        CallTypes.Add "Network Protocol"
        CallTypes.Add "Event Log File"
        CallTypes.Add "Event Log Events"
        CallTypes.Add "Operating System"
        CallTypes.Add "System Processes"
        CallTypes.Add "System Processor"
        CallTypes.Add "Physical Memory"
        CallTypes.Add "System Slots"
        CallTypes.Add "Page File"
        CallTypes.Add "Logical Disks"
        CallTypes.Add "HotFixes"
        CallTypes.Add "Shares"
        CallTypes.Add "Startup Items"
        
        
       Set WMICalls = CallTypes
        
End Property


Friend Property Let ProgMax(lMax As Long)
On Error GoTo ErrorHandler:
       mvar_CurrentRec = 0
       mvar_ProgMax = lMax
       If Not Me.ProgressBar Is Nothing Then
            Me.ProgressBar.Max = lMax
            Me.ProgressBar.value = 0
       End If
       
ErrorHandler:
End Property
Friend Function UpdateProgress()
On Error GoTo ErrorHandler:
    
    mvar_CurrentRec = mvar_CurrentRec + 1
    If Not Me.ProgressBar Is Nothing Then
        Me.ProgressBar.value = Me.ProgressBar.value + 1
    End If
    
ErrorHandler:
End Function

Friend Function AddressReturnData(strID As String, strComputerID As String, strComputer As String, strSubSys As String, strHeading As String, strData As String)
        RaiseEvent ItemDetail(strID, strComputerID, strComputer, strSubSys, strHeading, strData, mvar_CurrentRec)
End Function

