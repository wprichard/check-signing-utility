VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Weitz Export"
   ClientHeight    =   5160
   ClientLeft      =   1650
   ClientTop       =   2370
   ClientWidth     =   8235
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Billing"
            Object.ToolTipText     =   "Backup Billing"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Check"
            Object.ToolTipText     =   "Sign Checks"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Form"
            Object.ToolTipText     =   "Open Workflow Form"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4890
      Visible         =   0   'False
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6324
            Text            =   "Status"
            TextSave        =   "Status"
            Key             =   "Status"
            Object.ToolTipText     =   "Program Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Key             =   "SID"
            Object.ToolTipText     =   "Specialist ID"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Key             =   "User"
            Object.ToolTipText     =   "User"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "11:43 PM"
            Key             =   "Time"
            Object.ToolTipText     =   "Time"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1080
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2824
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C76
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuViewSettings 
      Caption         =   "&Settings"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuSettingsLog 
         Caption         =   "Use Debug &Log"
         Checked         =   -1  'True
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsBilling 
         Caption         =   "&Backup Billing"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuToolsChecks 
         Caption         =   "&Check Signing"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"390DBAFD03A1"
'*****************************************************************************************
' Class:        frmMain
' Author:       Wes Prichard, Optika
' Date:         November 1999
' Description:  MDI form that implements top-level business logic.
' Edit History:
' 11/27/2002 - Modified by W. Prichard, Optika
'   Disabled log file so multiple instances can run on the same server.
'   Also disabled and hid Settings menu.
' 04/30/2003 - Modified by W. Prichard, Optika
'   Activated Backup Billing toolbar button

'Public Interface:
'Properties:
'none

'Methods:
'Display - makes form visible

'Events:
'Initialized

'Dependencies:
'TODO

'ActiveX controls and references used by this form:
'Toolbar control
'CommonDialog control
'Imagelist control
'StatusBar control
'*****************************************************************************************

Option Explicit

Const CLASSNAME = "WeitzExport" 'used for app-specific registry values

'TODO Capture form close events and destroy

''Win API function declarations
'Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

'Event declarations
Public Event Initialized()
'Indicates that initialization is complete.
'Public Event Trace(Message As String)
''Contains a string containing trace information that can be logged.

'Object references
'TODO add custom unload events so object refs can be destroyed
Private WithEvents mfrmCheckSign As frmCheckSign
Attribute mfrmCheckSign.VB_VarHelpID = -1
Private WithEvents mfrmBillingBackup As frmBillingBackup
Attribute mfrmBillingBackup.VB_VarHelpID = -1
Private mobjLog As cLog            'Logging object

'Private class data
Dim mstrLogPath As String            'path to log files
Dim mblnTraceLogEnabled As Boolean  'Logging enable
Dim mlngLogMin As Long              'Log carry-over record count
Dim mlngLogMax As Long              'Log reduction threshold
Dim mstrLogFileName As String       'Log file name (no path because app.path is used)
Dim mudtLogType As enuLogType       'Log type
Private m_blnLoggedIn As Boolean    'indicates if logged in to mainframe
Private m_blnNoNextOnUpdate  As Boolean 'indicates that no Next to be performed in Update

Public Sub Display()     'UserID As String, Password As String)
'Public entry point for the form.

'Indicate app is busy
    Me.MousePointer = vbHourglass

'Show the form to the user so that the child forms will size properly
    Me.Show


'Indicate not busy
    Me.MousePointer = vbNormal
    
    mobjLog.WriteLine "Main form display complete"

End Sub

Private Sub MDIForm_Load()
'Perform any form-specific initialization.
'(Application-specific initialization is done in Display method.)
    
'Set window state
    Me.WindowState = vbMaximized
    
'Get all stored program settings
    Call Initialize

'Update status bar
    sbStatusBar.Panels("Status").Text = ""
    
'Indicate initialized
    mobjLog.WriteLine "Initialization complete"
    RaiseEvent Initialized

End Sub

Private Sub Initialize()
'Perform custom initialization for this application.
      
'Dim Result As Login     'eMedia login result
Dim i As Integer        'loop counter

'Enable error handler
    On Error GoTo errHandler
    
'Read the registry settings
'    Set m_frmSettings = New frmSettings
'    Load m_frmSettings  'load it to get settings but don't display it
    Call GetRegSettings
    
'Start Trace Log
    Set mobjLog = New cLog    'explicitly instantiate object
    'Configure log object according to registry settings
    mobjLog.LogType = mudtLogType
    mobjLog.MaxLines = 10000
    mobjLog.MinLines = 8000
    mobjLog.LoggingEnabled = mblnTraceLogEnabled
'    mobjLog.LoggingEnabled = False 'True (true for testing/debugging)
    mobjLog.OpenLogFile mstrLogPath & "\" & mstrLogFileName
'    mobjLog.OpenLogFile App.Path & "\Log-" & App.Title & ".txt"
    mobjLog.WriteLine ""
    mobjLog.WriteLine "Log file opened for new instance of " & App.Title & " - version " & _
        App.Major & "." & App.Minor & "." & App.Revision
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            mobjLog.WriteLine "Raising Error in Initialize: " & Err.Number & "-" & Err.Description
            Err.Raise Err.Number, Err.Source, Err.Description
    End Select
'Resume '@@@

End Sub

Private Sub GetRegSettings()
'Get values from the registry that are needed to execute.
'Note - do not write to log from this procedure (since log object could not be initialized)

'The following requires a project reference to Optika Registry Tool (optikareg.dll)
Dim cReg As cRegistry
Dim strVal As String    'a temporary holder for a registry value
Dim intCount As Integer 'array index
Dim i As Integer        'loop counter
Dim strHoldEvents  As String    'temp var
Dim intComma As Integer         'character position in string

Const APPKEYOPTIKA = "Software\Optika\Common"
Const APPKEYCLASS = "Software\OptikaCustom\" & CLASSNAME
    
'Enable error trap
    On Error GoTo errHandler

'Instantiate registry object
    Set cReg = New cRegistry

'Trace log values
    'Put log in same location as other Acorde service log files
    mstrLogPath = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYOPTIKA, "EventLogPath", App.Path)
    'Get log values from the other key
    mblnTraceLogEnabled = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "LogEnabled", True)
    mlngLogMin = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "LogMin", 1000)
    mlngLogMax = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "LogMax", 5000)
    mstrLogFileName = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "LogFileName", CLASSNAME & "Log.txt")
    strVal = cReg.QueryValue(regHKEY_LOCAL_MACHINE, APPKEYCLASS, "LogType", "Daily")
    'If value is daily then...
    If LCase$(strVal) = "daily" Then
        mudtLogType = logDaily
    Else
        mudtLogType = logCircular
    End If
    
'Insert new properties here

'Destroy the reg object reference
    Set cReg = Nothing

Exit Sub

errHandler:

    Select Case Err.Number
        Case Else
            Call RaiseError(Err.Number, CLASSNAME & ".GetRegSettings", _
                Err.Number & "-" & Err.Description)
        End Select
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'This event fires when the MDI form is being closed and occurs before the Unload
'or Terminate events.

Dim i As Integer

'Close all forms
    For i = Forms.Count - 1 To 0 Step -1
        Unload Forms(i)
    Next i

End Sub

Private Sub MDIForm_Terminate()
'Clean up memory and terminate the application.  This event will not occur if any
'form are open in the application.

'Logout
    mobjLog.WriteLine "Disconnecting from Workflow"
    mobjLog.WriteLine "Disconnecting from the database"

'Destroy global object references
    Set mfrmCheckSign = Nothing
    mobjLog.WriteLine "Ending the application"
    Set mobjLog = Nothing
    
'End the application
    End

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
'If not minimized then...
    If Me.WindowState <> vbMinimized Then
        'Save window size and position
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
'End the application
    End
    
End Sub

Private Sub mfrmBillingBackup_LogMsg(Message As String)
    mobjLog.WriteLine Message
End Sub

Private Sub mfrmBillingBackup_MousePtr(PointerStyle As MousePointerConstants)
    Me.MousePointer = PointerStyle
End Sub

Private Sub mfrmBillingBackup_Status(Message As String)
   sbStatusBar.Panels("Status").Text = Message
End Sub

Private Sub mfrmCheckSign_LogMsg(Message As String)
    mobjLog.WriteLine Message
End Sub

Private Sub mfrmCheckSign_Status(Message As String)
   sbStatusBar.Panels("Status").Text = Message
End Sub

Private Sub mnuSettingsLog_Click()
    mobjLog.WriteLine "menu Settings Log"
    mnuSettingsLog.Checked = Not mnuSettingsLog.Checked
    mobjLog.LoggingEnabled = Not mobjLog.LoggingEnabled

End Sub

Private Sub mnuToolsBilling_Click()

    mobjLog.WriteLine "menu Tools Billing"
    Set mfrmBillingBackup = New frmBillingBackup
    mfrmBillingBackup.Display

End Sub

Private Sub mnuToolsChecks_Click()

    mobjLog.WriteLine "menu Tools Checks"
    Set mfrmCheckSign = New frmCheckSign
    mfrmCheckSign.Display

End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.key
        Case "Billing"
            Call mnuToolsBilling_Click
        
        Case "Check"
            Call mnuToolsChecks_Click
        
        Case Else
            MsgBox "Button key not found", , "tbToolBar_ButtonClick"
    End Select

End Sub

Private Sub mnuHelpAbout_Click()
    mobjLog.WriteLine "menu Help About"
    frmAbout.Show vbModal, Me
End Sub

'Private Sub mnuHelpContents_Click()
''VB App Wizard generated
'
'    Dim nRet As Integer
'
'    'if there is no helpfile for this project display a message to the user
'    'you can set the HelpFile for your application in the
'    'Project Properties dialog
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
'    Else
'        On Error Resume Next
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
'        If Err Then
'            MsgBox Err.Description
'        End If
'    End If
'
'End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    mobjLog.WriteLine "menu File Exit"
    Unload Me

End Sub

Private Sub RaiseError(ErrorNumber As Long, Source As String, Description As String)
'Log and raise the error.

'Log the error
    mobjLog.WriteLine "(frmMain.RaiseError) Error in " & Source & ": " & ErrorNumber & " - " & Description
    
'Raise an error up to the client
    Err.Raise ErrorNumber, Source, Description
    
End Sub



