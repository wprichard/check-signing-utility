VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D718B306-B8FB-11D1-B34F-00C04FD0D58E}#1.0#0"; "OptPage.ocx"
Begin VB.Form frmCheckSign 
   Caption         =   "Check Signing"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12360
   Icon            =   "frmCheckSign.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   12360
   Begin OPTPAGELib.OptPage OptPage1 
      Height          =   8055
      Left            =   5520
      TabIndex        =   12
      Top             =   1080
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
      _ExtentY        =   14208
      _StockProps     =   98
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      ToolTipText     =   "Retrieves the selected invoice nto the viewer"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtCount 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Page 0 of 0"
      ToolTipText     =   "Page Number and Page Count"
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   7920
      TabIndex        =   8
      ToolTipText     =   "This is the check report being processed"
      Top             =   0
      Width           =   7215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Open a Check Report..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      ToolTipText     =   "Opens a file dialog to select a check report"
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "&Last"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      ToolTipText     =   "Selects the last invoice in the list"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      ToolTipText     =   "Selects the next invoice in the list"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "&Previous"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      ToolTipText     =   "Selects the previous invoice in the list"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "&First"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Select the first invoice in the list"
      Top             =   480
      Width           =   855
   End
   Begin MSComctlLib.TreeView trvChecks 
      Height          =   9135
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Treeview containing the list of checks, the invoices being paid, and the invoice coding"
      Top             =   960
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   16113
      _Version        =   393217
      Style           =   6
      Appearance      =   1
   End
   Begin VB.ComboBox cboApps 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "This is the Invoice application that will be searched to find the document images"
      Top             =   0
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   7920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   14640
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":23D2
            Key             =   "Larger"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":2824
            Key             =   "Smaller"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":2C76
            Key             =   "Fit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":30C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":351A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":396C
            Key             =   "Previous"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":3DBE
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":4210
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":4322
            Key             =   "FitSides"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":463C
            Key             =   "FitWindow"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":4956
            Key             =   "NextPage"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":4C70
            Key             =   "PrevPage"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":4F8A
            Key             =   "ZoomIn"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":5B5C
            Key             =   "ZoomOut"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Height          =   420
      Left            =   5520
      TabIndex        =   11
      Top             =   600
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   3000
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Zoomin"
            Object.ToolTipText     =   "Zoom In"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Zoomout"
            Object.ToolTipText     =   "Zoom Out"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Fit"
            Object.ToolTipText     =   "Fit To Window"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Previous"
            Object.ToolTipText     =   "Previous Page"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            Object.ToolTipText     =   "Next Page"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print Document"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList imgTree 
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":5E76
            Key             =   "Invoice"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckSign.frx":62C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblApp 
      Caption         =   "Select Search:"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmCheckSign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
' Class:        frmCheckSign
' Author:       Wes Prichard, Optika
' Date:         August 2002
' Description:  Provides the interface to support check signing.
' Edit History:
' 11/18/2002 - Wes Prichard, Optika
'   Modifed ParseCheckReport to parse a different report format. Previously, the
'   report did not include any commas after the last field of the row. The new
'   report format always includes all commas. The number of leading consecutive
'   commas is now used to identify the information on the row, rather than the
'   number of fields in the row.

' 11/27/02 - Wes Prichard, Optika
'   Modified LoadInvoice to unload the previous image. Removed the same code from
'   FindAndDisplayInvoice so that the previous image would always be cleared when
'   an attempt is made to load a new image.

' 04/30/03 - Wes Prichard, Optika
'   ParseCheckReport - commented out message box in error handler that warns of a
'   duplicate invoice in the input file, per Weitz request.

' 06/15/2004 - Wes Prichard, Optika
'   Modified ParseTXTCheckReport to handle multiple commas inside quotes. v2.1.2

' 01/04/2005 - Wes Prichard, Optika, v2.2.1
'   Modified procedures that set file paths from registry settings to trap errors
'   resulting from invalid drives:
'   cmdSelect_Click

'06/05/2006 - Rafael Geraldino, Stellent V2.3.0
'   Compiled for IBPM 7.6

'Public Interface:
'Properties:
'None

'Methods:
'Display - makes form visible

'Events:
'Status - reports a status for the status bar
'LogMsg - Provides a log entry

'Dependencies:
'Microsoft Scripting Runtime (scrrun.dll)
'Microsoft ActiveX Data Objects 2.6 Library (msado15.dll)
'IBPM SDK User Security 1.0 Type Library (OTUsrSec.dll)
'IBPM SDK Schema 1.0 Type Library (OTSchema.dll)
'IBPM SDK ObjectID 1.0 Type Library (OTObjID.dll)
'IBPM SDK Query 1.0 Type Library (OTQuery.dll)
'*****************************************************************************************

Option Explicit

'Viewer flag bit definition
'Value  Notes
'1      Reserved
'2      Reserved
'4      Reserved
'8      Reserved

'16     Establish a user as annotation administrator for a particular application.  User must have Modify rights to be an annotation administrator.
'32     Reserved
'64     Allows the user to add or modify annotations.
'128    Reserved

'256    If set to true, the currently viewed file will be deleted when the page is closed.
'512    Disables the ability to print
'1024   Disables the ability to Fax
'2048   Disables export capability

'4096   Reserved
'8192   Disables the overlay
'16384  Reserved
'32768  Reserved

'65536  Enables stamp administration
'131072 Enables black redaction
'262144 Enables white redaction
'524288 Enables Span Documents

'1048576    Disables annotation capability
'2097152    Show hits on page
'4194304    Reserved
'8388608    Disable launch capability
Private Const VIEWERFLAGS = 393494  '393494 = 060116 Hex = 0000 0110 0000 0001 0001 0110 Bin

Private Const ERRORBASE = ErrorBase7    'see modErrorHandling - used to avoid
                                        '   overlapping error numbers
Private Const MODULE = "frmCheckSign"     'used in reporting errors


'Class Error enumeration
Public Enum errCDatabase
    errcmdNavigate = ERRORBASE + 0
    errParseCheckReport = ERRORBASE + 1
    errLoad = ERRORBASE + 2
    errGetSavedSearches = ERRORBASE + 3
    errFindAndDisplayInvoice = ERRORBASE + 4
    errToolBarClick = ERRORBASE + 5
    errExecuteNamedQuery = ERRORBASE + 6
    errLoadInvoice = ERRORBASE + 7
    errGeneric = ERRORBASE + 8
'    errDBRollbackTrans = ERRORBASE + 9
'    errDBRollbackTrans = ERRORBASE + 10
End Enum

Private Enum ParsingState
    NewCheck
    NewInvoice
    NewGL
End Enum

'Public event declarations
Public Event Status(Message As String)  'status to the user
Public Event LogMsg(Message As String)  'message to the log

'Structure for carrying invoice query parameters
Private Type udtInvoice
    InvoiceNum As String
    SupplierNum As String
End Type

'Module-scope variables
Dim mstrReportFile As String    'the path and name of the check report file
Dim mstrChkInfo() As String      'dynamic array for the check report content
Dim mobjUserToken As OTACORDELib.UserToken     'Acorde login user token
Dim blnGoodLogin As Boolean     'indicates if login was successful
Dim mobjNameParser As New OTCONTEXTLib.NameParser    'TODO make local
Dim mintPageNumber As Integer   'viewer page number
Dim mintPageCount As Integer    'viewer page count
'Dim mstrSchemaPrefix As String  'prefix to fully qualify a schema table

'Public Sub Display(Top As Single, Left As Single, Height As Single, Width As Single, UserID As String, Password As String)
Public Sub Display()
'Custom entry point for loading and displaying form so that modality and other
'properties can be controlled.
 
'Make me visible
    RaiseEvent LogMsg("(frmCheckSign.Display) Entering procedure")
    Me.Show vbModeless  'this triggers Form_Load procedure
    Me.WindowState = vbMaximized
    Me.Refresh
    
End Sub

Private Sub Form_Load()
'Initialize the form

Dim strUser As String
Dim strPassword As String
Dim ofrmLogin As frmLogin
Dim objUser As OTACORDELib.User

'Enable error trap
    On Error GoTo errHandler

'Setup treeview
'    Debug.Print trvChecks.Indentation
'    Debug.Print trvChecks.PathSeparator
    trvChecks.SingleSel = True      'expand node when selected
    trvChecks.Style = tvwPictureText
    trvChecks.ImageList = imgTree
    trvChecks.LineStyle = tvwTreeLines

'Disable navigation buttons until the treeview is laoded
    cmdFirst.Enabled = False
    cmdLast.Enabled = False
    cmdNext.Enabled = False
    cmdPrev.Enabled = False

'Disable view buttons until the treeview is loaded
    cmdView.Enabled = False

'Login to IBPM
    Do While mobjUserToken Is Nothing
        'Get the last user name
        strUser = GetSetting(App.Title, "Settings", "Acorde User")
        'Display the login form
        Set ofrmLogin = New frmLogin
        RaiseEvent LogMsg("(frmCheckSign.Form_Load) Opening Login dialog...")
        ofrmLogin.Display strUser
        strUser = ofrmLogin.UserName
        strPassword = ofrmLogin.Password
        
        'If user did not cancel login then...
        If ofrmLogin.OK = True Then
            Set ofrmLogin = Nothing
            Me.Refresh
            
            'Save the name of the successful login
            SaveSetting App.Title, "Settings", "Acorde User", strUser
            
            'Call the IBPM login
            Screen.MousePointer = vbHourglass
            Set objUser = New OTACORDELib.User
            RaiseEvent LogMsg("(frmCheckSign.Form_Load) User completed Login dialog - logging into IBPM as user: " & strUser)
            Set mobjUserToken = objUser.Login(strUser, strPassword, False) 'hide dialog
            Screen.MousePointer = vbNormal
            
            'Notify user if login unsuccessful
            If mobjUserToken Is Nothing Then
                MsgBox "Login was unsuccessful. Check user name and remember " & _
                    "that password are case-sensitive.", vbExclamation + vbOKOnly, _
                    "IBPM Login Failed"
            End If
        Else
            RaiseEvent LogMsg("(frmCheckSign.Form_Load) User cancelled Login dialog...")
            Exit Do
    '        Set ofrmLogin = Nothing
    '        'user cancelled the login dialog
    '        MsgBox "A successful Acorde login is required to run the check signing tool.", _
    '            vbInformation + vbOKOnly, "Acorde Login Cancelled"
    '        Exit Sub
        End If
    
    Loop

'If login successful then...
    If Not (mobjUserToken Is Nothing) Then
        RaiseEvent LogMsg("(frmCheckSign.Form_Load) IBPM login successful")
        blnGoodLogin = True
    
        'Initialize the application combo
        Call GetInvoiceApps(cboApps)
        Me.Refresh
        'cboApps.SetFocus   'can't do this until form is visible
    
        'cboApps.Text = GetSetting(App.Title, "Settings", "Application")
    Else    'login cancelled
        RaiseEvent LogMsg("(frmCheckSign.Form_Load) IBPM login failed")
        blnGoodLogin = False
        MsgBox "You will not be able to view invoice images because the " & _
            "login to IBPM was not successful.", vbExclamation + vbOKOnly, _
            "IBPM Login Result"
        cmdView.Enabled = False     'disable the view button since it won't work
        cmdSelect.Enabled = True    'allow user to view a report file
    End If
    
    RaiseEvent LogMsg("(frmCheckSign.Form_Load) Form Load complete")
Exit Sub

errHandler:
    Select Case Err.Number
        Case -2147220640
            MsgBox Err.Description, vbInformation + vbOKOnly, "IBPM Login Error"
            Resume Next
        Case -2147023570    'Failed Login method.
            MsgBox Err.Description, vbInformation + vbOKOnly, "IBPM Login Error"
            Resume Next
        Case -2147467259    'The Domain property is not set or is invalid.
            MsgBox Err.Description, vbInformation + vbOKOnly, "IBPM Login Error"
            Resume Next
        Case Else
            Call RaiseError(errLoad, MODULE & ".Form_Load", Err.Number & "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub Form_Resize()
'Scale the controls when the form is resized.

'Enable error trap
    On Error GoTo errHandler
        
'Keep the bottom of treeview at the bottom of the form
    trvChecks.Height = Me.Height - trvChecks.Top - 400
    
'Keep the bottom of the viewer at the bottom of the form
    OptPage1.Height = Me.Height - OptPage1.Top - 400
    
'Keep the right side of the viewer on the right edge of the form
    OptPage1.Width = Me.Width - OptPage1.Left - 150
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".Form_Resize", Err.Number & "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub cboApps_Click()
'Enable the report selection button now that the user has select an application.

'Enable error trap
    On Error GoTo errHandler
        
'If an application was selected then...
    If cboApps.Text <> vbNullString Then
        cmdSelect.Enabled = True
    End If
    
    SaveSetting App.Title, "Settings", "Saved Search", cboApps.Text
    RaiseEvent LogMsg("(frmCheckSign.cboApps_Click) Saved Search '" & cboApps.Text & "' selected")
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".cboApps_Click", Err.Number & "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub cmdView_Click()
'View the selected invoice in the tree

'Enable error trap
    On Error GoTo errHandler
        
'Make sure an invoice is selected in the tree view
    If trvChecks.SelectedItem.SelectedImage = 2 Then
        RaiseEvent LogMsg("(frmCheckSign.cmdView_Click) Locating and viewing invoice: " & trvChecks.SelectedItem.Text)
        Call LoadInvoice
    Else
        MsgBox "The selected item in the tree view is not an invoice and does not have an associated image.", vbInformation + vbOKOnly, "View Invoice"
    End If
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".cmdView_Click", Err.Number & "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub cmdFirst_Click()
'Expand the first check node in the tree view and expand all children, and select
'the first invoice and display it in the viewer.

'Enable error trap
    On Error GoTo errHandler     'Trap it if user cancels dialog

'Note:
'Nodes(1) = "Checks"
'Nodes(2) = first check
'Nodes(3) = first invoice
    

'Force scroll to top of Treeview by selecting first node
    trvChecks.Nodes(1).Selected = True
    trvChecks.SelectedItem.EnsureVisible
    
'Select the first invoice
    trvChecks.Nodes(3).Selected = True
'    trvChecks.Nodes(3).Checked = True
    trvChecks.SelectedItem.EnsureVisible
    Me.Refresh
'    trvChecks.SelectedItem.Bold = True
    
'Locate the invoice document and load it in the viewer
    RaiseEvent LogMsg("(frmCheckSign.cmdFirst_Click) Loading Invoice")
    LoadInvoice
    
Exit Sub

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errcmdNavigate, MODULE & ".cmdFirst_Click", Err.Number & "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub cmdPrev_Click()
'Load the previous invoice in the treeview.

Dim strInvoice As String
Dim strCheckInfo As String

'Enable error trap
    On Error GoTo errHandler     'Trap it if user cancels dialog

'If there is a current selection then...
    If Not (trvChecks.SelectedItem Is Nothing) Then
    
        Select Case Left(trvChecks.SelectedItem.key, 1)
            Case "c"    'top node
                'Select the first invoice node
                Call cmdFirst_Click
            
            Case "*"    'check node
                'Select the last child of the previous check if there is one
                'If there is a previous check then...
                If Not (trvChecks.SelectedItem.Previous Is Nothing) Then
                    trvChecks.SelectedItem.Previous.Child.LastSibling.Selected = True
                Else    'at beginning of list
                    MsgBox "You are at the beginning of the list", , "Previous"
                End If
            
            Case "#"    'invoice node
                'Proceed with selecting the previous invoice
                'If there is not a previous invoice for this check then...
                If trvChecks.SelectedItem.Previous Is Nothing Then
                    'If there is a previous check then...
                    If Not (trvChecks.SelectedItem.Parent.Previous Is Nothing) Then
                        'Select the last child invoice of the previous check
                        trvChecks.SelectedItem.Parent.Previous.Child.LastSibling.Selected = True
                    Else
                        MsgBox "You are at the beginning of the list", , "Previous"
                    End If
                Else    'there is a previous sibling invoice
                    trvChecks.SelectedItem.Previous.Selected = True
                End If
            
            Case ""     'coding node
                'Select the parent invoice
                trvChecks.SelectedItem.Parent.Selected = True
            
        End Select
        
    Else    'there is no node selected
        'select the first invoice node
        Call cmdFirst_Click
    End If
            
    trvChecks.SelectedItem.EnsureVisible
    Me.Refresh

'Locate the invoice document and load it in the viewer
    RaiseEvent LogMsg("(frmCheckSign.cmdPrev_Click) Loading Invoice")
    LoadInvoice
    
Exit Sub

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errcmdNavigate, MODULE & ".cmdPrev_Click", Err.Number & "-" & Err.Description)
    End Select
End Sub

Private Sub cmdNext_Click()
'Load the next invoice in the treeview.

Dim strInvoice As String
Dim strCheckInfo As String

'Enable error trap
    On Error GoTo errHandler     'Trap it if user cancels dialog

'If there is a current selection then...
    If Not (trvChecks.SelectedItem Is Nothing) Then
    
        Select Case Left(trvChecks.SelectedItem.key, 1)
            Case "c"    'top node
                'Select the first invoice node
                Call cmdFirst_Click
            
            Case "*"    'check node
                'Select the first child of the check
                trvChecks.SelectedItem.Child.Selected = True
                trvChecks.SelectedItem.Child.Bold = True
            
            Case "#"    'invoice node
                'Proceed with selecting the next invoice
                'If there is not another invoice for this check then...
                If trvChecks.SelectedItem.Next Is Nothing Then
                    'If there is another check then...
                    If Not (trvChecks.SelectedItem.Parent.Next Is Nothing) Then
                        'Select the first child invoice of the next check
                        trvChecks.SelectedItem.Parent.Next.Child.Selected = True
                        trvChecks.SelectedItem.Bold = True
                    Else
                        MsgBox "You are at the end of the list!", , "Next"
                    End If
                Else    'there is another sibling invoice
                    trvChecks.SelectedItem.Next.Selected = True
                     trvChecks.SelectedItem.Bold = True
                End If
            
            Case ""     'coding node
                'Select the parent invoice
                trvChecks.SelectedItem.Parent.Selected = True
                
                'Proceed with selecting the next invoice
                'If there is not another invoice for this check then...
                If trvChecks.SelectedItem.Next Is Nothing Then
                    'If there is another check then...
                    If Not (trvChecks.SelectedItem.Parent.Next Is Nothing) Then
                        'Select the first child invoice of the next check
                        trvChecks.SelectedItem.Parent.Next.Child.Selected = True
                         trvChecks.SelectedItem.Bold = True
                    Else
                        MsgBox "You are at the end of the list!", , "Next"
                    End If
                Else    'there is another sibling invoice
                    trvChecks.SelectedItem.Next.Selected = True
                    trvChecks.SelectedItem.Bold = True
                End If
            
        End Select
        
    Else    'there is no node selected
        'select the first invoice node
        Call cmdFirst_Click
    End If
    
    trvChecks.SelectedItem.EnsureVisible
    Me.Refresh

'Locate the invoice document and load it in the viewer
    RaiseEvent LogMsg("(frmCheckSign.cmdNext_Click) Loading Invoice")
    Call LoadInvoice
    
Exit Sub

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errcmdNavigate, MODULE & ".cmdPrev_Click", Err.Number & "-" & Err.Description)
    End Select
End Sub

Private Sub cmdLast_Click()
'Load the last invoice in the treeview.

Dim strInvoice As String
Dim strCheckInfo As String

'Enable error trap
    On Error GoTo errHandler     'Trap it if user cancels dialog

'Select the last invoice of the parent check's last sibling
    trvChecks.SelectedItem.Parent.LastSibling.Child.LastSibling.Selected = True
    trvChecks.SelectedItem.EnsureVisible
    Me.Refresh

'Locate the invoice document and load it in the viewer
    RaiseEvent LogMsg("(frmCheckSign.cmdLast_Click) Loading Invoice")
    Call LoadInvoice
    
Exit Sub

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errcmdNavigate, MODULE & ".cmdPrev_Click", Err.Number & "-" & Err.Description)
    End Select
End Sub

Private Sub cmdSelect_Click()
'Open a file dialog and get the fully qualified name of the input report
'to be processed. Then cause the report to be parsed and displayed.

Dim strPath As String

'Enable error trap
    On Error GoTo errHandler     'Trap it if user cancels dialog


'Open a file dialog to select the output file
    dlgFile.CancelError = True
    dlgFile.Filter = "All (*.*)|*.*|Text (*.txt)|*.txt|CSV (*.csv)|*.csv"
    'dlgFile.Filter = "All (*.*)|*.*|CSV (*.csv)|*.csv"
    On Error Resume Next
    dlgFile.InitDir = GetSetting(App.EXEName, "Files", "ReportPath", App.Path)
    If Err.Number <> 0 Then
        'Assume path in registry is not valid
        dlgFile.InitDir = "C:\"
    End If
    On Error GoTo errHandler
    dlgFile.ShowOpen
    
'Get file information selected by user
    strPath = Left(dlgFile.FileName, InStr(1, dlgFile.FileName, dlgFile.FileTitle) - 1)
    SaveSetting App.EXEName, "Files", "ReportPath", strPath
    mstrReportFile = dlgFile.FileName
    txtFilePath.Text = dlgFile.FileName
    
'Clear the tree view
    trvChecks.Nodes.Clear
    Me.Refresh
   
'Parse the report and load the tree view

    frmMain.sbStatusBar.Panels("Status").Text = "Loading Check File, please wait..."
    
    RaiseEvent LogMsg("(frmCheckSign.cmdSelect_Click) Parsing check report file: " & mstrReportFile)
    'ParseCSVCheckReport (mstrReportFile)
    ParseTXTCheckReport (mstrReportFile)
    
    frmMain.sbStatusBar.Panels("Status").Text = "Loading Check File, please wait...Done."
    
'Select the first check in the treeview
    Call cmdFirst_Click

Exit Sub

errHandler:
    Select Case Err.Number
        Case cdlCancel
            Exit Sub
        Case Else
            Call RaiseError(errGeneric, MODULE & ".cmdSelect_Click", Err.Number & "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub GetInvoiceApps(Combo As ComboBox)
'Load the specified combo control with Invoice applications.

'Requires a project reference to Acorde SDK Schema 1.0 Type Library (OTSchema.dll)
Dim objSchema As OTCONTEXTLib.Connections
Dim objConnection As OTCONTEXTLib.Connection
Dim intConn As Integer
Dim intTable As Integer
Dim strTableName As String
    
'Enable error trap
    On Error GoTo errHandler

'Set the user token and refresh the schema
    Set objSchema = New OTCONTEXTLib.Connections
    objSchema.UserToken = mobjUserToken
    
'Add application table names to the combo box
    'For each schema...
    For intConn = 1 To objSchema.Count
        'Get the connection
        Set objConnection = objSchema.Item(intConn)
        
        'for each table in the connection...
        For intTable = 1 To objConnection.Tables.Count
            strTableName = objConnection.Tables.Item(intTable).FullName
            'If this is an invoice application then...
            If strTableName Like "*Invoice*" Then
                'Add name to the combo box list
                cboApps.AddItem strTableName
            End If
        Next intTable
    Next intConn

    RaiseEvent LogMsg("(" & MODULE & ".GetInvoiceApps) " & cboApps.ListCount & _
        " applications added to the combo box list.")
    Set objConnection = Nothing
    Set objSchema = Nothing

Exit Sub

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errGetInvoiceApps, MODULE & ".GetInvoiceApps", Err.Number & "-" & Err.Description)
    End Select
End Sub

'
'
Private Sub GetSavedSearches(Combo As ComboBox)
'Load the specified combo control with Saved Searches.

  Dim objNamedQueries As New OTCONTEXTLib.NamedQueries
  Dim objNamedQuery As OTCONTEXTLib.NamedQuery
  Dim strSavedSetting As String
  
  objNamedQueries.UserToken = mobjUserToken
  objNamedQueries.Refresh
  
  strSavedSetting = GetSetting(App.Title, "Settings", "Saved Search")
  
  For Each objNamedQuery In objNamedQueries
    cboApps.AddItem objNamedQuery.Name
    If (strSavedSetting = objNamedQuery.Name) Then
      cboApps.Text = objNamedQuery.Name
    End If
  Next



    RaiseEvent LogMsg("(frmCheckSign.GetSavedSearches) " & cboApps.ListCount & _
        " searches added to the combo box list.")
    
    Set objNamedQueries = Nothing

Exit Sub

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errGetSavedSearches, MODULE & ".SavedSearches", Err.Number & "-" & Err.Description)
    End Select
End Sub

Private Function ParseTXTCheckReport(strFile As String) As String()
'Parse the specified report and extract the check information, invoice
'information, and coding information into an array.
'Inputs:
'strFile - the fully qualifed path and name of the report file to be processed

'Requires project reference to Microsoft Scripting Runtime (scrrun.dll)
Dim fs As New FileSystemObject
Dim TextStreamIn As TextStream
'strChkInfo(col, row)   0 based
Dim strLine As String       'line from the report file
Dim strSplit() As String    'dynamic array containing all fields in one report line
Dim strReport() As String   'dynamic array of report content without blank columns
Dim lngRow As Long          'strReport array row index
Dim intBefore As Integer    'position of prior comma in the string
Dim intAfter As Integer     'position of next comma in the string
'Dim strFieldVal As String   'a field from the
'Dim uRecState As udtRecState    'indicates state of record processing
Dim objNodes As Nodes       'treeview nodes collection
Dim objNode As Node         'treeview node object
Dim strCheckKey As String   'key of current check node in treeview
Dim strInvoiceKey As String 'key of current invoice node in treeview
Dim intFirstQuote As Integer
Dim intSecondQuote As Integer
Dim strSub As String
Dim strFixed As String
Dim strLastCheck As String
Dim strLastInvoice As String
Dim intState As ParsingState
Dim i As Integer            'loop counter

'Array index constants
'These are the indexes of the fields in the first dimension of the strReport array
Const ICHECKNUM = 0
Const IPAYEE = 1
Const ICHKAMT = 2
Const IINVOICE = 3
Const ISUPPLIER = 4
Const IACCT = 5
Const IGLAMT = 6

'Enable error trap
    On Error GoTo errHandler

'Create top node in tree
    Set objNodes = trvChecks.Nodes
    'Set objNodes = TreeView1.Nodes  '@@@
    Set objNode = objNodes.Add(, , "c", "Checks")

'Open the check report file
    Set TextStreamIn = fs.OpenTextFile(strFile, ForReading, False)
    RaiseEvent LogMsg("(frmCheckSign.ParseCheckReport) Report file '" & strFile & "' opened for reading")
    
'No header lines to skip
    
'Init the array
    ReDim strReport(6, 0)   '1 row with 8 fields
    strLastCheck = ""
    strLastInvoice = ""
    
'Populate the report array with the desired fields from the file
    'For all lines in the file...
    Do While Not TextStreamIn.AtEndOfStream
        strLine = TextStreamIn.ReadLine
        Debug.Print "Line # " & TextStreamIn.Line
        
        'If line contains quotes then replace commas inside quotes so Split function can be used
        intSecondQuote = 0
        Do Until intSecondQuote = Len(strLine)
            'Find the quote
            intFirstQuote = InStr(intSecondQuote + 1, strLine, """")
            'If quote found then...
            If intFirstQuote > 0 Then
                'Find the second quote
                intSecondQuote = InStr(intFirstQuote + 1, strLine, """")
                'if a comma between the quotes then
                strSub = Mid(strLine, intFirstQuote + 1, intSecondQuote - intFirstQuote - 1)
                If InStr(1, strSub, ",") Then
                    'Replace the comma with a semicolon
                    'strFixed = Replace(strLine, ",", ".", intFirstQuote + 1, 1)
                    'strFixed = Replace(strSub, ",", ".", 1, 1)
                    strFixed = Replace(strSub, ",", "~", 1, -1)
                    strLine = Left(strLine, intFirstQuote) & strFixed & Right(strLine, Len(strLine) - intSecondQuote + 1)
'                    Debug.Print strLine
                End If
            End If
        Loop
        
        'Remove quotes so values can be trimmed
        strLine = Replace(strLine, """", "")
        Debug.Print strLine
        
        'Split the row into an array of substrings
        strSplit = Split(strLine, ",")
        
        'Re-insert the commas that had previously been replaced
        For i = 0 To UBound(strSplit)
            strSplit(i) = Replace(strSplit(i), "~", ",", 1, -1)
        Next i
        
        'If the check number has changed then...
        If strSplit(0) <> strLastCheck Then
            intState = NewCheck
        Else    'not a new check record
            'If this a new invoice then...
            If strSplit(6) <> strLastInvoice Then
                intState = NewInvoice
            Else    'not a new invoice
                'Just add another GL amount to the existing invoice node
                intState = NewGL
            End If
        End If
            
        Select Case intState
            Case NewCheck
                Debug.Print "New Check"
                'Add a new check node
                'Add a new row for the check record
                lngRow = UBound(strReport, 2) + 1 'get the index of the next new row
                ReDim Preserve strReport(6, lngRow)
                
                'Extract Payee name to the existing check record
                strReport(IPAYEE, lngRow) = Trim(strSplit(1))
                'Extract Check Number and Check Amount
                strReport(ICHECKNUM, lngRow) = RDepad(LDepad(strSplit(0)))
                strLastCheck = strSplit(0) 'save the check number
                strCheckKey = "*" & strReport(ICHECKNUM, lngRow)
                strReport(ICHKAMT, lngRow) = LDepad(strSplit(2))
                
                'Create a new check node in tree view
                Set objNode = objNodes.Add("c", tvwChild, _
                    strCheckKey, _
                    strReport(ICHECKNUM, lngRow) & ", " & strReport(IPAYEE, lngRow) & ", $" & strReport(ICHKAMT, lngRow))
                objNode.EnsureVisible
                
                'Add a new Invoice node
                'Add a new row for the invoice record
                'lngRow = UBound(strReport, 2) + 1 'get the index of the next new row
                'ReDim Preserve strReport(6, lngRow)
                'Extract Invoice Number
                'Jobnum 8
                'Vendor Name 12
                strReport(IINVOICE, lngRow) = Trim(strSplit(6))
                strLastInvoice = strSplit(6)    'save the invoice number
                'Extract Supplier Number
                strReport(ISUPPLIER, lngRow) = LDepad(Left(strSplit(4), 21))
                strInvoiceKey = "#" & strReport(ISUPPLIER, lngRow) & "-" & strReport(IINVOICE, lngRow)
                
                'Create a new invoice node in the tree view
                Set objNode = objNodes.Add(strCheckKey, tvwChild, _
                    strInvoiceKey, _
                    strReport(IINVOICE, lngRow), 1)
                objNode.SelectedImage = 2
                
                'Add the GL node
                'Extract GL Acct Number and GL Amount
                strReport(IACCT, lngRow) = Trim(strSplit(7))
                strReport(IGLAMT, lngRow) = LDepad(strSplit(8))
                
                'Create a new coding node in the tree view
                Set objNode = objNodes.Add(strInvoiceKey, tvwChild, _
                    , _
                    strReport(IACCT, lngRow) & ", $" & strReport(IGLAMT, lngRow))
            
            
            Case NewInvoice
                Debug.Print "New Invoice"
                'Add a new Invoice node
                'Add a new row for the invoice record
                lngRow = UBound(strReport, 2) + 1 'get the index of the next new row
                ReDim Preserve strReport(6, lngRow)
                
                'Extract Payee name to the existing check record
                strReport(IPAYEE, lngRow) = Trim(strSplit(1))
                'Extract Check Number and Check Amount
                strReport(ICHECKNUM, lngRow) = RDepad(LDepad(strSplit(0)))
                strReport(ICHKAMT, lngRow) = LDepad(strSplit(2))
                
                'Extract Invoice Number
                strReport(IINVOICE, lngRow) = Trim(strSplit(6))
                strLastInvoice = strSplit(6)    'save the invoice number
                
                'Extract Supplier Number
                strReport(ISUPPLIER, lngRow) = LDepad(Left(strSplit(4), 21))
                strInvoiceKey = "#" & strReport(ISUPPLIER, lngRow) & "-" & strReport(IINVOICE, lngRow)
                
                'Create a new invoice node in the tree view
                Set objNode = objNodes.Add(strCheckKey, tvwChild, _
                    strInvoiceKey, _
                    strReport(IINVOICE, lngRow), 1)
                objNode.SelectedImage = 2
                
                'Add the GL node
                'Extract GL Acct Number and GL Amount
                strReport(IACCT, lngRow) = Trim(strSplit(7))
                strReport(IGLAMT, lngRow) = LDepad(strSplit(8))
                
                'Create a new coding node in the tree view
                Set objNode = objNodes.Add(strInvoiceKey, tvwChild, _
                    , _
                    strReport(IACCT, lngRow) & ", $" & strReport(IGLAMT, lngRow))
            
            Case NewGL
                Debug.Print "New GL"
                'Add the GL node
                'Add a new row for the coding record
                lngRow = UBound(strReport, 2) + 1 'get the index of the next new row
                ReDim Preserve strReport(6, lngRow)
                
                'Extract Payee name to the existing check record
                strReport(IPAYEE, lngRow) = Trim(strSplit(1))
                'Extract Check Number and Check Amount
                strReport(ICHECKNUM, lngRow) = RDepad(LDepad(strSplit(0)))
                strReport(ICHKAMT, lngRow) = LDepad(strSplit(2))
                
                'Extract Invoice Number
                strReport(IINVOICE, lngRow) = Trim(strSplit(6))
                
                'Extract Supplier Number
                strReport(ISUPPLIER, lngRow) = LDepad(Left(strSplit(4), 21))
                
                'Extract GL Acct Number and GL Amount
                strReport(IACCT, lngRow) = Trim(strSplit(7))
                strReport(IGLAMT, lngRow) = LDepad(strSplit(8))
                
                'Create a new coding node in the tree view
                Set objNode = objNodes.Add(strInvoiceKey, tvwChild, _
                    , _
                    strReport(IACCT, lngRow) & ", $" & strReport(IGLAMT, lngRow))
            
        End Select
            
'            Case 12 'invoice row 1 (for a credit with no invoice number)
'                'Add a new row for the invoice record
'                lngRow = UBound(strReport, 2) + 1 'get the index of the next new row
'                ReDim Preserve strReport(6, lngRow)
'                'Extract Invoice Number
'                'Jobnum 8
'                'Vendor Name 12
'                strReport(IINVOICE, lngRow) = "(none)"
            
        
    Loop
    'objNode.EnsureVisible
    
'Enable navigation buttons now that the treeview is loaded
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    cmdPrev.Enabled = True

'Enable view buttons now that the treeview is loaded
    If blnGoodLogin = True Then cmdView.Enabled = True

'Set focus on the next button
    cmdNext.SetFocus
    
    RaiseEvent LogMsg("(frmCheckSign.ParseCheckReport) Parsing of check report file '" _
        & strFile & "' complete. " & UBound(strReport, 2) & " records extracted.")
    

Exit Function

errHandler:
    Select Case Err.Number
        
        Case 35602  'Key is not unique in collection
            'MsgBox "A duplicate invoice record '" & strInvoiceKey & "' was found in the report for check: " & strCheckKey
            strInvoiceKey = strInvoiceKey & objNodes.Count
            Resume
'        Case 35601  'element not found
'            MsgBox "The report file contains a record that is out of sequence." & vbCrLf & "The record cannot be displayed."
'            Resume Next
'        Case 9  'subscript out of range
'            Resume Next
        Case Else
            Call RaiseError(errParseCheckReport, MODULE & ".ParseCheckReport", Err.Number & "-" & Err.Description)
    End Select
'Resume  '@@@

End Function

Private Function ParseCSVCheckReport(strFile As String) As String()
'Parse the specified report and extract the check information, invoice
'information, and coding information into an array.
'Inputs:
'strFile - the fully qualifed path and name of the report file to be processed

'Requires project reference to Microsoft Scripting Runtime (scrrun.dll)
Dim fs As New FileSystemObject
Dim TextStreamIn As TextStream
'strChkInfo(col, row)   0 based
Dim strLine As String       'line from the report file
Dim strSplit() As String    'dynamic array containing all fields in one report line
Dim strReport() As String   'dynamic array of report content without blank columns
Dim lngRow As Long          'strReport array row index
Dim intBefore As Integer    'position of prior comma in the string
Dim intAfter As Integer     'position of next comma in the string
'Dim strFieldVal As String   'a field from the
'Dim uRecState As udtRecState    'indicates state of record processing
Dim objNodes As Nodes       'treeview nodes collection
Dim objNode As Node         'treeview node object
Dim strCheckKey As String   'key of current check node in treeview
Dim strInvoiceKey As String 'key of current invoice node in treeview
Dim intFirstQuote As Integer
Dim intSecondQuote As Integer
Dim strSub As String
Dim strFixed As String
    
'Array index constants
'These are the indexes of the fields in the first dimension of the strReport array
Const ICHECKNUM = 0
Const IPAYEE = 1
Const ICHKAMT = 2
Const IINVOICE = 3
Const ISUPPLIER = 4
Const IACCT = 5
Const IGLAMT = 6

'Enable error trap
    On Error GoTo errHandler

    Set objNodes = trvChecks.Nodes
    'Set objNodes = TreeView1.Nodes  '@@@
    Set objNode = objNodes.Add(, , "c", "Checks")

'Open the check report file
    Set TextStreamIn = fs.OpenTextFile(strFile, ForReading, False)
    RaiseEvent LogMsg("(frmCheckSign.ParseCheckReport) Report file '" & strFile & "' opened for reading")
    
'Skip first two lines containing headers
    TextStreamIn.SkipLine
    TextStreamIn.SkipLine
    
'Init the array
    ReDim strReport(7, 0)   '1 row with 7 fields
    
'Populate the report array with the desired fields from the file
    'For all lines in the file...
    Do While Not TextStreamIn.AtEndOfStream
        strLine = TextStreamIn.ReadLine
        RaiseEvent LogMsg("(frmCheckSign.ParseCheckReport) " & strLine)
        'If line contains quotes then remove commas inside quotes
        'Find the quote
        intFirstQuote = InStr(1, strLine, """")
        'If quote found then...
        If intFirstQuote > 0 Then
            'Find the second quote
            intSecondQuote = InStr(intFirstQuote + 1, strLine, """")
            'if a comma between the quotes then
            strSub = Mid(strLine, intFirstQuote + 1, intSecondQuote - intFirstQuote - 1)
            If InStr(1, strSub, ",") Then
                'Replace the comma with a semicolon
                strFixed = Replace(strLine, ",", ".", intFirstQuote + 1, 1)
                strLine = Left(strLine, intFirstQuote) & strFixed '& Right(strLine, Len(strLine) - intSecondQuote + 1)
            End If
        End If
        
        strSplit = Split(strLine, ",")
        'Select Case UBound(strSplit)
        Select Case LeadingCommas(strLine)
            Case 2  'check row 1
                'Add a new row for the check record
                lngRow = UBound(strReport, 2) + 1 'get the index of the next new row
                ReDim Preserve strReport(6, lngRow)
                
                'Extract Payee name to the existing check record
                strReport(IPAYEE, lngRow) = strSplit(2)
                
            'Case 6  'check row 2
            Case 0  'check row 2
                'Extract Check Number and Check Amount
                strReport(ICHECKNUM, lngRow) = strSplit(0)
                strCheckKey = "*" & strSplit(0)
                strReport(ICHKAMT, lngRow) = strSplit(6)
               
                'Create a new check node in tree view
                Set objNode = objNodes.Add("c", tvwChild, _
                    strCheckKey, _
                    strReport(ICHECKNUM, lngRow) & ", " & strReport(IPAYEE, lngRow) & ", $" & strReport(ICHKAMT, lngRow))
                objNode.EnsureVisible
            
            'Case 16 'invoice row 1
            Case 8 'invoice row 1
                'Add a new row for the invoice record
                lngRow = UBound(strReport, 2) + 1 'get the index of the next new row
                ReDim Preserve strReport(6, lngRow)
                'Extract Invoice Number
                'Jobnum 8
                'Vendor Name 12
                strReport(IINVOICE, lngRow) = strSplit(16)
            
'            Case 12 'invoice row 1 (for a credit with no invoice number)
'                'Add a new row for the invoice record
'                lngRow = UBound(strReport, 2) + 1 'get the index of the next new row
'                ReDim Preserve strReport(6, lngRow)
'                'Extract Invoice Number
'                'Jobnum 8
'                'Vendor Name 12
'                strReport(IINVOICE, lngRow) = "(none)"
            
            Case 10 'invoice row 2
                'Extract Supplier Number
                strReport(ISUPPLIER, lngRow) = strSplit(10)
                strInvoiceKey = "#" & strReport(ISUPPLIER, lngRow) & "-" & strReport(IINVOICE, lngRow)
                
                'Create a new invoice node in the tree view
                Set objNode = objNodes.Add(strCheckKey, tvwChild, _
                    strInvoiceKey, _
                    strReport(IINVOICE, lngRow), 1)
                objNode.SelectedImage = 2
            
            'Case 22 'coding row
            Case 19 'coding row
                'Add a new row for the coding record
                lngRow = UBound(strReport, 2) + 1 'get the index of the next new row
                ReDim Preserve strReport(6, lngRow)
                'Extract GL Acct Number and GL Amount
                strReport(IACCT, lngRow) = Trim(strSplit(19))
                strReport(IGLAMT, lngRow) = strSplit(22)
                
                'Create a new coding node in the tree view
                Set objNode = objNodes.Add(strInvoiceKey, tvwChild, _
                    , _
                    strReport(IACCT, lngRow) & ", $" & strReport(IGLAMT, lngRow))
                
            Case Else
                MsgBox "Invalid input line detected:" & vbCrLf & strLine & vbCrLf & _
                    "This line will not be displayed in the check list."
        End Select
        
    Loop
    'objNode.EnsureVisible
    
'Enable navigation buttons now that the treeview is loaded
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    cmdPrev.Enabled = True

'Enable view buttons now that the treeview is loaded
    If blnGoodLogin = True Then cmdView.Enabled = True

'Set focus on the next button
    cmdNext.SetFocus
    
    RaiseEvent LogMsg("(frmCheckSign.ParseCheckReport) Parsing of check report file '" _
        & strFile & "' complete. " & UBound(strReport, 2) & " records extracted.")
    

Exit Function

errHandler:
    Select Case Err.Number
        
        Case 35602  'Key is not unique in collection
            'MsgBox "A duplicate invoice record '" & strInvoiceKey & "' was found in the report for check: " & strCheckKey
            strInvoiceKey = strInvoiceKey & objNodes.Count
            Resume
        Case 35601  'element not found
            MsgBox "The report file contains a record that is out of sequence." & vbCrLf & "The record cannot be displayed."
            Resume Next
'        Case 9  'subscript out of range
'            Resume Next
        Case Else
            Call RaiseError(errParseCheckReport, MODULE & ".ParseCheckReport", Err.Number & "-" & Err.Description)
    End Select
'Resume  '@@@

End Function

Private Function ExecuteAdHocQuery(strInvoice As String, strSupplierNumber As String) _
    As ADODB.Recordset
'Run an Acorde ad hoc query to locate the specified invoice document.
'Return a recordset containing the matching document record(s).

'This requires a project reference to OTQuery.dll
Dim objFromTable As OTCONTEXTLib.FromItemTable
Dim objWhereClause As OTCONTEXTLib.WhereClause
Dim objSearchCondition As OTCONTEXTLib.SearchCondition
Dim objPredicate As OTCONTEXTLib.BinaryCondition
Dim intResultSize As Integer
Dim objQManager As OTCONTEXTLib.QueryConnectionManager
Dim objQConnection As OTCONTEXTLib.QueryConnection
Dim objSQLSelect As OTCONTEXTLib.SQLSelect
Dim objField As OTCONTEXTLib.SelectItemExpression
Dim objQStat As OTCONTEXTLib.QueryStatus
Dim datStart As Date

'Enable error trap
    On Error GoTo errHandler
        
'Select RECID FROM <checkapptable> WHERE InvoiceNum = <val> AND SupplierNum = <val>
'   AND DocType = 'Invoice'

'Build the SELECT clause
    RaiseEvent LogMsg("Building SELECT")
    Set objSQLSelect = New OTCONTEXTLib.SQLSelect
    Set objField = New OTCONTEXTLib.SelectItemExpression
    objField.Expression = cboApps.Text & ".PAGECOUNT"
    objSQLSelect.SelectClause.SelectItems.Add objField

'Build the FROM clause
'Get the table name
    RaiseEvent LogMsg("Building FROM")
    Set objFromTable = New OTCONTEXTLib.FromItemTable
    objFromTable.TableName = cboApps.Text
        
'Populate SQL Object with the table (schema) to be searched
    RaiseEvent LogMsg("Populate SQL object with table to be searched")
    objSQLSelect.FromClause.FromItems.Add objFromTable
        
'Build the WHERE clause
'Instantiate and init the search conditions object
    RaiseEvent LogMsg("Building WHERE")
    Set objSearchCondition = New OTCONTEXTLib.SearchCondition
    
'Add the InvoiceNum field to the search conditions
    RaiseEvent LogMsg("Add Invoice number")
    objSearchCondition.Conjunction = otSearchConditionConjunction_NOP
    Set objPredicate = New OTCONTEXTLib.BinaryCondition
    objPredicate.Field.Name = cboApps.Text & ".InvoiceNum"
    objPredicate.Operator = otComparisonOperator_Equal
    objPredicate.Field.Value = strInvoice   'from tree view
'    objPredicate.Operator = otComparisonOperator_Like
'    objPredicate.Field.Value = "*" & strInvoice & "*"   'from tree view
    'add predicate to the search condition collection
    objSearchCondition.Conditions.Add objPredicate
    
'Add search conditions to SQL Object
    RaiseEvent LogMsg("Add Search conditions")
    objSQLSelect.WhereClause.SearchCondition.Conditions.Add objSearchCondition
    Set objSearchCondition = Nothing
    Set objPredicate = Nothing
    
'Add the SupplierNum field to the search conditions
    RaiseEvent LogMsg("Add Supplier Number to the search")
    Set objSearchCondition = New OTCONTEXTLib.SearchCondition
    objSearchCondition.Conjunction = otSearchConditionConjunction_AND
    Set objPredicate = New OTCONTEXTLib.BinaryCondition
    objPredicate.Field.Name = cboApps.Text & ".SupplierNum"
    objPredicate.Operator = otComparisonOperator_Equal
    objPredicate.Field.Value = strSupplierNumber 'from tree view
    'add predicate to the search condition collection
    objSearchCondition.Conditions.Add objPredicate
    
'Add search conditions to SQL Object
    RaiseEvent LogMsg("Add search conditions (again)")
    objSQLSelect.WhereClause.SearchCondition.Conditions.Add objSearchCondition
    Set objSearchCondition = Nothing
    Set objPredicate = Nothing
    
    
'Add the DocType = 'Invoice' to the search conditions
    RaiseEvent LogMsg("Adding DocType")
    Set objSearchCondition = New OTCONTEXTLib.SearchCondition
    objSearchCondition.Conjunction = otSearchConditionConjunction_AND
    Set objPredicate = New OTCONTEXTLib.BinaryCondition
    objPredicate.Field.Name = cboApps.Text & ".DocType"
    objPredicate.Operator = otComparisonOperator_Equal
    objPredicate.Field.Value = "Invoice"
    'add predicate to the search condition collection
    objSearchCondition.Conditions.Add objPredicate
    
'Add search conditions to SQL Object
    RaiseEvent LogMsg("Add Search conditions (again)")
    objSQLSelect.WhereClause.SearchCondition.Conditions.Add objSearchCondition
    '// RaiseEvent LogMsg("(" & MODULE & ".ExecuteAdHocQuery) SQL query statement = " & objSQLSelect.SQLString)
    '// RaiseEvent LogMsg("(" & MODULE & ".ExecuteAdHocQuery) SQL Select Valid = " & objSQLSelect.Validate)
    
'Instantiate and init query connection
    RaiseEvent LogMsg("Instantiate and init query")
    Set objQManager = New OTCONTEXTLib.QueryConnectionManager
    objQManager.UserToken = mobjUserToken
    Set objQConnection = objQManager.CreateConnection
    
'Set a default result set size
    intResultSize = 5

'Execute search
    RaiseEvent LogMsg("(" & MODULE & ".ExecuteAdHocQuery) Executing query...")
    datStart = Now
    objQConnection.ExecuteQuery objSQLSelect, otQueryStyleUnmanagedRecordset, _
        intResultSize
    RaiseEvent LogMsg("(" & MODULE & ".ExecuteAdHocQuery) Query results returned in " & DateDiff("s", datStart, Now) & " secs")
    
'Return the query results
    Set ExecuteAdHocQuery = objQConnection.GetResults(otReturnAll, objQStat)
    
'Clean up
    Set objSQLSelect = Nothing
    Set objField = Nothing
    Set objFromTable = Nothing
    Set objSearchCondition = Nothing
    Set objPredicate = Nothing
    Set objQManager = Nothing
    Set objQConnection = Nothing
 
Exit Function

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errExecuteAdHocQuery, MODULE & ".ExecuteAdHocQuery", _
                Err.Number & "-" & Err.Description)
    End Select
'Resume  '@@@
End Function

'
Private Sub FindAndDisplayInvoice(udtInvoiceDat As udtInvoice)
'Query for the specified invoice in the application designated
'by the user and load it into the viewer.

Dim obj As Object
Dim key As Variant
Dim SC As String
Dim i As Integer
Dim strInvoice As String
Dim strSupplierNum As String
'Requires a project reference to OTObjID.dll
Dim objObjID As OTOBJIDLib.ObjectID
'Requires a project reference to ADO
Dim objQResults As ADODB.Recordset
Dim datStart As Date

'Enable error trap
    On Error GoTo errHandler

'Query for the invoice document
    RaiseEvent LogMsg("(frmCheckSign.FindAndDisplayInvoice) Querying for invoice #: " & udtInvoiceDat.InvoiceNum & " with supplier #:" & udtInvoiceDat.SupplierNum)
    datStart = Now
    Set objQResults = ExecuteAdHocQuery(udtInvoiceDat.InvoiceNum, udtInvoiceDat.SupplierNum)
    RaiseEvent LogMsg("(frmCheckSign.FindAndDisplayInvoice) Query processing completed in " & DateDiff("s", datStart, Now) & " secs")

'Instantiate and init and ObjectId object
    Set objObjID = New OTOBJIDLib.ObjectID
    objObjID.UserToken = mobjUserToken
    objObjID.AutoResolve = False

'Get the object ID (LUCID) for the resulting document(s)
    Do While Not objQResults.EOF
        i = i + 1
        objObjID.Value = objQResults.Fields(0)
        objQResults.MoveNext
    Loop
    'If there were no records then
    If i = 0 Then
        MsgBox "The invoice's image was not found. " & vbCrLf & _
            "The image could be stored in an application that is not being search or " & _
            "it could mean the invoice has not been imaged.", vbInformation + vbOKOnly, _
            "View Image"
    Else
        If i > 1 Then
            MsgBox i & " Invoices were found in IBPM with this invoice " & _
                "number and supplier number. The last one will be displayed. " & _
                vbCrLf & _
                "You can query for this invoice in IBPM to see all the " & _
                "invoice documents.", vbInformation + vbOKOnly, "Invoice Query Results"
        End If
        Debug.Print "Resolving document " & i & " from the query results"
        RaiseEvent LogMsg("(frmCheckSign.FindAndDisplayInvoice) " & i & _
            " documents returned by the query; the last one is being resolved...")
        datStart = Now
        objObjID.Resolve
        RaiseEvent LogMsg("(frmCheckSign.FindAndDisplayInvoice) Resolution completed in " & DateDiff("s", datStart, Now) & " secs")
        
        'Load the image into the viewer
'        Call UnloadImage    'unload any page in the viewer in case the new image
                                'is not loaded
        Me.Refresh
        RaiseEvent LogMsg("(frmCheckSign.FindAndDisplayInvoice) Preparing to load the image into the viewer...")
        Call ShowImage(objObjID, udtInvoiceDat.InvoiceNum, True)
    End If

Exit Sub

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errFindAndDisplayInvoice, MODULE & ".FindAndDisplayInvoice", Err.Number & "-" & Err.Description)
    End Select
End Sub

Private Sub trvChecks_GotFocus()
'Enable the Previous and Next buttons

'Enable error trap
    On Error GoTo errHandler
        
    cmdNext.Enabled = True
    cmdPrev.Enabled = True

Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".trvChecks_GotFocus", Err.Number & "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
'Handle button clicks on the viewer's toolbar.

Const STEP = 10     'zoom increment
    
'Enable error trap
    On Error GoTo errHandler

    Select Case Button.key
        Case "Zoomin"
            'Zoom in one step
            RaiseEvent LogMsg("(frmCheckSign.tbToolbar_ButtonClick) Zoom In")
            OptPage1.Zoom = OptPage1.Zoom + STEP
            'optPage1.VertScrollOffset = 50
            'optPage1.HorzScrollOffset = 50
            
        Case "Zoomout"
            'Zoom out one step
            RaiseEvent LogMsg("(frmCheckSign.tbToolbar_ButtonClick) Zoom Out")
            OptPage1.Zoom = OptPage1.Zoom - STEP
'            optPage1.VertScrollBarDisabled = False
'            optPage1.HorzScrollBarDisabled = False
            'optPage1.VertScrollOffset = 50
            'optPage1.HorzScrollOffset = 50
            
        Case "Fit"
            'Fit to window
            RaiseEvent LogMsg("(frmCheckSign.tbToolbar_ButtonClick) Fit To Window")
            OptPage1.FitWindow = True
        
        Case "Previous"
'            If optPage1.PageNumber > 1 Then _
'                optPage1.PageNumber = optPage1.PageNumber - 1
            'Navigate to previous page
            RaiseEvent LogMsg("(frmCheckSign.tbToolbar_ButtonClick) Previous Page")
            OptPage1.PrevPage
            'Update buttons
            If OptPage1.IsPrevPage = True Then
                tbToolBar.Buttons("Previous").Enabled = True
            Else
                tbToolBar.Buttons("Previous").Enabled = False
            End If
            If OptPage1.IsNextPage = True Then
                tbToolBar.Buttons("Next").Enabled = True
            Else
                tbToolBar.Buttons("Next").Enabled = False
            End If
            mintPageNumber = mintPageNumber - 1
            txtCount.Text = "Page " & mintPageNumber & " of " & mintPageCount
        
        Case "Next"
'            If optPage1.PageNumber < optPage1.PageCount Then _
'                optPage1.PageNumber = optPage1.PageNumber + 1
            'Navigate to next page
            RaiseEvent LogMsg("(frmCheckSign.tbToolbar_ButtonClick) Next Page")
            OptPage1.NextPage
            'Update buttons
            If OptPage1.IsPrevPage = True Then
                tbToolBar.Buttons("Previous").Enabled = True
            Else
                tbToolBar.Buttons("Previous").Enabled = False
            End If
            If OptPage1.IsNextPage = True Then
                tbToolBar.Buttons("Next").Enabled = True
            Else
                tbToolBar.Buttons("Next").Enabled = False
            End If
            mintPageNumber = mintPageNumber + 1
            txtCount.Text = "Page " & mintPageNumber & " of " & mintPageCount
        
        Case "Print"
            RaiseEvent LogMsg("(frmCheckSign.tbToolbar_ButtonClick) Print")
            OptPage1.DoPrintDlg
        
        Case Else
            MsgBox "This button has not been programmed"
    
    End Select

'    'If on first page then...
'    If optPage1.PageNumber = 1 Then
'        tbToolbar.Buttons("Previous").Enabled = False
'        tbToolbar.Buttons("Next").Enabled = True
'    'Else if on last page then...
'    ElseIf optPage1.PageNumber = optPage1.PageCount Then
'        tbToolbar.Buttons("Previous").Enabled = True
'        tbToolbar.Buttons("Next").Enabled = False
'    Else 'on a middle page
'        tbToolbar.Buttons("Previous").Enabled = True
'        tbToolbar.Buttons("Next").Enabled = True
'    End If

Exit Sub

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errToolBarClick, MODULE & ".tlbToolBar_ButtonClick", Err.Number & "-" & Err.Description)
    End Select
End Sub

Private Sub ShowImage(objObjID As OTOBJIDLib.ObjectID, Optional Name As String, Optional ClearViewer As Boolean)
'Load an IBPM object into the viewer. This assumes the objectid is already resolved.
'Inputs:
'ObjectKey - a subset of an eMedia ObjectId that is stored in workflow.
'Name - the name of the image that should appear on the image tab.
'ClearViewer - True causes all existing images in the viewer to be cleared before
'   loading the new one. False (default) cases the image to be added as a new
'   tab in the viewer.

Dim bReturn As Boolean      'function return value
Dim i As Integer            'loop counter
Dim datStart As Date        'image retrieval start time hack

'Enable error trap
    On Error GoTo errHandler

'Init the viewer with the user token
    Set OptPage1.UserToken = mobjUserToken

'Load the object into the viewer
    RaiseEvent LogMsg("(frmCheckSign.ShowImage) Loading page into viewer...")
    datStart = Now
    bReturn = OptPage1.LoadOptikaPage(objObjID, 393434, 9, _
        mobjUserToken.UserName)
    RaiseEvent LogMsg("(frmCheckSign.ShowImage) Page Loaded=" & bReturn)
    RaiseEvent LogMsg("(frmCheckSign.ShowImage) Image loaded, Image Retrieval Time = " & DateDiff("s", datStart, Now) & " secs")

''Fit image to viewer width
'    optPage1.FitSides = True
'
'Fit image to viewer size
    OptPage1.FitWindow = True

'Check page count
'    Debug.Print optPage1.PageCount
'    Debug.Print optPage1.PageNumber
'    Debug.Print optPage1.PageComponents
'    If optPage1.PageCount > 1 Then
'        'tbToolBar.Buttons("Previous").Enabled = True
'        tbToolbar.Buttons("Next").Enabled = True
'    Else
'        tbToolbar.Buttons("Previous").Enabled = False
'        tbToolbar.Buttons("Next").Enabled = False
'    End If
'Update paging buttons
    If OptPage1.IsNextPage = True Then
        tbToolBar.Buttons("Next").Enabled = True
    Else
        tbToolBar.Buttons("Next").Enabled = False
    End If

    If OptPage1.IsPrevPage = True Then
        tbToolBar.Buttons("Previous").Enabled = True
    Else
        tbToolBar.Buttons("Previous").Enabled = False
    End If

'Reset the page navigation because of page count bug in viewer control
    For i = 1 To 50
        OptPage1.PrevPage
    Next i

'Update page count control
    'NOTE - there is a bug in the viewer; the count is not correct
    'txtCount.Text = "Page " & optPage1.PageNumber & " of " & optPage1.PageCount
    mintPageNumber = 1
    mintPageCount = objObjID.PageCount
    txtCount.Text = "Page " & mintPageNumber & " of " & mintPageCount
    'txtCount.Text = "Page " & optPage1.PageNumber & " of ?"
    RaiseEvent LogMsg("(frmCheckSign.ShowImage) Page " & mintPageNumber & " of " & mintPageCount & " now displayed")

Exit Sub

errHandler:
    Select Case Err.Number
        Case -2147467261    'can't resolve object key
            RaiseEvent LogMsg("(frmCheckSign.ShowImage) Can't resolve object key: " & objObjID.WorkFlowValue)
            MsgBox "Unable to display this type of image", vbInformation + vbOKOnly, "Show Image"
            Screen.MousePointer = vbDefault
            Err.Clear
        Case -2147467259    'Failed to Auto Resolve ObjectID in Put WorkFlowValue
                            'Property. Unspecified Error
            RaiseEvent LogMsg("(frmCheckSign.ShowImage)" & Err.Description)
            RaiseEvent LogMsg("(frmCheckSign.ShowImage) Can't resolve object key: " & objObjID.WorkFlowValue)
            MsgBox "Unable to retrieve this image for display at this time ", vbInformation + vbOKOnly, "Show Image"
            Screen.MousePointer = vbDefault
            Err.Clear
        Case Else
            Call RaiseError(errGeneric, MODULE & ".ShowImage", Err.Number & "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub UnloadImage()
'Unload the image from the viewer.

'Enable error trap
    On Error GoTo errHandler
        
    RaiseEvent LogMsg("(frmCheckSign.UnloadImage) Unloading image from viewer")
    OptPage1.UnloadPage
    OptPage1.Refresh
    'Me.Refresh
    
Exit Sub

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errGeneric, MODULE & ".tlbToolBar_ButtonClick", Err.Number & "-" & Err.Description)
    End Select
End Sub

Private Function ExecuteNamedQuery(strInvoice As String, strSupplierNum As String) _
    As ADODB.Recordset
'Run an IBPM Named query to locate the specified invoice document.
'Return a recordset containing the matching document record(s).

'This requires a project reference to OTQuery.dll
  Dim objQStat As OTCONTEXTLib.QueryStatus
  Dim datStart As Date
  Dim intResultSize As Integer
  Dim objQConnection As OTCONTEXTLib.QueryConnection
  Dim objQManager As OTCONTEXTLib.QueryConnectionManager
  Dim objNamedQueries As New OTCONTEXTLib.NamedQueries
  Dim objNamedQuery As OTCONTEXTLib.NamedQuery
  Dim objNamedQueryCondition As OTCONTEXTLib.NamedQueryCondition
  objNamedQueries.UserToken = mobjUserToken
  objNamedQueries.Refresh

  Set objNamedQuery = objNamedQueries(cboApps.Text)
  RaiseEvent LogMsg("(frmCheckSign.ExecuteNamedQuery) Executing named query: " & objNamedQuery.Name)
  
  For Each objNamedQueryCondition In objNamedQuery.Conditions
    If objNamedQueryCondition.DisplayStyle = otPresentationAttributeDisplayStyle_Prompt _
       Or objNamedQueryCondition.DisplayStyle = otPresentationAttributeDisplayStyle_PromptReq Then
        RaiseEvent LogMsg("(frmCheckSign.ExecuteNamedQuery) Parameter: " & objNamedQueryCondition.LeftValue)
       'Set the RightValues...
        If (InStr(1, objNamedQueryCondition.LeftValue, "InvoiceNum") > 0) Then
           objNamedQueryCondition.RightValue = strInvoice
           RaiseEvent LogMsg("(frmCheckSign.ExecuteNamedQuery) Setting RightValue to: " & strInvoice)
        End If
       
        If (InStr(1, objNamedQueryCondition.LeftValue, "SupplierNum") > 0) Then
           objNamedQueryCondition.RightValue = strSupplierNum
           RaiseEvent LogMsg("(frmCheckSign.ExecuteNamedQuery) Setting RightValue to: " & strSupplierNum)
        End If
        
    End If

Next
 
 
'Instantiate and init query connection
    Set objQManager = New OTCONTEXTLib.QueryConnectionManager
    objQManager.UserToken = mobjUserToken
    Set objQConnection = objQManager.CreateConnection
    
'Set a default result set size
    intResultSize = 5

'Execute search
    RaiseEvent LogMsg("(frmCheckSign.ExecuteNamedQuery) Executing query...")
    datStart = Now
    objQConnection.ExecuteQuery objNamedQuery, otQueryStyleUnmanagedRecordset, _
        intResultSize
    RaiseEvent LogMsg("(frmCheckSign.ExecuteNamedQuery) Query results returned in " & DateDiff("s", datStart, Now) & " secs")
    
'Return the query results
    Set ExecuteNamedQuery = objQConnection.GetResults(otReturnAll, objQStat)
    
'Clean up
    
    Set objQManager = Nothing
    Set objQConnection = Nothing
 
Exit Function

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errExecuteNamedQuery, MODULE & ".ExecuteNamedQuery", Err.Number & "-" & Err.Description)
    End Select
End Function

Private Sub LoadInvoice()
'Get the selected invoice number and associate supplier number, find the invoice
'and load it into the viewer.

Dim strInvoice As String
Dim strCheckInfo As String
Dim udtInvoiceDat As udtInvoice
Dim intComma As Integer
Dim strInvoiceKey As String
    
'Enable error trap
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    frmMain.sbStatusBar.Panels("Status").Text = "Loading image..."
    
'Unload any page in the viewer in case the new image is not loaded
    Call UnloadImage
    
    If blnGoodLogin = True Then
        strInvoice = trvChecks.SelectedItem.Text
        udtInvoiceDat.InvoiceNum = strInvoice
        strInvoiceKey = trvChecks.SelectedItem.key
        'strCheckInfo = trvChecks.SelectedItem.Parent.Text
        'intComma = InStr(1, strCheckInfo, ",")
        'udtInvoiceDat.SupplierNum = Left(strCheckInfo, intComma - 1)
        intComma = InStr(2, strInvoiceKey, "-")
        udtInvoiceDat.SupplierNum = Mid(strInvoiceKey, 2, intComma - 2)
        RaiseEvent LogMsg("(frmCheckSign.ExecuteNamedQuery) Loading invoice '" & strInvoice & "' for Supplier Number '" & udtInvoiceDat.SupplierNum)
        Call FindAndDisplayInvoice(udtInvoiceDat)
    End If

    Screen.MousePointer = vbNormal
    frmMain.sbStatusBar.Panels("Status").Text = ""
Exit Sub

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errLoadInvoice, MODULE & ".LoadInvoice", Err.Number & "-" & Err.Description)
    End Select
End Sub

Private Function LeadingCommas(strText As String) As Long
'Return the number of consecutive leading commas in the input string.

Dim lngCount As Long
Dim strChar As String

'Initialize count to indicate to commas found
    lngCount = 0
    
'Loop to test each character in string beginning at first
    Do
        'Get the next character
        strChar = Mid(strText, lngCount + 1, 1)
        'If character is a string then...
        If strChar = "," Then
            'Count it
            lngCount = lngCount + 1
        End If
    Loop While strChar = ","
    
'Return the result
    LeadingCommas = lngCount
    
Exit Function

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errGeneric, MODULE & ".LeadingCommas", Err.Number & "-" & Err.Description)
    End Select
End Function

Private Function LDepad(strField) As String
'"De-pad" the string by removing leading zeros

Dim strNew As String
Dim dblNew As String
Dim i As Integer

    If IsNumeric(strField) Then
        dblNew = CDbl(strField)
        LDepad = CStr(dblNew)
    Else
        strNew = strField
        'If value is negative then
        If Left(strField, 1) = "-" Then
            For i = 1 To Len(strField)
                If Left(strNew, i) = "0" Then
                    strNew = Right(strNew, Len(strNew) - 1)
                End If
            Next i
        Else
            For i = 1 To Len(strField)
                If Left(strNew, i) = "0" Then
                    strNew = Right(strNew, Len(strNew) - 1)
                End If
            Next i
        End If
    
        LDepad = strNew
    End If


End Function

Private Function RDepad(strField) As String
'"De-pad" the string by removing trailing zeros

Dim strNew As String
Dim dblNew As String
Dim i As Integer

    If IsNumeric(strField) Then
        dblNew = CDbl(strField)
        dblNew = Format(dblNew, "########################################.00")
        RDepad = CStr(dblNew)
    Else
        strNew = strField
        For i = 1 To Len(strField)
            If Right(strNew, i) = "0" Then
                strNew = Left(strNew, Len(strNew) - 1)
            End If
        Next i
        RDepad = strNew
    End If
    
End Function

Private Sub RaiseError(ErrorNumber As Long, Source As String, Description As String)
'Log and raise the error.

'Log the error
    RaiseEvent LogMsg("(frmCheckSign.RaiseError) Error in " & Source & ": " & ErrorNumber & " - " & Description)
    
'Raise an error up to the client
    Err.Raise ErrorNumber, Source, Description
    
End Sub




Private Sub trvChecks_NodeClick(ByVal Node As MSComctlLib.Node)
  'View the selected invoice in the tree

  'Enable error trap
    On Error GoTo errHandler
        
'Make sure an invoice is selected in the tree view
    If trvChecks.SelectedItem.SelectedImage = 2 Then
        RaiseEvent LogMsg("(frmCheckSign.cmdView_Click) Locating and viewing invoice: " & trvChecks.SelectedItem.Text)
        Call LoadInvoice
    End If
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".trvChecks_NodeClick", Err.Number & "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub


