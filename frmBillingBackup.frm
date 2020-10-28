VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBillingBackup 
   Caption         =   "Billing Backup"
   ClientHeight    =   11025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15225
   Icon            =   "frmBillingBackup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11025
   ScaleWidth      =   15225
   Begin VB.Frame fraExport 
      Caption         =   "Export"
      Height          =   5535
      Left            =   8880
      TabIndex        =   6
      Top             =   4560
      Width           =   6255
      Begin VB.Frame fraExportFormat 
         Caption         =   "File Format"
         Height          =   2895
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   2295
         Begin VB.OptionButton optExportType 
            Caption         =   "PDF (optional)"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   27
            Tag             =   "8"
            ToolTipText     =   "Each document is written to a portable document format file"
            Top             =   2520
            Width           =   2055
         End
         Begin VB.OptionButton optExportType 
            Caption         =   "EMF"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   26
            Tag             =   "7"
            ToolTipText     =   "Each individual goes into an Enhanced Metafile"
            Top             =   2160
            Width           =   2055
         End
         Begin VB.OptionButton optExportType 
            Caption         =   "Bitmap (BMP)"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   25
            Tag             =   "1"
            ToolTipText     =   "Each individual page goes into a bitmap files"
            Top             =   1800
            Width           =   2055
         End
         Begin VB.OptionButton optExportType 
            Caption         =   "PCX"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   24
            Tag             =   "4"
            ToolTipText     =   "Each individual page goes into a PCX file"
            Top             =   1440
            Width           =   2055
         End
         Begin VB.OptionButton optExportType 
            Caption         =   "JPEG"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Tag             =   "3"
            ToolTipText     =   "Each individual page goes into a compressed JPEG"
            Top             =   1080
            Width           =   2055
         End
         Begin VB.OptionButton optExportType 
            Caption         =   "Single page TIFF"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Tag             =   "2"
            ToolTipText     =   "Each individual page goes into a compressed TIFF"
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton optExportType 
            Caption         =   "Multi-page TIFF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Tag             =   "2"
            ToolTipText     =   "All pages of the document go into a single, compressed TIFF"
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Opens a file dialog to select the export folder"
         Top             =   720
         Width           =   2535
      End
      Begin VB.DirListBox dirExport 
         Height          =   4815
         Left            =   2760
         TabIndex        =   17
         ToolTipText     =   "The export folder"
         Top             =   240
         Width           =   3375
      End
      Begin VB.DriveListBox drvExport 
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         ToolTipText     =   "The drive containing the export folder"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export Selected Documents..."
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4680
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Export Folder:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraSelect 
      Caption         =   "Select Documents for Export"
      Height          =   3975
      Left            =   8880
      TabIndex        =   5
      Top             =   480
      Width           =   6255
      Begin VB.Frame fraFilter 
         Caption         =   "Filter Selection"
         Height          =   2175
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   5655
         Begin VB.ComboBox cboCostType 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   28
            ToolTipText     =   "The cost types that are in the report"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtAmt 
            Height          =   285
            Left            =   3240
            TabIndex        =   15
            ToolTipText     =   "The dollar amount used for the filter"
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdSelFilter 
            Caption         =   "Select Filtered"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            ToolTipText     =   "Selects invoices in the treeview based on the filter chosen"
            Top             =   1560
            Width           =   1575
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Option3"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.OptionButton optCostType 
            Caption         =   "By Cost Type"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            ToolTipText     =   "Selects only those invoices with the specified Cost Type"
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton optAmt 
            Caption         =   "By Amount: Greater than or equal to $"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            ToolTipText     =   "Selects only those invoices that have a dollar amount equal to or greater than the specified amount"
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.CommandButton cmdUnselect 
         Caption         =   "Unselect All"
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         ToolTipText     =   "Unselects all invoices in the treeview"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "Select All"
         Height          =   375
         Left            =   480
         TabIndex        =   8
         ToolTipText     =   "Selects all invoices in the treeview"
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.ComboBox cboApps 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "This is the Invoice application that will be searched to find the documents to export"
      Top             =   0
      Width           =   3975
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Open a Job Cost Report..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      ToolTipText     =   "Opens a file dialog to select a job cost report"
      Top             =   0
      Width           =   2055
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   7920
      TabIndex        =   0
      ToolTipText     =   "This is the job cost report being processed"
      Top             =   0
      Width           =   7215
   End
   Begin MSComctlLib.TreeView trvInvoices 
      Height          =   9495
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Treeview containing the list of invoices in the report"
      Top             =   600
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   16748
      _Version        =   393217
      Style           =   6
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   7920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
            Picture         =   "frmBillingBackup.frx":23D2
            Key             =   "Invoice"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBillingBackup.frx":2824
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblApp 
      Caption         =   "Select the Division's Invoice Application:"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmBillingBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
' Class:        frmBillingBackup
' Author:       Wes Prichard, Optika
' Date:         March 2003
' Description:  Implements the billing backup export functionality.
' Edit History:
' 06/17/2004 - Wes Prichard, Optika
'   Modified ParseBillingReport procedure to recognize cost codes as a string pattern
'   rather than a numeric value. This is due to Weitz changing to a new cost code system.
'
' 08/18/2004 - Wes Prichard, Optika, v2.1.5
'   In order for the invoices to match up with the detail billing, the format needs to be
'   Cat code, cost code, cost type, invoice number, date, #unique id.
'   looking like:
'   02G, A101001U, 2450, 60404, 6-19-2004, #15003704.tif
'   Modified procedure ParseBillingReport to find and extract the cat code and include
'   it in the report array. Modified cmdExport_Click to put CatCode in the udtInvoice
'   structure and FindAndExportInvoice to create the exported file name in the new format.

' 08/26/2004 - Wes Prichard, Optika, v2.1.6
'   Added GL Date to the name of the exported file. GLDate added to udtInvoice. Modified
'   cmdExport to init udtInv.GLDate before calling FindAndExportInvoice. Modified
'   FindAndExportInvoice to include GL Date in the exported file's title.

' 10/08/2004 - Wes Prichard, Optika, v2.1.7
'   Added code to procedure FindAndExportInvoice to replace any illegal characters in the
'   final name for the exporeted file. This is because the file name include the invoice
'   number which can sometimes include slashes or potentiall other illegal characters.

' 10/15/2004 - Wes Prichard, Optika, v2.1.8
'   Modified procedure FindAndExportInvoice to change order of fields in name of exported
'   image file. The new order is Cat code, cost code, cost type, date, Invoice num,
'   #unique id.

' 10/15/2004 - Wes Prichard, Optika, v2.2.0
'   Modified to extract Supplier Number from new report format and use it to find invoice.
'   Added SupplierNum to udtInvoice.
'   Added Supplier Number to report array.
'   Modified procedure ParseBillingReport to extract from new format and
'   include Supplier Number.
'   Modified procedure FindAndExportInvoice to pass Supplier Number to ExecuteAdHocQuery
'   Modified procedure ExecuteAdHocQuery to use SupplierNum instead of JobNum in the
'   invoice query and removed Vendor Name.

' 01/04/2005, 01/18/205 - Wes Prichard, Optika, v2.2.1
'   Modified procedures that set file paths from registry settings to trap errors
'   resulting from invalid drives:
'   cmdBrowse_Click, Form_Load, cmdSelect_Click
'   Modified procedure ParseBillingReport to make the cost code key unique in the
'   treeview by adding the line number to the node key. This resolves an issue
'   with the JDE reports where they contain non-consecutive records for the same
'   cost code.
'   Also fixed a bug where the export report is not accurate if multiple exports are
'   run against the same report without reloading the report. Added code to
'   GenerateReports to check the export flag so that only records selected in the
'   treeview are included in the report. Also added code to clear the InAcorde flag
'   in the mstrReport array after the report is generated so that if a subsequent
'   export is cancelled, the report will be accurate.
'   Also added the Form_Unload procedure to release the Acorde license by logging
'   out.
'
' 02/28/2005 - Wes Prichard, Optika, v2.3.0
'   Upgraded code to work with Acorde version 4.0 SP1.
'   Modified procedure ParseBillingReport to only save report records that contained
'   the JDE doctype P2, PV, PD, or PM. Also added some log messages pertaining to
'   parsing.
'   Modified Form_Unload to skip logout of mobjUserToken is nothing to correct an error
'   that would occur if the user had not logged in and then closed the tool window.

' 02/28/2005 - Wes Prichard, Optika, v2.3.1
'   Corrected a bug where the Supplier Number was not getting passed into the Acorde search
'   and the incorrect image was being exported.
'   Modified cmdExport_Click to load the Supplier Number from the report array into the
'   udtInvoice structure.

' 07/26/2005 - Wes Prichard, Optika, v2.3.2
'   Modified code to support a new cost code format.
'   Modified procedure ParseBillingReport to validate the cost code field as a string
'   of 8 or fewer characters. (Previously, it was validated as always starting with a
'   letter followed by 4, 5, or 6 digits and maybe a letter.)

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
'Acorde SDK User Security 1.0 Type Library (OTUsrSec.dll)
'Acorde SDK Schema 1.0 Type Library (OTSchema.dll)
'Acorde SDK DocumentID 1.0 Type Library (OTDocID.dll)
'Acorde SDK Export Manager 1.0 Type Library (OTExportManager.dll)
'Acorde SDK Query 1.0 Type Library (OTQuery.dll)
'Acorde SDK Object ID 1.0 Type Library (OTObjID.dll)
'*****************************************************************************************

Option Explicit

Private Const ERRORBASE = ErrorBase8    'see modErrorHandling - used to avoid
                                        '   overlapping error numbers
Private Const MODULE = "frmBillingBackup"     'used in reporting errors

'Class Error enumeration
Public Enum errCEP
    errNoPDF = ERRORBASE + 0
    errParseBillingReport = ERRORBASE + 1
'    errLoad = ERRORBASE + 2
    errGetInvoiceApps = ERRORBASE + 3
    errFindAndExportInvoice = ERRORBASE + 4
'    errToolBarClick = ERRORBASE + 5
   errExecuteAdHocQuery = ERRORBASE + 6
'    errLoadInvoice = ERRORBASE + 7
    errGeneric = ERRORBASE + 8
'    errDBRollbackTrans = ERRORBASE + 9
'    errDBRollbackTrans = ERRORBASE + 10
End Enum

'Public event declarations
Public Event Status(Message As String)  'status to the user
Public Event LogMsg(Message As String)  'message to the log
Public Event MousePtr(PointerStyle As MousePointerConstants)

'Structure for carrying invoice query parameters and results
Private Type udtInvoice
    InvoiceNum As String    'invoice number
    JobNum As String        'job number
    InAcorde As Boolean     'True if invoice found in Acorde
    CostCode As String      'cost code from report
    CostType As String      'cost type from report
    ExportPath As String    'fully qualified reference to exported file
    Vendor As String        'Vendor name from expense description
    CatCode As String       'Category Code
    GLDate As String        'General Ledger Date
    SupplierNum As String   'Supplier Number
End Type

Const EXPORTFOLDER = "AcordeData" 'target folder for Export Manager

'Report Array index constants
Const REPTARRAYCOLS = 13  'indicates how many columns are in the array
'These are the indexes of the fields in the first dimension of the strReport array
Const ICOSTCODE = 0
Const ICOSTTYPE = 1
Const ICCDESC = 2
Const IEXPDESC = 3
Const IINVNUM = 4
Const IGLDATE = 5
Const IGLAMT = 6
Const IEXPORT = 7
Const IINACORDE = 8
Const IEXPPATH = 9
Const ICATCODE = 10
Const ISUPNUM = 11
Const IDOCTYPE = 12

'Module-scope variables
Dim mstrReportFile As String    'the path and name of the report file
Dim mstrReport() As String       'dynamic array of report content without blank columns
'The following declaration requires a project reference to Acorde SDK User Security 1.0 type Library (OTUsrSec.dll)
Dim mobjUserToken As OTACORDELib.UserToken     'Acorde login user token
Dim mblnGoodLogin As Boolean     'indicates if login was successful
Dim mstrJobNum As String        'job number for report
Dim WithEvents mfrmExp As frmExportProgress
Attribute mfrmExp.VB_VarHelpID = -1
Dim mblnCancelExport As Boolean 'indicates export to be cancelled if true
'Requires a project reference to Acorde SDK User Security 1.0 type Library (OTUsrSec.dll)
Dim mobjUser As OTACORDELib.User    'Acorde user object

Public Sub Display()
'Custom entry point for loading and displaying form so that modality and other
'properties can be controlled.
 
'Make me visible
    RaiseEvent LogMsg("(" & MODULE & ".Display) Entering procedure")
    Me.Show vbModeless  'this triggers Form_Load procedure
    Me.WindowState = vbMaximized
    Me.Refresh
    
End Sub

Private Sub cmdBrowse_Click()
'Allow the user to browse for the export folder using a common dialog.
'Update the directory control with the selection.

Dim strPath As String
Dim intSubString        'position of substring within string

'Enable error trap
    On Error GoTo errHandler     'Trap it if user cancels dialog

'Open a file dialog to select the output file
    dlgFile.Flags = cdlOFNExplorer
    dlgFile.CancelError = True
    dlgFile.Filter = "All (*.*)|*.*|TXT (*.txt)|*.txt"
    On Error Resume Next
    dlgFile.InitDir = GetSetting(App.EXEName, "Files", "ExportPath", App.Path)
    If Err.Number <> 0 Then
        'Assume path in registry is not valid
        dirExport.Path = "C:\"
    End If
    On Error GoTo errHandler
    dlgFile.FileName = "Export.txt"
    dlgFile.ShowOpen
    
'Get file information selected by user
    strPath = Left(dlgFile.FileName, InStr(1, dlgFile.FileName, dlgFile.FileTitle) - 1)
    SaveSetting App.EXEName, "Files", "ExportPath", strPath
    intSubString = InStr(1, dlgFile.FileName, dlgFile.FileTitle)
    dirExport.Path = Left(dlgFile.FileName, intSubString - 2)
    intSubString = InStr(1, dlgFile.FileName, ":")
    drvExport.Drive = Left(dlgFile.FileName, intSubString)
Exit Sub

errHandler:
    Select Case Err.Number
        Case cdlCancel
            Exit Sub
        Case Else
            Call RaiseError(errGeneric, MODULE & ".cmdBrowse_Click", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub dirExport_Change()
'Save the new dir in the registry

'Enable error trap
    On Error GoTo errHandler

    SaveSetting App.EXEName, "Files", "ExportPath", dirExport.Path
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".dirExport_Change", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub drvExport_Change()
'Change the associated directory control to be the selected drive.

'Enable error trap
    On Error GoTo errHandler

    dirExport.Path = drvExport.Drive

Exit Sub

errHandler:
    Select Case Err.Number
        Case 68
            'Invalid drive
            MsgBox "'" & Err.Description & "' Please select another drive"
            drvExport.Drive = dirExport.Path
            Exit Sub
        Case Else
            Call RaiseError(errGeneric, MODULE & ".drvExport_Change", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub Form_Load()
'Initialize the form

Dim strUser As String
Dim strPassword As String
Dim ofrmLogin As frmLogin
'Requires a project reference to Acorde SDK User Security 1.0 type Library (OTUsrSec.dll)
'Dim objUser As OTACORDELib.User
Dim objControl As Control
Dim i As Integer

'Enable error trap
    On Error GoTo errHandler

'Setup treeview
    trvInvoices.SingleSel = False   'True      'expand node when selected
    trvInvoices.Style = 7   'tvwPictureText
    trvInvoices.ImageList = imgTree
    trvInvoices.LineStyle = tvwTreeLines
    trvInvoices.Checkboxes = True
    trvInvoices.Indentation = 300   'default = 566.9291

'Disable buttons until the treeview is loaded
    For i = 1 To Me.Controls.Count
        On Error Resume Next
        If (Me.Controls(i).Container.Name = fraSelect.Name) Or (Me.Controls(i).Container.Name = fraExport.Name) Then Me.Controls(i).Enabled = False
        If Err.Number <> 438 And Err.Number <> 0 Then Err.Raise 1, , Err.Description
        On Error GoTo errHandler
    Next i
    
'Disable the Select filtered button until an option is selected
    cmdSelFilter.Enabled = False

'Select the default export format
    optExportType(0).Value = True   'Multi-page TIFF
    
'Init the export path
    On Error Resume Next
    dirExport.Path = GetSetting(App.EXEName, "Files", "ExportPath", App.Path)
    If Err.Number <> 0 Then
        'Assume path in registry is not valid
        dirExport.Path = "c:\"
    End If
    On Error GoTo errHandler

'Login to Acorde
    Do While mobjUserToken Is Nothing
        'Get the last user name
        strUser = GetSetting(App.Title, "Settings", "Acorde User")
        'Display the login form
        Set ofrmLogin = New frmLogin
        RaiseEvent LogMsg("(" & MODULE & ".Form_Load) Opening Login dialog...")
        ofrmLogin.Display strUser
        strUser = ofrmLogin.UserName
        strPassword = ofrmLogin.Password
        
        'If user did not cancel login then...
        If ofrmLogin.OK = True Then
            Set ofrmLogin = Nothing
            Me.Refresh
            
            'Save the name of the successful login
            SaveSetting App.Title, "Settings", "Acorde User", strUser
            
            'Call the Acorde login
            Me.MousePointer = vbHourglass
            Set mobjUser = New OTACORDELib.User
            RaiseEvent LogMsg("(" & MODULE & ".Form_Load) User completed Login dialog - logging into Acorde as user: " & strUser)
            Set mobjUserToken = mobjUser.Login(strUser, strPassword, False) 'hide dialog
'            Set mobjUserToken = mobjUser.Login(strUser, strPassword, True) 'hide dialog
            Me.MousePointer = vbDefault
            
            'Notify user if login unsuccessful
            If mobjUserToken Is Nothing Then
                MsgBox "Login was unsuccessful. Check user name and remember " & _
                    "that password are case-sensitive.", vbExclamation + vbOKOnly, _
                    "Acorde Login Failed"
            End If
        Else
            'user cancelled the login dialog
            RaiseEvent LogMsg("(" & MODULE & ".Form_Load) User cancelled Login dialog...")
            Exit Do
            Set ofrmLogin = Nothing
        End If
    
    Loop

'If login successful then...
    If Not (mobjUserToken Is Nothing) Then
        RaiseEvent LogMsg("(" & MODULE & ".Form_Load) Acorde login successful")
        mblnGoodLogin = True
    
        'Initialize the application combo
        Call GetInvoiceApps(cboApps)
        Me.Refresh
        'cboApps.SetFocus   'can't do this until form is visible
    
        'cboApps.Text = GetSetting(App.Title, "Settings", "Application")
        'Enable the controls
'        fraSelect.Enabled = True
'        fraExport.Enabled = True
        
    Else    'login cancelled
        RaiseEvent LogMsg("(" & MODULE & ".Form_Load) Acorde login failed")
        mblnGoodLogin = False
        MsgBox "You will not be able to export documents because the " & _
            "login to Acorde was not successful.", vbExclamation + vbOKOnly, _
            "Acorde Login Result"
        'Enable selected controls so that user can only view a report file
        cmdSelect.Enabled = True    'allow user to view a report file
    End If
    
    RaiseEvent LogMsg("(" & MODULE & ".Form_Load) Form Load complete")
Exit Sub

errHandler:
    Select Case Err.Number
        Case -2147220640
            MsgBox Err.Description, vbInformation + vbOKOnly, "Acorde Login Error"
            Resume Next
        Case -2147023570    'Failed Login method.
            MsgBox Err.Description, vbInformation + vbOKOnly, "Acorde Login Error"
            Resume Next
        Case -2147467259    'The Domain property is not set or is invalid.
            MsgBox Err.Description, vbInformation + vbOKOnly, "Acorde Login Error"
            Resume Next
        Case Else
            Call RaiseError(errLoad, MODULE & ".Form_Load", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub Form_Resize()
'Scale the controls when the form is resized.

'Enable error trap
    On Error GoTo errHandler
        
'Keep the bottom of treeview at the bottom of the form
    If Me.Height - trvInvoices.Top - 400 > 0 Then
        trvInvoices.Height = Me.Height - trvInvoices.Top - 400
    End If
    
'TODO resize frames
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".Form_Resize", Err.Number & _
                "-" & Err.Description)
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
    RaiseEvent LogMsg("(" & MODULE & ".cboApps_Click) Invoice application '" & cboApps.Text & "' selected")
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".cboApps_Click", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub cmdSelect_Click()
'Open a file dialog and get the fully qualified name of the input report
'to be processed. Then cause the report to be parsed and displayed.

Dim strPath As String
Dim i As Integer
Dim j As Long

'Enable error trap
    On Error GoTo errHandler     'Trap it if user cancels dialog


'Open a file dialog to select the output file
    dlgFile.Flags = cdlOFNExplorer
    dlgFile.CancelError = True
    dlgFile.Filter = "All (*.*)|*.*|CSV (*.csv)|*.csv"
    On Error Resume Next
    dlgFile.InitDir = GetSetting(App.EXEName, "Files", "ReportPath", App.Path)
    'dlgFile.InitDir = "c:"  '@@@
    If Err.Number <> 0 Then
        'Assume path in registry is not valid
        dlgFile.InitDir = "c:\"
    End If
    On Error GoTo errHandler
    dlgFile.ShowOpen
    
'Get file information selected by user
    strPath = Left(dlgFile.FileName, InStr(1, dlgFile.FileName, dlgFile.FileTitle) - 1)
    SaveSetting App.EXEName, "Files", "ReportPath", strPath
    mstrReportFile = dlgFile.FileName
    txtFilePath.Text = dlgFile.FileName
    
'Clear the tree view
    trvInvoices.Nodes.Clear
    Me.Refresh

'Parse the report and load the tree view
    RaiseEvent LogMsg("(" & MODULE & ".cmdSelect_Click) Parsing check report file: " & mstrReportFile)
    ParseBillingReport (mstrReportFile)
    Me.Refresh

'Enable the report-related controls
    If mblnGoodLogin = True Then
        For i = 1 To Me.Controls.Count
            On Error Resume Next
            'If the control is in one of the frames then...
            If (Me.Controls(i).Container.Name = fraSelect.Name) Or _
                (Me.Controls(i).Container.Name = fraExport.Name) Then _
                Me.Controls(i).Enabled = True
            If (Err.Number <> 438) And (Err.Number <> 0) Then
                Err.Raise 1, , Err.Description
            End If
            On Error GoTo errHandler
        Next i
    End If

'Load cost type combo box from array
    cboCostType.Clear
    For j = 1 To UBound(mstrReport, 2)
        If IsDuplicate(mstrReport(ICOSTTYPE, j), ICOSTTYPE, j) = False Then
            cboCostType.AddItem mstrReport(ICOSTTYPE, j)
        End If
    Next j

Exit Sub

errHandler:
    Select Case Err.Number
        Case cdlCancel
            Exit Sub
        Case Else
            Call RaiseError(errGeneric, MODULE & ".cmdSelect_Click", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub cmdExport_Click()
'Export the selected documents.

Dim intResponse As Integer
Dim strExportFolder As String
Dim i As Integer
Dim lngRow As Long
Dim udtInv As udtInvoice
Dim lngToExport As Long     'number of documents selected for export
Dim lngExported As Long     'number of documents actually exported
Dim blnSomethingSelected As Boolean 'indicates if at least one invoice is selected
'Requires project reference to Microsoft Scripting Runtime (scrrun.dll)
Dim objFSO As FileSystemObject
Dim strFolder As String
Dim strExportTempFolder As String
Dim lngInvoicesToExport As Long 'number of invoices to export for progress bar
Dim datStart As Date             'reference time
Dim sngUnitTime As Single       'average time per query/export operation
Dim lngRemainingUnits As Long   'number of invoices left to export
Dim intPos As Integer           'position of character in string

'Enable error trap
    On Error GoTo errHandler

'Get the export path
    strExportFolder = dirExport.Path
        
'Are you sure?
    intResponse = MsgBox("Export the selected documents to " & _
        strExportFolder & "?", vbOKCancel + vbQuestion, "Export")
    If intResponse <> vbOK Then
        Exit Sub
    
    Else
        'Export...
        RaiseEvent LogMsg("(" & MODULE & ".cmdExport_Click) Exporting to " & strExportFolder)
        'Save the export path as a user preference
        SaveSetting App.EXEName, "Files", "ExportFolder", strExportFolder
                
        'Set the cancel flag to false to allow the user to cancel
        mblnCancelExport = False
        
        'Update the array with the selected documents
        lngRow = 1
        blnSomethingSelected = False
        For i = 1 To trvInvoices.Nodes.Count    'for all nodes in the treeview...
            'If this is an invoice node then...
            If trvInvoices.Nodes(i).key = vbNullString Then
                'If this invoice node is checked then...
                If trvInvoices.Nodes(i).Checked = True Then
                    mstrReport(IEXPORT, lngRow) = True
                    blnSomethingSelected = True
                    lngInvoicesToExport = lngInvoicesToExport + 1
                Else
                    mstrReport(IEXPORT, lngRow) = False
                End If
                'Increment the array row index
                lngRow = lngRow + 1
            End If
        Next i
        
        'If at least one invoice was selected for export then...
        If blnSomethingSelected = True Then
        
            'Display the export form
            Set mfrmExp = New frmExportProgress
            'mfrmExp.Top = (Me.Height / 2) - (mfrmExp.Height / 2)
            'mfrmExp.Left = (Me.Width / 2) - (mfrmExp.Width / 2)
            mfrmExp.Top = 1725  '(Me.Height / 2) + (mfrmExp.Height / 2)
            mfrmExp.Left = 3690 '(Me.Width / 2) + (mfrmExp.Width / 2)
            mfrmExp.Display
            Me.WindowState = vbMinimized
            Me.Enabled = False
            DoEvents    'to allow resize events to fire
            
            'Initialize the progress bar
            mfrmExp.prgExport.Min = 0
            mfrmExp.prgExport.Max = lngInvoicesToExport
            mfrmExp.prgExport.Value = 0
            mfrmExp.prgExport.Visible = True
            'mfrmexp.prgexport.Refresh
            
            'start the progress timer
            datStart = Now
            
            'Query for and export the documents
            lngToExport = 0
            lngExported = 0
            Me.MousePointer = vbHourglass
            For i = 1 To UBound(mstrReport, 2)  'for all invoices in the array...
                'If the invoice is to be exported then...
                If mstrReport(IEXPORT, i) = CStr(True) Then
                    lngToExport = lngToExport + 1
                    mfrmExp.lblCount = "Processing " & CStr(lngToExport) & " of " & CStr(lngInvoicesToExport)
                    mfrmExp.Refresh
                    
                    udtInv.InvoiceNum = mstrReport(IINVNUM, i)
                    udtInv.JobNum = mstrJobNum
                    udtInv.CostCode = mstrReport(ICOSTCODE, i)
                    udtInv.CostType = mstrReport(ICOSTTYPE, i)
                    udtInv.Vendor = mstrReport(IEXPDESC, i)
                    udtInv.CatCode = mstrReport(ICATCODE, i)
                    udtInv.GLDate = mstrReport(IGLDATE, i)
                    udtInv.SupplierNum = mstrReport(ISUPNUM, i)
                    RaiseEvent LogMsg("(" & MODULE & ".cmdExport_Click) InvNum = " & _
                        udtInv.InvoiceNum & ", JobNum = " & udtInv.JobNum & ", CostCode = " & _
                        udtInv.CostCode & ", CostType = " & udtInv.CostType & ", Vendor = " & _
                        udtInv.Vendor & ", CatCode = " & udtInv.CatCode & ", GLDate = " & _
                        udtInv.GLDate & ", SupplierNum = " & udtInv.SupplierNum)
                    
                    'TODO - have it check to see if already in acorde and exported
                    'udtInv.InAcorde = mstrReport(IINACORDE, i) 'might be set if previously processed
                    strFolder = FindAndExportInvoice(udtInv)
                    
                    'Save report information
                    mstrReport(IINACORDE, i) = udtInv.InAcorde
                    mstrReport(IEXPPATH, i) = udtInv.ExportPath
                    If strFolder <> vbNullString Then
                        strExportTempFolder = strFolder
                        'Strip off the last subfolder
                        intPos = InStrRev(strExportTempFolder, "\")
                        strExportTempFolder = Left(strExportTempFolder, intPos - 1)
                    End If
                    
                    'Update the count of exported docs
                    If udtInv.InAcorde = True Then lngExported = lngExported + 1
                    
                    'Update the progress form
                    mfrmExp.prgExport.Value = mfrmExp.prgExport.Value + 1
                    mfrmExp.prgExport.Refresh
                    mfrmExp.lblElapsed = Format((DateDiff("s", datStart, Now) / 60), "##0.0")
                    'Calculate time remaining
                    sngUnitTime = (mfrmExp.lblElapsed * 60) / mfrmExp.prgExport.Value
                    lngRemainingUnits = mfrmExp.prgExport.Max - mfrmExp.prgExport.Value
                    mfrmExp.lblRemaining = Format((sngUnitTime * lngRemainingUnits) / 60, "##0.0")
                    mfrmExp.Refresh
                End If
                
                'Check for cancel
                DoEvents
                If mblnCancelExport = True Then
                    Exit For
                End If
            Next i
            Me.MousePointer = vbDefault
            
            'Delete the temporary export folder used by export server
            'now that all the files have been moved out of it.
            If lngExported > 0 Then
                Set objFSO = New FileSystemObject
                If objFSO.FolderExists(strExportFolder & "\" & EXPORTFOLDER) = True Then
                    objFSO.DeleteFolder strExportFolder & "\" & EXPORTFOLDER, True
                End If
            End If
            
            'Generate reports
'            GenerateReports strExportFolder & "\ExportReport.csv"
            GenerateReports strExportFolder & "\ExportReport.txt", mblnCancelExport
            
            'Generate index html file
            Select Case True
                Case optExportType(0).Value  'multipage TIFF
                Case optExportType(1).Value  'single page TIFF
                Case optExportType(2).Value  'JPEG
                    'GenerateIndex strExportFolder & "\Invoice Image Index.htm"
                    'TODO modify GenerateReports for single page exports
                Case optExportType(3).Value  'PCX
                    'GenerateIndex strExportFolder & "\Invoice Image Index.htm"
                    'TODO modify GenerateReports for single page exports
                Case optExportType(4).Value  'BMP
                Case optExportType(5).Value  'EMF
                Case optExportType(6).Value  'PDF
                    GenerateIndex strExportFolder & "\Invoice Image Index.htm"
            End Select
            
            'Notify User
            If mblnCancelExport = True Then
                MsgBox "Export cancelled. " & lngExported & " of " & lngToExport & _
                    " documents exported." & vbCrLf & _
                    "See the complete report in the export folder.", vbInformation + vbOKOnly, _
                    "Export"
                RaiseEvent LogMsg("(" & MODULE & ".cmdExport_Click) Export cancelled. " & _
                    lngExported & " of " & lngToExport & " documents exported.")
            Else
                MsgBox "Export complete! " & lngExported & " of " & lngToExport & _
                    " documents exported." & vbCrLf & _
                    "See the complete report in the export folder.", vbInformation + vbOKOnly, _
                    "Export"
                RaiseEvent LogMsg("(" & MODULE & ".cmdExport_Click) Export complete. " & _
                    lngExported & " of " & lngToExport & " documents exported.")
            End If
            
            mfrmExp.Hide
            Set mfrmExp = Nothing
            Me.WindowState = vbMaximized
            Me.Enabled = True
        Else
            MsgBox "There are no invoices selected for export.", _
                vbInformation + vbOKOnly, "Export"
            RaiseEvent LogMsg("(" & MODULE & ".cmdExport_Click) There are no invoices selected for export.")
        End If
        
    End If

Exit Sub

errHandler:
    Select Case Err.Number
        Case errNoPDF   'no PDF export engine
            MsgBox "The export has been terminated because the requested output " & _
                "format could not be provided", vbExclamation + vbOKOnly, "Export"
            Me.MousePointer = vbDefault
            mfrmExp.prgExport.Visible = False
            'exit
        Case Else
            Call RaiseError(errGeneric, MODULE & ".cmdExport_Click", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub cmdSelAll_Click()
'Select all documents in the tree

Dim i As Integer

'Enable error trap
    On Error GoTo errHandler

    For i = 1 To trvInvoices.Nodes.Count
        CheckNode trvInvoices.Nodes(i), True
    Next i
    RaiseEvent LogMsg("(" & MODULE & ".cmdSelAll_Click) All invoices selected")
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".cmdSelAll_Click", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub cmdSelFilter_Click()
'Select nodes according to the filter selected

'Enable error trap
    On Error GoTo errHandler

    Select Case True
        Case optAmt.Value
            If IsNumeric(txtAmt.Text) Then
                Call SelectByAmount(txtAmt.Text)
                RaiseEvent LogMsg("(" & MODULE & ".cmdSelFilter_Click) Invoices selected by amount " & txtAmt.Text)
            Else
                MsgBox "You must enter an amount.", vbExclamation + vbOKOnly, "Filter by Amount"
                txtAmt.SetFocus
            End If
        
        Case optCostType.Value
            If IsNumeric(cboCostType.Text) Then
                Call SelectByCostType(cboCostType.Text)
                RaiseEvent LogMsg("(" & MODULE & ".cmdSelFilter_Click) Invoices selected by cost type " & cboCostType.Text)
            Else
                MsgBox "You must select a cost type.", vbExclamation + vbOKOnly, "Filter by Cost Type"
                cboCostType.SetFocus
            End If
        
        Case Option3.Value
    End Select

Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".cmdSelFilter_Click", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub cmdUnselect_Click()
'Uncheck all nodes

Dim i As Integer

'Enable error trap
    On Error GoTo errHandler

    For i = 1 To trvInvoices.Nodes.Count
        CheckNode trvInvoices.Nodes(i), False
    Next i
    RaiseEvent LogMsg("(" & MODULE & ".cmdUnselect_Click) All invoices unselected")
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".cmdUnselect_Click", Err.Number & _
                "-" & Err.Description)
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

Private Function ParseBillingReport(strFile As String)
'Parse the specified report and extract the invoice
'information into an array and populate the treeview.
'Inputs:
'strFile - the fully qualifed path and name of the report file to be processed

'Requires project reference to Microsoft Scripting Runtime (scrrun.dll)
Dim fs As New FileSystemObject  'used to access report file
Dim TextStreamIn As TextStream  'used to read report file
Dim strLine As String           'line from the report file
Dim strSplit() As String        'dynamic array containing all fields in one report line
Dim lngRow As Long              'strReport array row index
Dim objNodes As Nodes           'treeview nodes collection
Dim objNode As Node             'treeview node object
Dim strCostCodeKey As String    'key of current check node in treeview
Dim strCostTypeKey As String    'key of current cost type node in treeview
Dim intFirstQuote As Integer    'position of first quote character in report line
Dim intSecondQuote As Integer   'position of second quote character in report line
Dim strSub As String            'a substring used in parsing
Dim strFixed As String          'a string that has had the commas replaced
Dim blnInCostRec As Boolean     'state variable indicating the beginning line of a cost record has been processed
Dim strLastCostCode As String   'last cost code added to treeview
Dim lntLastCostType As Long     'last cost type added to treeview
Dim i As Long                   'loop counter
Dim strErr As String            'error number used in error handler
Dim strDesc As String           'error description used in error handler
Dim strCatCode As String        'current cat code extracted from report

Const KEY_TOPNODE = "T"
    
'Enable error trap
    On Error GoTo errHandler

'Show hourglass
    'Screen.MousePointer = vbHourglass
    'trvInvoices.MousePointer = vbHourglass
    RaiseEvent MousePtr(vbHourglass)
    
'Open the report file
    Set TextStreamIn = fs.OpenTextFile(strFile, ForReading, False)
    RaiseEvent LogMsg("(" & MODULE & ".ParseBillingReport) Report file '" & strFile & "' opened for reading")
       
'Proceed with parsing the report:
'3 lines per record
'Example:
',,,,,,,,,,W301003,2450
'01A,,,,,,,,,,,,P2,258567,06/03/04,,,82.39
',Pickup Truck Fuel,,,BANK OF AMERICA - PHOENIX,,,5-15-04 KB,,127474
'Note some records don't contain an invoice number. If this is the case, ignore the record.
'Also, if the record is for one of the following doctypes, keep it, otherwise ignore it.
'Keep: P2, PV, PD, and PM
'The parsing algorithm reads one line of the csv report at a time. It looks for the first line
'of a record and when found, sets a state variable (blnInCostRec). When a record is found, the
'desired data is extracted from the report and added to an array (mstrReport). The treeview
'control is loaded from the report array (mstrReport) after the complete record is extracted.

'10 leading blank fields is the cost code and cost type
'Cost Code - field 11
'Cost Type - field 12

'0 leading blank fields contains
'Category Code - field 1
'G/L Date - field 15
'Amount - field 18
'Note JDE doctype is in field 13.

'1 leading blank field contains
'Cost Code description - field 2
'Expense Description - field 5
'Invoice Num - field 8
'Supplier Number - field 10
'anything else is ignored

'If none of the above are true and Field 1 contains "Job:"
'Job Number - field 2

'Clear the job number for a new report
    mstrJobNum = vbNullString

'Init the array
    ReDim mstrReport(REPTARRAYCOLS, 0)   '1 row with 8 fields
    
'Clear the tree view
    For i = 1 To trvInvoices.Nodes.Count
        trvInvoices.Nodes.Remove i
    Next i
    
'Init state varialbe
    blnInCostRec = False
    
'Populate the report array with the desired fields from the file
    'Get the treeview nodes collection
    Set objNodes = trvInvoices.Nodes
    'Add a top-level node so that there is an expansion control with each child node
    Set objNode = objNodes.Add(, , KEY_TOPNODE, "Invoices")
    'For all lines in the file...
    Do While Not TextStreamIn.AtEndOfStream
        strLine = TextStreamIn.ReadLine
        Debug.Print TextStreamIn.Line; strLine
        
        'If line contains quotes then replace commas inside quotes
        ' so Split will work right
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
                strFixed = Replace(strLine, ",", "|", intFirstQuote + 1, 1)
                strLine = Left(strLine, intFirstQuote) & strFixed '& Right(strLine, Len(strLine) - intSecondQuote + 1)
            End If
        End If
        
        'Split the CSV field into an array
        strSplit = Split(strLine, ",")
        'Identify the content of the row by the number of leading commas (i.e. blank fields)
        Select Case LeadingCommas(strLine)
            Case 10  'doc row 1
                'If the field is no more than 8 characters then it is assumed to be a cost code
                ' and a new row is started.
                'Debug.Print "Possible cost code: "; strSplit(11 - 1)   '@@@used for testing
                If (Len(strSplit(11 - 1))) <= 8 Then
                    'Add a new row for the document record
                    lngRow = UBound(mstrReport, 2) + 1 'get the index of the next new row
                    ReDim Preserve mstrReport(REPTARRAYCOLS, lngRow)
                    
                    'Extract cost code and cost type from the existing record
                    'Cost Code - field 11
                    'Cost Type - field 12
                    mstrReport(ICOSTCODE, lngRow) = strSplit(11 - 1)
                    mstrReport(ICOSTTYPE, lngRow) = strSplit(12 - 1)
                    
                    blnInCostRec = True     'set state
                    RaiseEvent LogMsg("(" & MODULE & ".ParseBillingReport) New record found, report line " & TextStreamIn.Line)
                    
                End If
                
            Case 0  'doc row 2
                If blnInCostRec = True Then
                    'If doctype is one of the desired values then...
                    Select Case strSplit(13 - 1)
                        Case "P2", "PV", "PD", "PM"
                            'Keep this record
                            'Extract:
                            'Cat Code - Field 1
                            'G/L Date - field 15
                            'Amount - field 18
                            'Doctype - field 13
                            mstrReport(ICATCODE, lngRow) = strSplit(1 - 1)
                            mstrReport(IGLDATE, lngRow) = strSplit(15 - 1)
                            mstrReport(IGLAMT, lngRow) = strSplit(18 - 1)
                            mstrReport(IDOCTYPE, lngRow) = strSplit(13 - 1)
                            
                        Case Else
                            'Ignore this record by removing the last row in the
                            ' report array and clearing the state variable
                            ReDim Preserve mstrReport(REPTARRAYCOLS, lngRow - 1)
                            blnInCostRec = False
                            RaiseEvent LogMsg("(" & MODULE & ".ParseBillingReport) Ignoring record because Doctype = '" & _
                                strSplit(13 - 1) & "'")
                    End Select
                    
                Else
                    'Look for jobnum here
                    If strSplit(0) = "Job:" Then
                        mstrJobNum = Trim(strSplit(1))
                        RaiseEvent LogMsg("(" & MODULE & ".ParseBillingReport) Job number found")
                    End If
                End If
                            
            Case 1  'doc row 3
                'If in a valid record then...
                If blnInCostRec = True Then
                    'If there is no invoice number in the record then...
                    If UBound(strSplit) < (8 - 1) Then
                    'If strSplit(8 - 1) = vbNullString Then
                        'Discard this row of the extracted data array
                        lngRow = lngRow - 1 'decrement the number of rows
                        ReDim Preserve mstrReport(REPTARRAYCOLS, lngRow)
                        blnInCostRec = False
                    
                    Else
                        'Extract:
                        'Cost Code description - field 2
                        'Expense Description - field 5
                        'Invoice Num - field 8
                        'Supplier Number - field 10
                        mstrReport(ICCDESC, lngRow) = strSplit(2 - 1)
                        mstrReport(IEXPDESC, lngRow) = FixName(strSplit(5 - 1))
                        mstrReport(IINVNUM, lngRow) = strSplit(8 - 1)
                        mstrReport(ISUPNUM, lngRow) = strSplit(10 - 1)
                            
                        'Now that we know the record is valid...
                        'Create new nodes whenever the cost code or cost type change
                        'If the cost code has change then...
                        If mstrReport(ICOSTCODE, lngRow) <> strLastCostCode Then
                            
                            'Create a new cost code node (child of top node)
                            strCostCodeKey = "*" & CStr(lngRow) & mstrReport(ICOSTCODE, lngRow)
                            Set objNode = objNodes.Add(KEY_TOPNODE, tvwChild, _
                                strCostCodeKey, _
                                mstrReport(ICOSTCODE, lngRow) & ", " & mstrReport(ICCDESC, lngRow))
                            objNode.EnsureVisible
                            strLastCostCode = mstrReport(ICOSTCODE, lngRow)
                            
                            'Create a new cost type node (child of cost code node)
                            strCostTypeKey = "#" & CStr(lngRow) & mstrReport(ICOSTTYPE, lngRow)
                            Set objNode = objNodes.Add(strCostCodeKey, tvwChild, _
                                strCostTypeKey, _
                                mstrReport(ICOSTTYPE, lngRow))
                            objNode.EnsureVisible
                            lntLastCostType = mstrReport(ICOSTTYPE, lngRow)
                            
                        'else if the cost type only has changed then...
                        ElseIf mstrReport(ICOSTTYPE, lngRow) <> lntLastCostType Then
                            'Create a new cost type node (child of cost code node)
                            strCostTypeKey = "#" & CStr(lngRow) & mstrReport(ICOSTTYPE, lngRow)
                            Set objNode = objNodes.Add(strCostCodeKey, tvwChild, _
                                strCostTypeKey, _
                                mstrReport(ICOSTTYPE, lngRow))
                            objNode.EnsureVisible
                            lntLastCostType = mstrReport(ICOSTTYPE, lngRow)
                        End If
                        
                        'Create a new invoice entry in tree view (child of cost type node)
                        Set objNode = objNodes.Add(strCostTypeKey, tvwChild, _
                            , _
                            mstrReport(IEXPDESC, lngRow) & ", " & mstrReport(IINVNUM, lngRow) & _
                                ", " & mstrReport(ISUPNUM, lngRow) & ", $" & mstrReport(IGLAMT, lngRow))
                        objNode.EnsureVisible
                        
                        'Indicate the record has been parsed
                        blnInCostRec = False
                        RaiseEvent LogMsg("(" & MODULE & ".ParseBillingReport) " & mstrReport(IDOCTYPE, lngRow) & " Invoice '" & _
                            mstrReport(IINVNUM, lngRow) & "' added to treeview")
                    End If
                
                End If  'blnInCostRec
                
            Case Else
                'Ignore other rows
        End Select
        
    Loop
    
    If mstrJobNum = vbNullString Then
        Err.Raise 1, "ParseBillingReport", "Job Number Not Found"
    End If
    
    RaiseEvent LogMsg("(" & MODULE & ".ParseBillingReport) Parsing of report file '" _
        & strFile & "' complete. " & UBound(mstrReport, 2) & " records extracted.")
    
'Restore mousepointer
    RaiseEvent MousePtr(vbDefault)

Exit Function

errHandler:
    Select Case Err.Number
        
        Case 1  'No job number
            MsgBox "The report selected is not in the expected format. The job " & _
                "number was not found indicating it might be an older version " & _
                "of the Job Cost report. " & vbCrLf & _
                "Make sure you select the Job Cost report in .csv format.", _
                vbExclamation + vbOKOnly, "Parse Billing Report"
            RaiseEvent MousePtr(vbDefault)
            'exit function
            
'        Case 35602  'Key is not unique in collection
'            MsgBox "A duplicate record was found in the report"
'            'strInvoiceKey = strInvoiceKey & objNodes.Count
'            Resume
        Case 35601  'element not found
            MsgBox "The report file contains a record that is out of sequence." & _
                vbCrLf & "The record cannot be displayed.", , "Parse Billing Report"
            Resume Next
        Case 9  'subscript out of range
            MsgBox "The report selected is not in the expected format and it cannot be " & _
                "processed. Make sure you select a Job Cost report in .csv format.", _
                vbExclamation + vbOKOnly, "Parse Billing Report"
                RaiseEvent MousePtr(vbDefault)
                'exit function
        Case Else
            'Call RaiseError(errParseBillingReport, MODULE & ".ParseBillingReport", Err.Number & "-" & Err.Description)
            MsgBox "An error was encountered while trying to read the report and it cannot be " & _
                "processed. Make sure you select a Job Cost report in .csv format.", _
                vbExclamation + vbOKOnly, "Parse Billing Report"
                RaiseEvent MousePtr(vbDefault)
            strErr = Err.Number
            strDesc = Err.Description
            Err.Clear
            RaiseEvent LogMsg("(" & MODULE & ".ParseBillingReport) Error: '" & _
                strErr & "' Description: '" & strDesc & _
                "' encountered during parsing")
            'exit function
    End Select
'Resume  '@@@
End Function

Private Function FindAndExportInvoice(ByRef udtInvoiceDat As udtInvoice) As String
'Query for the specified invoice in the application designated
'by the user and load it into the viewer.
'Return the patch of the tem folder

Dim obj As Object
Dim key As Variant
Dim SC As String
Dim i As Integer
Dim strInvoice As String
Dim strSupplierNum As String
'Requires a project reference to OTObjID.dll
'Dim objObjID As OTOBJIDLib.ObjectID

'Requires a project reference to ADO
Dim objQResults As ADODB.Recordset
Dim datStart As Date
Dim lngTime As Long

'Requires a project reference to OTDocID.dll
Dim objDocId As OTACORDELib.DocumentID
Dim objIndexid As OTACORDELib.IndexID
'Requires a project reference to OTExportManager.dll
Dim objExpDoc As OTCONTEXTLib.ExportedDocument
Dim objExpEnv As OTCONTEXTLib.ExportEnvironment
Dim objPageRange As OTCONTEXTLib.PageRanges
Dim objSection As OTCONTEXTLib.ExportedSection
Dim objPages As OTCONTEXTLib.ExportedPages
Dim objSections As OTCONTEXTLib.ExportedSections
Dim objPage As OTCONTEXTLib.ExportedPage
Dim intExportType As Integer
Dim strExportPath As String
'Requires project reference to Microsoft Scripting Runtime (scrrun.dll)
Dim fs As FileSystemObject
Dim TextStreamIn As TextStream
Dim fsoFile As File
Dim strExt As String    'file extension
Dim strFinalName As String      'final name of exported document
Dim strExistingFile As String   'name of document already exported
Dim lngPageNum As Long          'page number
Dim lngSecNum As Long           'section number
Dim datGLDate  As Date          'converted GL date

'Enable error trap
    On Error GoTo errHandler

'Query for the invoice document
    Me.MousePointer = vbHourglass
    RaiseEvent LogMsg("(" & MODULE & ".FindAndExportInvoice) *****Querying for invoice: " & udtInvoiceDat.InvoiceNum)
    datStart = Now
    Set objQResults = ExecuteAdHocQuery(udtInvoiceDat.InvoiceNum, udtInvoiceDat.SupplierNum, udtInvoiceDat.Vendor)
    lngTime = DateDiff("s", datStart, Now)
    RaiseEvent LogMsg("(" & MODULE & ".FindAndExportInvoice) Query processing completed in " & lngTime & " secs")
    mfrmExp.lblQuery = lngTime
    mfrmExp.Refresh
    Me.MousePointer = vbDefault

'If there were no records then...
    If objQResults.EOF = True And objQResults.BOF = True Then
        udtInvoiceDat.InAcorde = False
        udtInvoiceDat.ExportPath = "not exported"
        RaiseEvent LogMsg("(" & MODULE & ".FindAndExportInvoice) Invoice not found in Acorde")

    'Export the image in the specified format
    Else
        udtInvoiceDat.InAcorde = True
        RaiseEvent LogMsg("(" & MODULE & ".FindAndExportInvoice) Invoice found in Acorde")
        Set fs = New FileSystemObject
        'Get the export location
        strExportPath = dirExport.Path
            
        'For each document found in the query results...
        Do While objQResults.EOF = False
        
            'Create the desired name for the file to be exported (without extension)
            'Note - LUCID is used in the file name in case there are multiple
            'documents that matched the search parameters. LUCID will make the name
            'of each unique
'           old format
'            strFinalName = udtInvoiceDat.CostCode & ", " & udtInvoiceDat.CostType & _
'                ", " & udtInvoiceDat.InvoiceNum & "." & objQResults("LUCID")
'           new format
            'Cat code, cost code, cost type, date, Invoice num, #unique id
            datGLDate = CDate(udtInvoiceDat.GLDate)
            strFinalName = udtInvoiceDat.CatCode & ", " & udtInvoiceDat.CostCode & ", " & udtInvoiceDat.CostType & _
                ", " & Format(datGLDate, "m-d-yyyy") & ", " & udtInvoiceDat.InvoiceNum & ", #" & objQResults("LUCID")
            
            'Replace any illegal characters in the file name
            strFinalName = Replace(strFinalName, "/", "-")
            strFinalName = Replace(strFinalName, "\", "-")
            strFinalName = Replace(strFinalName, ":", ";")
            strFinalName = Replace(strFinalName, "*", " ")
            strFinalName = Replace(strFinalName, "?", " ")
            strFinalName = Replace(strFinalName, "<", "(")
            strFinalName = Replace(strFinalName, ">", ")")
            strFinalName = Replace(strFinalName, "|", "-")
            strFinalName = Replace(strFinalName, """", "'")
            
            'Get the output file format from the form controls
            For i = 0 To optExportType.UBound   'Ubound = 6 currently
                If optExportType(i).Value = True Then
                    intExportType = optExportType(i).Tag
                    RaiseEvent LogMsg("(" & MODULE & ".FindAndExportInvoice) Document being exported in " & optExportType(i).Caption & " format")
                    If intExportType = 0 Or intExportType > otiExportSmartMode Then
                        Err.Raise 1, , "optExportType(" & i & ") has an invalid Tag property value. Set it to the appropriate ExportFileType Enum"
                    End If
                End If
            Next i
            
            'Get the export type extension for a previously exported file
            Select Case intExportType
                Case otiExportBMP
                    strExt = "*.bmp"
                Case otiExportEMF
                    strExt = "*.emf"
                Case otiExportJPG
                    strExt = "*.jpg"
                Case otiExportPCX
                    strExt = "*.pcx"
                Case otiExportPDF
                    strExt = ".pdf"
                Case otiExportTIF
                    'If multipage TIFF then...
                    If optExportType(0).Value = True Then
                        strExt = ".tif"
                    Else
                        strExt = "*.tif"
                    End If
            End Select
                    
            'If the file has not previously been exported then...
'            If fs.FileExists(strExportPath & "\" & strFinalName) = False Then
            strExistingFile = Dir(strExportPath & "\" & strFinalName & strExt)
            If strExistingFile = vbNullString Then
                'Export it
                'Init the export environment
                Set objExpEnv = New OTCONTEXTLib.ExportEnvironment
                objExpEnv.UseHighestResolution = True
                objExpEnv.MaxOutputPages = 1000
                objExpEnv.ExportType = intExportType
                objExpEnv.ExportBasePath = strExportPath
                objExpEnv.ExportNameSpace = EXPORTFOLDER
                
                'Init the DocId object for the specified document
                Set objDocId = New OTACORDELib.DocumentID
                objDocId.ProviderID = "{0BF3C340-4C13-11d3-8166-00C04F99E979}"
                objDocId.UniqueID = objQResults("LUCID")    'objQResults.Fields(0)  'RECID
                objDocId.MimeType = "image/tiff"
                'Note - this code only exports TIFFs from Acorde
                'Exporting Universals would require extra code.
                'Also objExpDoc.RecombineDocument does not work for all MIME types
        
                'Init the IndexId object for the specified document
                Set objIndexid = New OTACORDELib.IndexID
                objIndexid.IndexName = cboApps.Text
                objIndexid.IndexProvider = "{608FCB70-10BF-11d4-A931-00C04F94786A}"
                objIndexid.IndexID = objQResults("LUCID")   'objQResults.Fields(0)  'RECID
                
                'Init the exported document object
                Set objExpDoc = New OTCONTEXTLib.ExportedDocument
                objExpDoc.ExportEnvironment = objExpEnv
                objExpDoc.UserToken = mobjUserToken
                objExpDoc.DocumentID = objDocId
                objExpDoc.IndexID = objIndexid
                
                'Export the document
                'If export type is multipage tiff then...
                If (optExportType(0).Value = True) Or (intExportType = otiExportPDF) Then
                    Set objPageRange = New PageRanges
                    objPageRange.AddRange 1, objQResults("PAGECOUNT")    'must be actual page count of document
                    datStart = Now
                    Set objSection = objExpDoc.RecombineDocument(objPageRange)
                    lngTime = DateDiff("s", datStart, Now)
                    RaiseEvent LogMsg("(" & MODULE & ".FindAndExportInvoice) Export completed in " & lngTime & " secs")
                    mfrmExp.lblExport = lngTime
                    
                    RaiseEvent LogMsg("(" & MODULE & ".FindAndExportInvoice) Document (all pages) " & _
                        "exported to " & objSection.FullPath)
                    Debug.Print objSection.RelativePath 'use up to first backslash for temp path
                    'Debug.Print objSection.FullPath
                    'Rename of the exported document and move it
                    Debug.Print objSection.FullPath
                    Set fsoFile = fs.GetFile(objSection.FullPath)
                    FindAndExportInvoice = fsoFile.ParentFolder 'return the temp Acorde export path
                    strExt = Right(fsoFile.Name, (Len(fsoFile.Name) - InStrRev(fsoFile.Name, ".")))
                    fsoFile.Name = strFinalName & "." & strExt
                    fsoFile.Move dirExport.Path & "\" & fsoFile.Name
                    RaiseEvent LogMsg("(" & MODULE & ".FindAndExportInvoice) Document " & _
                      "moved to " & dirExport.Path & "\" & fsoFile.Name)
                
                Else    'non-multipage export format
                    lngPageNum = 0
                    'Get the pages collection of the export doc object
                    Set objPages = objExpDoc.Pages
                    'For each page of the document
                    For Each objPage In objPages
                        lngPageNum = lngPageNum + 1
                        'Export the page in the specified format
                        datStart = Now
                        Set objSections = objPage.Sections
                        lngTime = DateDiff("s", datStart, Now)
                        RaiseEvent LogMsg("(" & MODULE & ".FindAndExportInvoice) Export completed in " & lngTime & " secs")
                        mfrmExp.lblExport = lngTime
                        mfrmExp.Refresh
                      
                        lngSecNum = 0
                        For Each objSection In objSections
                            lngSecNum = lngSecNum + 1
                            RaiseEvent LogMsg("(" & MODULE & ".FindAndExportInvoice) Document page " & _
                              lngPageNum & " exported to " & objSection.FullPath)
                            'Rename of the exported document and move it
                            Debug.Print objSection.FullPath
                            Set fsoFile = fs.GetFile(objSection.FullPath)
                            FindAndExportInvoice = fsoFile.ParentFolder  'return the temp Acorde export path
                            strExt = Right(fsoFile.Name, (Len(fsoFile.Name) - InStrRev(fsoFile.Name, ".")))
                            fsoFile.Name = strFinalName & "." & lngPageNum & "." & lngSecNum & "." & strExt
                            fsoFile.Move dirExport.Path & "\" & fsoFile.Name
                            RaiseEvent LogMsg("(" & MODULE & ".FindAndExportInvoice) Document page " & _
                              lngPageNum & " moved to " & dirExport.Path & "\" & fsoFile.Name)
                            
                        Next objSection
                    Next objPage
                End If
                 
                'Return the final path in udtInvoiceDat
'                udtInvoiceDat.ExportPath = fsoFile.Path
                udtInvoiceDat.ExportPath = fsoFile.Name
                
                'Clean up
                Set objDocId = Nothing
                Set objIndexid = Nothing
                Set objExpDoc = Nothing
                Set objSection = Nothing
                Set objPageRange = Nothing
                Set objExpEnv = Nothing
                Set fsoFile = Nothing
            
            Else 'export file already exists
'                udtInvoiceDat.ExportPath = strExportPath & "\" & strExistingFile
                udtInvoiceDat.ExportPath = strExistingFile
                RaiseEvent LogMsg("(" & MODULE & ".FindAndExportInvoice) Document '" & _
                  strExportPath & "\" & strExistingFile & "' previously exported")
                    
            End If
            
            objQResults.MoveNext
        Loop
        
        Set fs = Nothing
    End If
    

Exit Function

errHandler:
    Select Case Err.Number
        Case -2147188504    'Cannot create export engine
            MsgBox "The Adobe Acrobat software was not detected on " & _
                "Acorde Export server. This must be obtained from Adobe. " & _
                "Please select another export format.", vbExclamation + vbOKOnly, "Export"
            'Disable the PDF option
            optExportType(6).Enabled = False
            optExportType(0).Value = True
            'Raise an error that the calling procedure can detect
            Err.Raise errNoPDF
        Case Else
            Call RaiseError(errFindAndExportInvoice, MODULE & ".FindAndExportInvoice", Err.Number & "-" & Err.Description)
    End Select
'Resume  '@@@
End Function

Private Sub Form_Unload(Cancel As Integer)
'Log out if previously logged-in to Acorde.

    If Not (mobjUserToken Is Nothing) Then
        'Release the license
        mobjUser.Logout mobjUserToken
    End If

End Sub

Private Sub mfrmExp_Cancel()
'This event fires when users whats to cancel the export operation.
'Set the Cancel boolean

    mblnCancelExport = True
    RaiseEvent LogMsg("(" & MODULE & ".mfrmExp_Cancel) User cancelled export operation")

End Sub

Private Sub optAmt_Click()

'Enable the Select Filtered button now that an option is selected
    cmdSelFilter.Enabled = True
    
End Sub

Private Sub optCostType_Click()

'Enable the Select Filtered button now that an option is selected
    cmdSelFilter.Enabled = True
    
End Sub

Private Function ExecuteAdHocQuery(strInvoice As String, strSupplierNumber As String, strVendor As String) _
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
Private Function LeadingCommas(strText As String) As Long
'Return the number of consecutive leading commas in the input string.

Dim lngCount As Long
Dim strChar As String

'Enable error trap
    On Error GoTo errHandler

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
            Call RaiseError(errGeneric, MODULE & ".LeadingCommas", Err.Number & _
                "-" & Err.Description)
    End Select
End Function

Private Function TrailingCommas(strText As String) As Long
'Return the number of consecutive trailing commas in the input string.

Dim lngCount As Long
Dim strChar As String
Dim lngLoop As Long

'Enable error trap
    On Error GoTo errHandler

'Initialize count to indicate to commas found
    lngCount = 0
    
'Initialize character count
    lngLoop = 0
    
'Loop to test each character in string beginning at last
    Do
        'Get the next character
        strChar = Mid(strText, Len(strText) - lngLoop, 1)
        'If character is a string then...
        If strChar = "," Then
            'Count it
            lngCount = lngCount + 1
        End If
        lngLoop = lngLoop + 1
    Loop While strChar = ","
    
'Return the result
    TrailingCommas = lngCount
    
Exit Function

errHandler:
    Select Case Err.Number
        
        Case Else
            Call RaiseError(errGeneric, MODULE & ".TrailingCommas", Err.Number & _
                "-" & Err.Description)
    End Select
End Function

Private Sub SelectByAmount(curAmount As Currency)
'Select nodes with an amount >= the specified dollar amount.

Dim i As Integer
Dim strText As String
Dim strAmount As String
Dim intComma2 As Integer

'Enable error trap
    On Error GoTo errHandler

'Check each node for a dollar amount.
'Check the node if it meets the dollar amount threshold.

    For i = 1 To trvInvoices.Nodes.Count
        'CheckNode trvInvoices.Nodes(i), False  'uncheck if previously checked
        'If this is an invoice node then...
        If trvInvoices.Nodes(i).key = vbNullString Then
            'Check the dollar amount
            strText = trvInvoices.Nodes(i).Text
            intComma2 = InStrRev(strText, ",")
            strAmount = Right(strText, Len(strText) - intComma2)
            If IsNumeric(strAmount) = True Then
                If CCur(strAmount) >= curAmount Then
                    CheckNode trvInvoices.Nodes(i), True
                End If
            End If
        End If
    Next i

Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".SelectByAmount", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub SelectByCostType(lngCostType As Long)
'Select nodes with a cost type = the specified value.

Dim i As Integer
Dim strCostType As String
Dim j As Long
Dim nodChild As MSComctlLib.Node

'Enable error trap
    On Error GoTo errHandler

'Check each node for a dollar amount.
'Check the node if it meets the dollar amount threshold.

    For i = 1 To trvInvoices.Nodes.Count
        'CheckNode trvInvoices.Nodes(i), False  'uncheck if prevoiusly checked
        'If this is an cost type node then...
        If Left(trvInvoices.Nodes(i).key, 1) = "#" Then
            'Get the cost type
            strCostType = trvInvoices.Nodes(i).Text
            'If cost type matches specified value then...
            If CLng(strCostType) = lngCostType Then
                CheckNode trvInvoices.Nodes(i), True
                'Check all of its children
                'Check children nodes
                Set nodChild = trvInvoices.Nodes(i).Child
                For j = 1 To trvInvoices.Nodes(i).Children
                    CheckNode nodChild, True
                    Set nodChild = nodChild.Next
                Next j
            End If
        End If
    Next i

Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".SelectByCostType", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub trvInvoices_NodeCheck(ByVal Node As MSComctlLib.Node)
'Change formatting of node depending on if checked or not.
'Also check or uncheck any children.

Dim i As Integer
Dim j As Integer
Dim nodChild As Node
Dim nodGrandchild As Node

'Enable error trap
    On Error GoTo errHandler

    If Node.Checked = True Then
        CheckNode Node, True
        'Check children nodes
        Set nodChild = Node.Child
'        Debug.Print nodChild.FirstSibling.Text
'        Debug.Print nodChild.LastSibling.Text
        For i = 1 To Node.Children
            CheckNode nodChild, True
            
            'Check grandchild nodes
            Set nodGrandchild = nodChild.Child
            For j = 1 To nodChild.Children
                CheckNode nodGrandchild, True
                Set nodGrandchild = nodGrandchild.Next
            Next j
            
            Set nodChild = nodChild.Next
        Next i
    Else
        CheckNode Node, False
        'Check children nodes
        Set nodChild = Node.Child
        For i = 1 To Node.Children
            CheckNode nodChild, False
            
            'Check grandchild nodes
            Set nodGrandchild = nodChild.Child
            For j = 1 To nodChild.Children
                CheckNode nodGrandchild, False
                Set nodGrandchild = nodGrandchild.Next
            Next j
            
            Set nodChild = nodChild.Next
        Next i
    End If

Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".trvInvoices_NodeCheck", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Sub CheckNode(ByVal Node As MSComctlLib.Node, CheckIt As Boolean)
'Set the node formatting according to the input boolean

'Enable error trap
    On Error GoTo errHandler

    'If it is to be checked then...
    If CheckIt = True Then
        Node.Checked = True
        Node.ForeColor = QBColor(1) 'blue
        Node.Bold = True
    
    Else
        Node.Checked = False
        Node.ForeColor = QBColor(0) 'black
        Node.Bold = False
    End If

Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".CheckNode", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Sub

Private Function GenerateReports(strReportFile As String, blnCancelled As Boolean) As Long
'Generate a CSV file that lists the documents that were exported and the exceptions.
'Return the count of documents exported.

'Requires project reference to Microsoft Scripting Runtime (scrrun.dll)
Dim objFileSys As FileSystemObject
Dim tsOut As TextStream
Dim strFileName As String
Dim i As Long

'Enable error trap
    On Error GoTo errHandler

'Create the report file for writing
    Set objFileSys = New FileSystemObject
    Set tsOut = objFileSys.OpenTextFile(strReportFile, ForWriting, True)
    
'Write the Job number
    tsOut.WriteLine "Export Results for Job Number " & mstrJobNum

'If the export was cancelled then...
    If blnCancelled = True Then
        'Write the records for the invoices not queried
        tsOut.WriteLine ""
        tsOut.WriteLine "The Export Operation was cancelled before completion."
        tsOut.WriteLine "Invoices not searched for in Acorde:"
        tsOut.WriteLine Pad("Supplier", 32) & vbTab & _
            Pad("Invoice Number", 25) & vbTab & _
            Pad("Date", 10) & vbTab & _
            Pad("Amount", 9) & vbTab & _
            "Exported File"
        'For each record in the report array...
        For i = 1 To UBound(mstrReport, 2)
            'If the invoice was to be exported but there is not status of it being in Acorde then...
            If (mstrReport(IINACORDE, i) = vbNullString) And (mstrReport(IEXPORT, i) = CStr(True)) Then
                tsOut.WriteLine Pad(mstrReport(IEXPDESC, i), 32) & vbTab & _
                    Pad(mstrReport(IINVNUM, i), 25) & vbTab & _
                    Pad(mstrReport(IGLDATE, i), 10) & vbTab & _
                    Pad(Format(mstrReport(IGLAMT, i), "####0.00"), 9) & vbTab & _
                    mstrReport(IEXPPATH, i)
            End If
        Next i
    End If
    
'Write the records for the invoices NOT found in Acorde
    tsOut.WriteLine ""
    tsOut.WriteLine "Invoices NOT found in Acorde:"
    tsOut.WriteLine Pad("Supplier", 32) & vbTab & _
        Pad("Invoice Number", 25) & vbTab & _
        Pad("Date", 10) & vbTab & _
        Pad("Amount", 9) & vbTab & _
        "Exported File"
    'For each record in the report array...
    For i = 1 To UBound(mstrReport, 2)
        If (mstrReport(IINACORDE, i) = CStr(False)) And (mstrReport(IEXPORT, i) = CStr(True)) Then
            tsOut.WriteLine Pad(mstrReport(IEXPDESC, i), 32) & vbTab & _
                Pad(mstrReport(IINVNUM, i), 25) & vbTab & _
                Pad(mstrReport(IGLDATE, i), 10) & vbTab & _
                Pad(Format(mstrReport(IGLAMT, i), "####0.00"), 9) & vbTab & _
                mstrReport(IEXPPATH, i)
        End If
    Next i

'Write the records for the invoices found in Acorde
    'For each record in the report array...
    tsOut.WriteLine ""
    tsOut.WriteLine "Invoices found in Acorde:"
    tsOut.WriteLine Pad("Supplier", 32) & vbTab & _
        Pad("Invoice Number", 25) & vbTab & _
        Pad("Date", 10) & vbTab & _
        Pad("Amount", 9) & vbTab & _
        "Exported File"
    For i = 1 To UBound(mstrReport, 2)
        If (mstrReport(IINACORDE, i) = CStr(True)) And (mstrReport(IEXPORT, i) = CStr(True)) Then
            tsOut.WriteLine Pad(mstrReport(IEXPDESC, i), 32) & vbTab & _
                Pad(mstrReport(IINVNUM, i), 25) & vbTab & _
                Pad(mstrReport(IGLDATE, i), 10) & vbTab & _
                Pad(Format(mstrReport(IGLAMT, i), "#####0.00"), 9) & vbTab & _
                mstrReport(IEXPPATH, i)
        End If
    Next i

'Clear the InAcorde status for the next export run so that the
'next report will not contain leftover status
    For i = 1 To UBound(mstrReport, 2)
            mstrReport(IINACORDE, i) = vbNullString
    Next i

'Close the file
    tsOut.Close

'Clean up
    Set tsOut = Nothing
    Set objFileSys = Nothing
    
    RaiseEvent LogMsg("(" & MODULE & ".GenerateReports) Report file generated: " & strReportFile)
Exit Function

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".GenerateReports", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Function

Private Function GenerateIndex(strIndexfile As String)
'Generate an HTML index file that lists the documents that were exported.
'Format the HTML to look like the treeview.

'Requires project reference to Microsoft Scripting Runtime (scrrun.dll)
Dim objFileSys As FileSystemObject
Dim tsOut As TextStream
Dim strFileName As String
Dim i As Long
Dim strLastCostCode As String
Dim strLastCostType As String
Dim strHyperlink As String

'Enable error trap
    On Error GoTo errHandler

'Create the index file for writing
    Set objFileSys = New FileSystemObject
    Set tsOut = objFileSys.OpenTextFile(strIndexfile, ForWriting, True)
    
'Write the HTML header and start the body
    tsOut.WriteLine "<html>"
    tsOut.WriteLine ""
    tsOut.WriteLine "<head>"
    tsOut.WriteLine "<meta http-equiv=""Content-Language"" content=""en-us"">"
    tsOut.WriteLine "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1252"">"
    tsOut.WriteLine "<meta name=""GENERATOR"" content=""Microsoft FrontPage 4.0"">"
    tsOut.WriteLine "<meta name=""ProgId"" content=""FrontPage.Editor.Document"">"
    tsOut.WriteLine "<title>Exported Document Index</title>"
    tsOut.WriteLine "</head>"
    tsOut.WriteLine ""
    tsOut.WriteLine "<body>"
    tsOut.WriteLine ""
    tsOut.WriteLine "<p>Invoice Image Index</p>"

'Write the index content
    strLastCostCode = ""
    strLastCostType = ""
    For i = 1 To UBound(mstrReport, 2)
        If mstrReport(IEXPORT, i) = CStr(True) Then
            If mstrReport(IINACORDE, i) = CStr(True) Then
                If mstrReport(ICOSTCODE, i) <> strLastCostCode Then
                    'Write the cost code
                    tsOut.WriteLine "<p>" & mstrReport(ICOSTCODE, i) & ", " & mstrReport(ICCDESC, i) & "</p>"
                    'Write the cost type
                    tsOut.WriteLine "<p>&nbsp;&nbsp;&nbsp; " & mstrReport(ICOSTTYPE, i) & "</p>"
                    'Save the cost code and cost type
                    strLastCostCode = mstrReport(ICOSTCODE, i)
                    strLastCostType = mstrReport(ICOSTTYPE, i)
                ElseIf mstrReport(ICOSTTYPE, i) <> strLastCostType Then
                    'Write the cost type
                    tsOut.WriteLine "<p>&nbsp;&nbsp;&nbsp; cost type1</p>"
                    'Save the cost type
                    strLastCostType = mstrReport(ICOSTTYPE, i)
                End If
                
                'Write the invoice
                tsOut.WriteLine "<p style=""margin-top: 0; margin-bottom: 0"">&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;"
                Debug.Print Hyperlink(mstrReport(IEXPPATH, i))
                strHyperlink = "<a href=""" & Hyperlink(mstrReport(IEXPPATH, i)) & """ target=""_blank"">" & _
                    mstrReport(IEXPDESC, i) & " " & mstrReport(IINVNUM, i) & " " & Format(mstrReport(IGLAMT, i), "$######0.00") & "</a></p>"
                tsOut.WriteLine strHyperlink
            End If
        End If
    Next i

'Close the body and HTML
    tsOut.WriteLine "</body>"
    tsOut.WriteLine ""
    tsOut.WriteLine "</html>"

'Close the file
    tsOut.Close

'Clean up
    Set tsOut = Nothing
    Set objFileSys = Nothing
    
    RaiseEvent LogMsg("(" & MODULE & ".GenerateIndex) Index file generated: " & strIndexfile)
Exit Function

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".GenerateIndex", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Function

Private Function Hyperlink(ByVal strFileSpec As String) As String
'Reformat a filespec for for an HTML hyperlink.
'Replace backslashes with forward slashes
'Replace spaces with %20

'Enable error trap
    On Error GoTo errHandler

    strFileSpec = Replace(strFileSpec, "\", "/")
    strFileSpec = Replace(strFileSpec, " ", "%20")
    Hyperlink = strFileSpec
    
Exit Function

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".Hyperlink", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Function

Private Function IsDuplicate(strComparisonValue As String, ValueIndex As Integer, _
    LastRow As Long) As Boolean
'Search the Report array up to the specified row to see if the specified value
'has already appeared in the array

Dim i As Long

    On Error GoTo errHandler
    
    IsDuplicate = False 'default value
    
    'For all array rows up to the last row (but not including it)
    For i = 1 To LastRow - 1
        'If the array value matches the specified value then...
        If mstrReport(ValueIndex, i) = strComparisonValue Then
            'It's a duplicate
            IsDuplicate = True
            Exit For
        End If
    Next i

Exit Function

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".IsDuplicate", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Function

Private Function Pad(strInput, intLen As Integer) As String
'Return the string padded to the specified number of characters

Dim strSpaces As String
Dim i As Integer

'Build the padding string
    For i = 1 To (intLen - Len(strInput))
        strSpaces = strSpaces & " "
    Next i
    
'Build the output string
    Pad = strInput & strSpaces
    
Exit Function

errHandler:
    Select Case Err.Number
        Case Else
            Call RaiseError(errGeneric, MODULE & ".Pad", Err.Number & _
                "-" & Err.Description)
    End Select
'Resume  '@@@
End Function

Private Function FixName(ByVal strName As String)
'Modify the Vendor name by removing quotes and replacing pipes with commas.

    strName = Replace(strName, "|", ",")    'replace pipes with commas
    strName = Replace(strName, """", "")    'eliminate quotes
    FixName = strName

End Function

Private Sub RaiseError(ErrorNumber As Long, Source As String, Description As String)
'Log and raise the error.

'Log the error
    RaiseEvent LogMsg("(" & MODULE & ".RaiseError) Error in " & Source & ": " & ErrorNumber & " - " & Description)
    
'Raise an error up to the client
    Err.Raise ErrorNumber, Source, Description
    
End Sub





