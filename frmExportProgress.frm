VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExportProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Progress"
   ClientHeight    =   5130
   ClientLeft      =   4035
   ClientTop       =   3105
   ClientWidth     =   7245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   5130
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel Export"
      Height          =   495
      Left            =   2880
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      Top             =   4200
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar prgExport 
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1800
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      Caption         =   "Processsing x of y"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Last Query (secs)"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Last Export (secs)"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblQuery 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblExport 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblRemaining 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblElapsed 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Minutes Remaining"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Minutes Elapsed"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmExportProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
' Class:        frmExportProgress
' Author:       Wes Prichard, Optika
' Date:         April 2003
' Description:  Displays export progress for the billing backup export functionality.
' Edit History:
' mm/dd/yyyy - name, company
'   Description

'Public Interface:
'Properties:
'None

'Methods:
'Display - makes form visible

'Events:
'None

'Dependencies:
'None
'*****************************************************************************************

Option Explicit

Private Const ERRORBASE = ErrorBase5    'see modErrorHandling - used to avoid
                                        '   overlapping error numbers
Private Const MODULE = "frmExportProgress"     'used in reporting errors

'Public event declarations
Public Event Cancel()  'indicates user wants to cancel

'Class Error enumeration
Public Enum errCBB
'    errNoPDF = ERRORBASE + 0
'    errParseBillingReport = ERRORBASE + 1
'    errLoad = ERRORBASE + 2
'    errGetInvoiceApps = ERRORBASE + 3
'    errFindAndExportInvoice = ERRORBASE + 4
'    errToolBarClick = ERRORBASE + 5
'    errExecuteAdHocQuery = ERRORBASE + 6
'    errLoadInvoice = ERRORBASE + 7
    errGeneric = ERRORBASE + 8
'    errDBRollbackTrans = ERRORBASE + 9
'    errDBRollbackTrans = ERRORBASE + 10
End Enum

Public Sub Display()
'Custom entry point for loading and displaying form so that modality and other
'properties can be controlled.
 
'Make me visible
    'RaiseEvent LogMsg("(" & MODULE & ".Display) Entering procedure")
    Me.Show vbModeless  'this triggers Form_Load procedure
    'Me.WindowState = vbMaximized
    'Me.ControlBox = False
'    Me.MaxButton = False
'    Me.MinButton = False
    Me.Refresh
    
End Sub

Private Sub cmdCancel_Click()
'Raise an event to indicate the user want's to cancel the export operation

Dim mbrResponse As VbMsgBoxResult
    
    mbrResponse = MsgBox("Cancel the Export operation?", vbQuestion + vbYesNo, "Cancel")
    If mbrResponse = vbYes Then
        cmdCancel.Enabled = False
        RaiseEvent Cancel
    End If
    
    
End Sub

