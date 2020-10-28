VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label W 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Weitz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2910
         TabIndex        =   3
         Top             =   480
         Width           =   1275
      End
      Begin VB.Image imgLogo 
         Height          =   465
         Left            =   1920
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   705
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   2400
         Width           =   6885
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2760
         TabIndex        =   2
         Top             =   1440
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"390DBE2D026E"
'*****************************************************************************************
' Class:        frmSplash
' Author:       Wes Prichard, Optika
' Date:         November 1999
' Description:  Implements the Splash screen.
' Edit History:

'Public Interface:
'Properties:
'none

'Methods:
'none

'Events:
'none

'Dependencies:
'ActiveX controls and references used by this form:
'Frame control
'Label control
'Image control
'*****************************************************************************************

Option Explicit

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    Unload Me
'End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

'Private Sub Frame1_Click()
'    Unload Me
'End Sub
