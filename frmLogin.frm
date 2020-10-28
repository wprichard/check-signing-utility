VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.TextBox txtConfirm 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1080
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   3780
      TabIndex        =   7
      Tag             =   "Cancel"
      Top             =   1560
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   2160
      TabIndex        =   6
      Tag             =   "OK"
      Top             =   1560
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   600
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   2985
      TabIndex        =   3
      Top             =   135
      Width           =   2325
   End
   Begin VB.Image Image2 
      Height          =   1485
      Left            =   120
      Picture         =   "frmLogin.frx":23D2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblConfirm 
      Caption         =   "&Confirm Password"
      Height          =   360
      Left            =   1800
      TabIndex        =   0
      Tag             =   "&Password:"
      Top             =   975
      Width           =   1080
   End
   Begin VB.Label lblPassword 
      Caption         =   "&Password:"
      Height          =   255
      Left            =   1785
      TabIndex        =   1
      Tag             =   "&Password:"
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblUserName 
      Caption         =   "&User Name:"
      Height          =   255
      Left            =   1785
      TabIndex        =   2
      Tag             =   "&User Name:"
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"390DBB07028D"
'*****************************************************************************************
' Class:        frmLogin
' Author:       Wes Prichard, Optika
' Date:         November 1999
' Description:  Provides a login dialog.
' Edit History:

'Public Interface:
'Properties:
'UserName
'Password
'OK

'Methods:
'Display - makes form visible

'Events:
'none

'Dependencies:
'advapi32.dll
'*****************************************************************************************

Option Explicit

'Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

'Public OK As Boolean
Private mvarUserName As String
Private mvarPassword As String
Private mvarOK As Boolean

Public Property Get UserName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinLines
    UserName = mvarUserName
End Property

Public Property Get Password() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinLines
    Password = mvarPassword
End Property

Public Property Get OK() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinLines
    OK = mvarOK
End Property

Public Sub Display(Optional UserName As String, Optional ConfirmPassword As Boolean = False)
'Public method to display the dialog

Dim sBuffer As String
Dim lSize As Long

    Unload frmSplash
    txtUserName = UserName
    lblConfirm.Visible = ConfirmPassword
    txtConfirm.Visible = ConfirmPassword
    Me.Show vbModal
        
End Sub

Private Sub cmdCancel_Click()
    mvarOK = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    
    If txtConfirm.Visible = True Then
        'If the passwords do not match then...
        If txtPassword <> txtConfirm Then
            MsgBox "The passwords do not match, please reenter", vbExclamation + vbOKOnly, "OK"
        
        Else
            'Save the contents of the text boxes
            mvarUserName = txtUserName
            mvarPassword = txtPassword
        
            'Indicate the user clicked OK
            mvarOK = True
        
            'Hide the form
            Me.Hide
        End If
    
    Else    'not confirming password
        'Save the contents of the text boxes
        mvarUserName = txtUserName
        mvarPassword = txtPassword
    
        'Indicate the user clicked OK
        mvarOK = True
    
        'Hide the form
        Me.Hide
    
    End If
    
End Sub

Private Sub Form_Activate()
    txtPassword.SetFocus
End Sub

