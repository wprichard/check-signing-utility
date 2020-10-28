VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"390DBE2D0372"
' Class:        cMain
' Author:       Wes Prichard, Optika
' Date:         November 1999
' Description:  Implements startup logic. (The class can receive the
'   Initialized event whereas a Basic module cannot.)
' Edit History:

'Public Interface:
'Properties:
'none

'Methods:
'Main

'Events:
'none

'Dependencies:
'none

Option Explicit


'global user-defined data types
'Public Type WindowPos         'values for positioning and sizing a window
'    Top As Long     'top of window
'    Left As Long    'left edge of window
'    Height As Long  'height of window
'    Width As Long   'width of window
'End Type




Public Sub Main()
'Entry point for application.
'Login then show the MDI form.

Dim fLogin As New frmLogin
Dim strUserID As String
Dim strPassword As String
    
''Load and display the login form
'    'fLogin.Show vbModal
'    fLogin.Display
'
''If login not successful then...
'    If Not fLogin.OK Then
'        'Login Failed so exit app
'        Unload fLogin
'        End
'    Else
'        strUserID = fLogin.UserName
'        strPassword = fLogin.Password
'    End If
'
''Unload the login form
'    Unload fLogin

    'Put up a splash screen
    frmSplash.Show vbModeless
    frmSplash.Refresh


'Load and display the main form
    Load frmMain  'begin initialization
    Unload frmSplash
    
    'make it visible
    frmMain.Display
    
    
'    'make it visible
'    fMainForm.Display strUserID, strPassword
    
End Sub


