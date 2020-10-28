Attribute VB_Name = "modMain"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"390DBAFC016F"
' Name:         modMain
' Author:       Wes Prichard, Optika
' Date:         November 1999
' Description:  Provides the entry point for the application and defines
'global
'   constants.
' Edit History:

Option Explicit

'global user-defined data types
'Public Type WindowPos         'values for positioning and sizing a window
'    Top As Long     'top of window
'    Left As Long    'left edge of window
'    Height As Long  'height of window
'    Width As Long   'width of window
'End Type

Global Const KeyPath = "Software\Optika\CheckSigning"     'path of keys in registry




Sub Main()
'Entry point for application.
'Login then show the MDI form.

Dim clsMain As cMain

    Set clsMain = New cMain
    Call clsMain.Main
    Set clsMain = Nothing

End Sub

