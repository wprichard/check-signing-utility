Attribute VB_Name = "modErrorHandling"
'*****************************************************************************************
' Module Name:  modErrorHandling
' Author:       Wes Prichard, Optika
' Date:         08/28/2000
' Description:  Provides constants for application error handling
'
' Edit History:
' mm/dd/yyyy - Modified by name, company
'   Description of change
'
'*****************************************************************************************
Option Base 0
Option Explicit


' Define your custom error numbers here.  Be sure to use numbers
' greater than 512, to avoid conflicts with OLE error numbers.

'Define blocks of 50 error number constants for each project module
Public Const ErrorBase1 = vbObjectError + 512 + 50  'assigned to
Public Const ErrorBase2 = ErrorBase1 + 50           'assigned to none
Public Const ErrorBase3 = ErrorBase2 + 50           'assigned to
Public Const ErrorBase4 = ErrorBase3 + 50           'assigned to
Public Const ErrorBase5 = ErrorBase4 + 50           'assigned to none
Public Const ErrorBase6 = ErrorBase5 + 50           'assigned to cRegistry
Public Const ErrorBase7 = ErrorBase6 + 50           'assigned to frmCheckSign
Public Const ErrorBase8 = ErrorBase7 + 50           'assigned to frmBackupBilling


