VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*****************************************************************************************
' Class Name:   cRegistry
' Author:       Optika
' Date:         ??
' Description:  Handles storage and retrieval of data to/from the registry in an
'               application-defind key.
'
' Edit History:
' 06/08/2000 - Modified by W. Prichard, Optika
'   Modified error handling in all procedures to raise class-defined errors so that
'   those errors can be trapped by class clients.

' Dependencies:
' WIN3 API declaration contained in modRegistry.

' Design/Maintenance Notes:
'
'*****************************************************************************************

Option Explicit

Private Const ERRORBASE = ErrorBase6

'Class Error enumeration
Public Enum errCRegistry
    errRegistry = ERRORBASE + 0          '
End Enum

Public Enum RegSection
   regHKEY_LOCAL_MACHINE = &H80000002
   regHKEY_CLASSES_ROOT = &H80000000
   regHKEY_CURRENT_USER = &H80000001
End Enum

Public Function QueryValue(Section As RegSection, KeyPath As String, KeyName As String, _
                 sDefaultValue As String) As String
   Dim lRetVal As Long         'result of the API functions
   Dim hKey As Long         'handle of opened key
   Dim sValue As String      'setting of queried value
       
       lRetVal = RegOpenKeyEx(Section, KeyPath, 0, _
                 KEY_ALL_ACCESS, hKey)
       
       If lRetVal <> 0 Then GoTo err_exit
       
       lRetVal = QueryValueEx(hKey, KeyName, sValue)
       
       If lRetVal <> 0 Then GoTo err_exit
       
       QueryValue = sValue
       RegCloseKey (hKey)
       Exit Function
err_exit:
      QueryValue = sDefaultValue
End Function

Public Sub SaveValue(Section As RegSection, KeyPath As String, KeyName As String, Value As String)
   Dim lRetVal As Long
   Dim hKey As Long
   Dim sError As String
   
   lRetVal = RegOpenKeyEx(Section, KeyPath, 0, KEY_ALL_ACCESS, hKey)
  
   If lRetVal = ERROR_BADKEY Then  'if the key does not exist
       CreateNewKey KeyPath, Section
   ElseIf lRetVal <> 0 Then
       Select Case lRetVal
                Case 0
                    sError = "Zero Error"
                Case ERROR_BADDB
                        sError = "Bad DB"
                Case ERROR_CANTOPEN
                        sError = "Can not Open"
                Case ERROR_CANTREAD
                        sError = "Can not Read"
                
                Case ERROR_CANTWRITE
                        sError = "Can not Write"
                                        
                Case ERROR_OUTOFMEMORY
                        sError = "Out of memory"
                
                Case ERROR_ARENA_TRASHED
                        sError = "Arena Trashed"
                
                Case ERROR_ACCESS_DENIED
                        sError = "Access Denied"
                
                Case ERROR_INVALID_PARAMETERS
                        sError = "Invalid Parmaters"
                
                Case ERROR_NO_MORE_ITEMS
                        sError = "No more Items"
                        
                Case Else
                    sError = "Error " & lRetVal
       End Select
       Err.Raise errRegistry, "OptikaReg.SaveValue", "RegOpenKeyEx error: " & sError
   End If
   
   SetKeyValue Section, KeyPath, KeyName, Value, REG_SZ
   RegCloseKey (hKey)   'fix for 1450 error
   
End Sub

Private Sub CreateNewKey(sKeyPath As String, lSection As Long)
       Dim hNewKey As Long         'handle to the new key
       Dim lRetVal As Long         'result of the RegCreateKeyEx function
   Dim sError As String
       
       lRetVal = RegCreateKeyEx(lSection, sKeyPath, 0&, _
                 vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                 0&, hNewKey, lRetVal)
       
       If lRetVal <> 0 Then
           Select Case lRetVal
                Case ERROR_BADDB
                        sError = "Bad DB"
                Case ERROR_CANTOPEN
                        sError = "Can not Open"
                Case ERROR_CANTREAD
                        sError = "Can not Read"
                
                Case ERROR_CANTWRITE
                        sError = "Can not Write"
                                        
                Case ERROR_OUTOFMEMORY
                        sError = "Out of memory"
                
                Case ERROR_ARENA_TRASHED
                        sError = "Arena Trashed"
                
                Case ERROR_ACCESS_DENIED
                        sError = "Access Denied"
                
                Case ERROR_INVALID_PARAMETERS
                        sError = "Invalid Parmaters"
                
                Case ERROR_NO_MORE_ITEMS
                        sError = "No more Items"
                        
                Case Else
                    sError = "Error " & lRetVal
           End Select
           Err.Raise errRegistry, "OptikaReg.CreateNewKey", "Error on RegCreateKeyEx: " & sError
       End If
       
       RegCloseKey (hNewKey)
End Sub

Private Sub SetKeyValue(lSection As Long, sKeyPath As String, sKeyName As String, _
                        vValue As Variant, lValueType As Long)
       Dim lRetVal As Long         'result of the SetValueEx function
       Dim hKey As Long         'handle of open key
   Dim sError As String
       
       'open the specified key
       lRetVal = RegOpenKeyEx(lSection, sKeyPath, 0, KEY_ALL_ACCESS, hKey)
       If lRetVal <> 0 Then
           Select Case lRetVal
                Case ERROR_BADDB
                        sError = "Bad DB"
                Case ERROR_CANTOPEN
                        sError = "Can not Open"
                Case ERROR_CANTREAD
                        sError = "Can not Read"
                
                Case ERROR_CANTWRITE
                        sError = "Can not Write"
                                        
                Case ERROR_OUTOFMEMORY
                        sError = "Out of memory"
                
                Case ERROR_ARENA_TRASHED
                        sError = "Arena Trashed"
                
                Case ERROR_ACCESS_DENIED
                        sError = "Access Denied"
                
                Case ERROR_INVALID_PARAMETERS
                        sError = "Invalid Parmaters"
                
                Case ERROR_NO_MORE_ITEMS
                        sError = "No more Items"
           
                Case Else
                    sError = "Error " & lRetVal
           End Select
           Err.Raise errRegistry, "OptikaReg.SetKeyValue", "RegOpenKeyEx error: " & sError
       End If
       
       
       lRetVal = SetValue(hKey, sKeyName, lValueType, vValue)
       
       If lRetVal <> 0 Then
           Select Case lRetVal
                Case ERROR_BADDB
                        sError = "Bad DB"
                Case ERROR_CANTOPEN
                        sError = "Can not Open"
                Case ERROR_CANTREAD
                        sError = "Can not Read"
                
                Case ERROR_CANTWRITE
                        sError = "Can not Write"
                                        
                Case ERROR_OUTOFMEMORY
                        sError = "Out of memory"
                
                Case ERROR_ARENA_TRASHED
                        sError = "Arena Trashed"
                
                Case ERROR_ACCESS_DENIED
                        sError = "Access Denied"
                
                Case ERROR_INVALID_PARAMETERS
                        sError = "Invalid Parmaters"
                
                Case ERROR_NO_MORE_ITEMS
                        sError = "No more Items"
           
                Case Else
                    sError = "Error " & lRetVal
           End Select
           Err.Raise errRegistry, "OptikaReg.SetKeyValue", "SetValue error: " & sError
       End If
       
       RegCloseKey (hKey)
End Sub


Private Function SetValue(ByVal hKey As Long, sValueName As String, _
                           lType As Long, vValue As Variant) As Long
       Dim lValue As Long
       Dim sValue As String
       Select Case lType
           Case REG_SZ
               sValue = vValue & Chr$(0)
               SetValue = RegSetValueExString(hKey, sValueName, 0&, _
                                                 lType, sValue, Len(sValue))
           Case REG_DWORD
               lValue = vValue
               SetValue = RegSetValueExLong(hKey, sValueName, 0&, _
                            lType, lValue, 4)
       End Select
End Function


Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
   
   Dim cch As Long
   Dim lrc As Long
   Dim lType As Long
   Dim lValue As Long
   Dim sValue As String
   
   On Error GoTo QueryValueExError
   ' Determine the size and type of data to be read
     lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
   If lrc <> ERROR_NONE Then Error 5
      Select Case lType
           Case REG_SZ:
               sValue = String(cch, 0)
               lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
               If lrc = ERROR_NONE Then
                   vValue = Left$(sValue, cch - 1)
               Else
                   vValue = Empty
               End If           ' For DWORDS
           Case REG_DWORD:
               lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
               If lrc = ERROR_NONE Then vValue = lValue
           Case Else
               'all other data types not supported               lrc = -1
    End Select
QueryValueExExit:
         QueryValueEx = lrc
         Exit Function
QueryValueExError:
         Resume QueryValueExExit
End Function
