VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'*****************************************************************************************
' Class Name:   cLog
' Author:       Jeff Anderson, Optika
' Date:         ??
' Description:  Provides methods to open, close, and add to a log file.
' Edit History:
' 06/03/1999 - Wes Prichard, Optika
'   Replaced public variable LoggingState with public property
'   LoggingEnabled.  Added ReduceFile method which reduces the log file to the
'   last mvarminlines.  Added logic to call ReduceFile when log file reaches
'   mvarmaxlines. Add properties for minlines and maxlines.

' 04/18/2000 - Modified by Wes Prichard, Optika
'   Added error handlers to procedures WriteLine and OpenLogFile.

' 12/11/2000 - Modified by Wes Prichard, Optika
'   Added new property LogType and modified OpenLogFile & WriteLine to optionally
'   create daily log files versus a single scrolling log file.

' Public Interface:
'Initialize properties then open the log before attempting to write to log.
'If there is an error reducing the log then logging is disabled.
' Methods:
'CloseLogFile - closes an open log file
'OpenLogFile - creates (if necessary) and opens a specified log file
'WriteLine - writes a line to an open log file

' Events:
'(none)

' Properties:
'FileName (ReadOnly) - name of most recently opened log file
'LoggingEnabled - True to enable output to log file
'MaxLines - maximum lines written to a file before it is reduced
'Note - If MaxLines is 0 (or not initialized) the log file will not be reduced
'MinLines - number of lines carried over to a reduced log file from the existing file
'LogType - used to specify the type of log file to create (Daily [default] or Circular)

'*****************************************************************************************
Option Explicit

'Requires project reference to Microsoft Scripting Runtime (scrrun.dll)
Private fs As New FileSystemObject
Private TextStreamOut As TextStream

Public Enum enuLogType
    logDaily               'new log file every day
    logCircular            'one circular log file
End Enum

Private mdatLogDay As Date              'the day associated with the current daily log file

'Local variables to hold property values
Private mvarLoggingEnabled As Boolean   'True enables writing to file
Private mvarFileName As String          'name and path of log file
Private mvarMaxLines As Long            'max lines in a log file
Private mvarMinLines As Long            'lines in a log file after reducing
Private mvarLogType As enuLogType       'type of log to create

Public Property Let MinLines(ByVal vData As Long)
Attribute MinLines.VB_Description = "Number of lines carried over when the log file is reduced"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MinLines = 5
    mvarMinLines = vData
End Property

Public Property Get MinLines() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinLines
    MinLines = mvarMinLines
End Property

Public Property Let MaxLines(ByVal vData As Long)
Attribute MaxLines.VB_Description = "Max lines added to the files before it is reduced"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxLines = 5
    mvarMaxLines = vData
End Property

Public Property Get MaxLines() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxLines
    MaxLines = mvarMaxLines
End Property

Public Property Get FileName() As String
Attribute FileName.VB_Description = "Read-only name and path of log file. Property is set through OpenLogFile method."
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property

Public Property Let LoggingEnabled(ByVal vData As Boolean)
Attribute LoggingEnabled.VB_Description = "True to enable logging."
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LoggingEnabled = True
    mvarLoggingEnabled = vData
End Property

Public Property Get LoggingEnabled() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LoggingEnabled
    LoggingEnabled = mvarLoggingEnabled
End Property

Public Property Let LogType(ByVal vData As enuLogType)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LoggingEnabled = True
    mvarLogType = vData
End Property

Public Property Get LogType() As enuLogType
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LoggingEnabled
    LogType = mvarLogType
End Property

Public Sub OpenLogFile(Optional strFileName As String)
'Open the file specified in the filename property.
'If the filename is blank, no file will be opened

Dim Folder As Folder
Dim strWorkingName As String    'actual log file name
Dim intExt As Integer           'position of file name extionsion in string
   
'Enable error trap
    On Error GoTo errHandler
    
'Init the FileName local property variable if provided
    'If a name was not provided then update property variable
    If strFileName <> vbNullString Then mvarFileName = strFileName
    
    '*** New code
    'If using daily log then...
    If mvarLogType = logDaily Then
        'Find the file extension
        intExt = InStrRev(mvarFileName, ".")
        If intExt > 0 Then
            'File name includes today's date
            strWorkingName = Left(mvarFileName, intExt - 1) & Format(Date, "YYYYMMDD ") & _
                Right(mvarFileName, Len(mvarFileName) - intExt + 1)
        Else    'no extension
            'tack date on to the end of filename
            strWorkingName = mvarFileName & Format(Date, "YYYYMMDD ")
        End If
        mdatLogDay = Date
    Else
        strWorkingName = mvarFileName
    End If
    '*** end new code
    
'If the folder does not exist then create it
    If Not fs.FolderExists(Left(strWorkingName, InStrRev(strWorkingName, "\") - 1)) Then
         Set Folder = fs.CreateFolder(Left(strWorkingName, InStrRev(strWorkingName, "\") - 1))
    End If

'If logging is enabled, and the file is not already open, and a filename has been defined then...
    If (mvarLoggingEnabled = True) And (TextStreamOut Is Nothing) And (strWorkingName <> vbNullString) Then
        Set TextStreamOut = fs.OpenTextFile(strWorkingName, ForAppending, True)
    End If

Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            Err.Raise Err.Number, , "cLog.OpenLogFile-> " & Err.Description
'Resume '@@@
    End Select
End Sub

Public Sub WriteLine(LogData As String)
    
'*** don't raise an error since this procedure can be called from an error handler
'Enable error trap
'    On Error GoTo ErrHandler
    
    If mvarLoggingEnabled = True Then
        '*** new code
        'If using a daily log then...
        Select Case mvarLogType
            Case logDaily
                'If the log day has changed then...
                If mdatLogDay <> Date Then
                    Call CloseLogFile   'close the old file
                    Call OpenLogFile    'open a new file for today
                End If
            
                'If log file is not open then attempt to open it
                If TextStreamOut Is Nothing Then Call OpenLogFile
                
                'Write to the log
                TextStreamOut.WriteLine Format(Now, "mm/dd/yy hh:mm:ss") & " " & LogData
                'Debug.Print TextStreamOut.Line & " " & LogData
                'No size limit to daily logs
            
            Case logCircular
                '*** orig code
                'If log file is not open then attempt to open it
                If TextStreamOut Is Nothing Then Call OpenLogFile
                
                'Write to the log
                TextStreamOut.WriteLine Format(Now, "mm/dd/yy hh:mm:ss") & " " & LogData
                'Debug.Print TextStreamOut.Line & " " & LogData
                
                'If MaxLines is initialized and threshold is reached, then Reduce log file
                If (mvarMaxLines <> 0) And (TextStreamOut.Line > mvarMaxLines) Then Call ReduceFile
                '*** end orig code
        
        End Select
        '*** end new code
        
    End If
    
'Copy log entry to debugging mode immediate window
    If TextStreamOut Is Nothing Then
        Debug.Print "Log: " & LogData
    Else
        Debug.Print TextStreamOut.Line & " " & LogData
    End If
    
'Exit Sub
'
'ErrHandler:
'    Select Case Err.Number
'        Case Else
'            Err.Raise Err.Number, , "cLog.WriteLine-> " & Err.Description
'    End Select
End Sub

Public Sub CloseLogFile()
'    If mvarLoggingEnabled <> False Then    '*** old code
    If Not (TextStreamOut Is Nothing) Then  '*** new code
        TextStreamOut.Close
        Set TextStreamOut = Nothing
    End If
End Sub

Private Sub ReduceFile()
'Reduce the size of the log file by saving up to the last n (mvarminlines) lines and
'throwing away the rest.

Dim lngStartLine As Long    'starting line number for copy
Dim strArray() As String    'array of lines from log file
Dim i As Long            'loop counter

'Enable error handler
    On Error GoTo errHandler

'Initialize the array
    ReDim strArray(mvarMinLines)

'Go to the beginning of the last min_lines lines in the old file
    lngStartLine = TextStreamOut.Line - mvarMinLines - 1       'get current line number
    If lngStartLine <= 0 Then lngStartLine = 1
    TextStreamOut.Close
    Set TextStreamOut = fs.OpenTextFile(mvarFileName, ForReading, False)
    For i = 1 To lngStartLine - 1
        TextStreamOut.SkipLine
    Next i

'Read the lines from the old file and write them to an array
    For i = 1 To mvarMinLines
        strArray(i) = TextStreamOut.ReadLine
    Next i

'Close the old file
    TextStreamOut.Close

'Create a new file, overwriting the old one
    fs.CreateTextFile mvarFileName, True

'Open the new file
    Set TextStreamOut = fs.OpenTextFile(mvarFileName, ForAppending, True)

'Write the line to the new file
    For i = 1 To mvarMinLines
        TextStreamOut.WriteLine strArray(i)
    Next i

    Call WriteLine("Log reduced to last " & mvarMinLines & " lines")
    
Exit Sub

errHandler:
    Select Case Err.Number
        Case Else
            'Disable logging
            mvarLoggingEnabled = False
            Err.Raise Err.Number, , "cLog.ReduceFile-> " & Err.Description
    End Select

End Sub

Private Sub Class_Terminate()
'Flush the I/O buffer by closing the log
    If Not (TextStreamOut Is Nothing) Then
        TextStreamOut.Close
    End If
End Sub
