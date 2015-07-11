VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bluLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'blu : A Modern Metro-esque graphical toolkit; Copyright (C) Kroc Camen, 2013-15
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS :: bluLog

'Dependencies
'Last Updated       12-JUL-15
'Last Update        Rewrote to blindly append to file, rather than keep in memory

'--------------------------------------------------------------------------------------

'A class designed for maximum-speed logging of text to disk

'/// CLASS ////////////////////////////////////////////////////////////////////////////

'Where the log will be saved
Private My_FilePath As String

'Used for doing the String formatting before adding a line
Private StrFormat As bluString

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    Set StrFormat = New bluString
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    Set StrFormat = Nothing
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

Public Property Get FilePath() As String
    Let FilePath = My_FilePath
End Property
Public Property Let FilePath(ByRef NewPath As String)
    Let My_FilePath = FilePath
End Property

'LogFormat : Add a line to the log, with marker replacement
'======================================================================================
'Text           | A VB String to add to our array
'[FormatValues] | A list of values to replace string markers "{$}" in the text
'======================================================================================
Public Sub LogFormat( _
         ByRef Text As String, _
    ParamArray FormatValues() As Variant _
)
    'Are there any values to format?
    If UBound(FormatValues) >= 0 Then
        'Format the given log message before sending it up
        Let StrFormat.Text = Text
        Let Text = StrFormat.Format(FormatValues).Text
    End If
    
    If LenB(My_FilePath) <> 0 Then
        'TODO: Replace with API methods
    
        Dim FileNumber As Long
        Let FileNumber = FreeFile
        
        Open My_FilePath For Append Lock Write As #FileNumber
            Print #FileNumber, Text
        Close #FileNumber
    End If
End Sub

'Log : Add a line to the log
'======================================================================================
'Text           | A VB String to add to our array
'======================================================================================
Public Sub Log( _
    Optional ByRef Text As String = vbNullString _
)
    If LenB(My_FilePath) <> 0 Then
        'TODO: Replace with API methods
    
        Dim FileNumber As Long
        Let FileNumber = FreeFile
        
        Open My_FilePath For Append Lock Write As #FileNumber
            Print #FileNumber, Text
        Close #FileNumber
    End If
End Sub
