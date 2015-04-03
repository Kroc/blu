VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bluText"
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
'CLASS :: bluString

'Dependencies       bluFileSystem.bas, bluString.cls
'Last Updated
'Last Update

'Always get the latest version of this file at <github.com/Kroc/blu>. _
 Feedback can be sent to <kroc@camendesign.com>

'--------------------------------------------------------------------------------------

'For multiple lines of text, such as generating logs and reading text files

'/// CLASS ////////////////////////////////////////////////////////////////////////////

Private Lines() As String
Private Line As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    Erase Lines: Let Line = 0
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

Private Sub Load( _
    ByRef FilePath As String _
)
    'All the messy API file handling is shared in bluFileSystem
    Dim FileBuffer() As Byte
    If bluFileSystem.ReadBinaryFile_AsArray( _
        FilePath, FileBuffer _
    ) Then
        Stop
        Exit Sub
    End If
End Sub
