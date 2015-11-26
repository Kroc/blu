Attribute VB_Name = "blu"
Option Explicit
'======================================================================================
'blu : A Modern Metro-esque graphical toolkit; Copyright (C) Kroc Camen, 2013-15
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'MODULE :: blu

'Dependencies       blu_Sys, blu_Strings
'Last Updated       12-JUL-15
'Last Update        Initial version - INCOMPLETE

'/// API //////////////////////////////////////////////////////////////////////////////

'Get the Command-line Argument for the current process: _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms683156(v=vs.85).aspx>
Private Declare Function api_GetCommandLine Lib "kernel32" Alias "GetCommandLineW" ( _
) As Long

Private Declare Function api_lstrcpyn Lib "kernel32" Alias "lstrcpynW" ( _
    ByVal lpString1 As Long, _
    ByVal lpString2 As Long, _
    ByVal iMaxLength As Long _
) As Long

'<msdn.microsoft.com/en-us/library/bb773602(v=vs.85).aspx>
Private Declare Function api_PathGetArgs Lib "shlwapi" Alias "PathGetArgsW" ( _
    ByVal SourcePathPointer As Long _
) As Long

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

Public sys As New blu_Sys

Public Strings As New blu_Strings

Public FileSystem As New blu_FileSystem

'CommandParams : Get the Command-Line Arguments, with Unicode support
'======================================================================================
'Returns        | A string array of each word from the command line parameters
'======================================================================================
Public Function CommandParams() As String()
    'With thanks to TannerHelland / PhotoDemon for the demonstration and "dilettante" _
     for the original: <xtremevbtalk.com/showthread.php?t=325868>
    
    'A bluString will be used to store the command line, _
     we can use its `SplitWords` method to make the rest easy
    Dim CommandStr As bluString
    Set CommandStr = New bluString
    Let CommandStr.Buffer = MAX_PATH
    
    'When in the IDE, use the command line provided in the Project Properties
    If blu.sys.InIDE Then
        Let CommandStr.Text = Command$
    Else
        'Use Windows API to get a pointer to the command line string
        Dim CommandStrPtr As Long
        Let CommandStrPtr = api_GetCommandLine()
        
        'The command string provided by API will contain the EXE path, _
         this API will move the pointer forward until after the EXE path
        Let CommandStrPtr = api_PathGetArgs(CommandStrPtr)
        
        If CommandStrPtr <> 0 Then
            'This is a null-terminated string and we _
             won't know the length until we count
            Dim CommandStrLen As Long
            Let CommandStrLen = bluW32.StringLengthUpToNull(CommandStrPtr)
            
            If CommandStrLen <> 0 Then
                Let CommandStr.Length = CommandStrLen
                'Copy the command line string into our bluString
                Call api_lstrcpyn(CommandStr.Pointer, CommandStrPtr, CommandStrLen + 1)
            End If
        End If
    End If
    
'    MsgBox CommandStr.Text
    
    Dim Words() As String
    Let Words = CommandStr.SplitWords(True)
    
    Dim i As Long
    For i = LBound(Words) To UBound(Words)
        Debug.Print Words(i)
    Next i
    
    Let CommandParams = Words
    
    'Free the bluString
    Set CommandStr = Nothing
End Function
