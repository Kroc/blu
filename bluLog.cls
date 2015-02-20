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

'Dependencies       bluString.cls [bluArrayStrings.cls, <bluArray.cls>]
'Last Updated       20-FEB-15
'Last Update        *   Use a buffer for array reservation
'                   *   'Add' method renamed 'LogFormat', a separate 'Log' method
'                       logs without formatting for extra speed

'--------------------------------------------------------------------------------------

'A class designed for maximum-speed logging of text, either to memory or disk. _
 An array of Strings is used instead of a single String to avoid ever having to _
 reallocate the text

'/// CLASS ////////////////////////////////////////////////////////////////////////////

Private My_Data() As String             'Our array of VB Strings
Private My_Count As Long                'Number of Strings in the array (1-based)
Private My_Buffer As Long

'Used for doing the String formatting before adding a line
Private StrFormat As bluString

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    Set StrFormat = New bluString
    
    Let My_Buffer = 1024
    ReDim Preserve My_Data(1 To My_Buffer) As String
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    'This will unload each of the Strings in our array
    Erase My_Data
    Set StrFormat = Nothing
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Add : Add a line to the log
'======================================================================================
'Text           | A VB String to add to our array
'[FormatValues] | A list of values to replace string markers "{$}" in the text
'======================================================================================
Public Sub LogFormat( _
         ByRef Text As String, _
    ParamArray FormatValues() As Variant _
)
    'The next index will be filled
    Let My_Count = My_Count + 1
    'If this goes over the number of reserved string indexes, _
     increase the size of the array
    If My_Count > My_Buffer Then
        'Increase the buffer exponentially
        Let My_Buffer = My_Buffer * 2
        'Increase the array bounds
        ReDim Preserve My_Data(1 To My_Buffer) As String
    End If
    
    'Are there any values to format?
    If UBound(FormatValues) < 0 Then
        'No, raise just the text given
        Let My_Data(My_Count) = Text
    Else
        'Format the given log message before sending it up
        Let StrFormat.Text = Text
        Let My_Data(My_Count) = StrFormat.Format(FormatValues).Text
    End If
End Sub

'Log : Add a line to the log
'======================================================================================
'Text           | A VB String to add to our array
'======================================================================================
Public Sub Log( _
    ByRef Text As String _
)
    'The next index will be filled
    Let My_Count = My_Count + 1
    'If this goes over the number of reserved string indexes, _
     increase the size of the array
    If My_Count > My_Buffer Then
        'Increase the buffer exponentially
        Let My_Buffer = My_Buffer * 2
        'Increase the array bounds
        ReDim Preserve My_Data(1 To My_Buffer) As String
    End If
    
    'Add the given text to the log
    Let My_Data(My_Count) = Text
End Sub

'Clear : Empty the array
'======================================================================================
Public Sub Clear()
    Erase My_Data: Let My_Count = 0
End Sub

'Concatenate : Return all Strings in the array as a combined String
'======================================================================================
'Returns        | A VB String of the combined array contents
'======================================================================================
Public Function Concatenate( _
) As String
    'Prepare a bluString buffer large enough for the whole array. _
     VB's concatenation will be too slow
    Dim ReturnString As bluString
    Set ReturnString = New bluString
    
    Dim Length As Long, i As Long
    For i = 1 To My_Count
        'TODO: Check for overflow
        Let Length = Length + Len(My_Data(i)) + Len(vbCrLf)
    Next i
    
    Let ReturnString.Buffer = Length
    
    For i = 1 To My_Count
        Call ReturnString.Append(My_Data(i)).Append(vbCrLf)
    Next i
    
    Let Concatenate = ReturnString.Text
    Set ReturnString = Nothing
End Function

'Count : Number of Strings in our array
'======================================================================================
Public Property Get Count() As Long: Let Count = My_Count: End Property

'Item : Get access to a String in the array
'======================================================================================
'Index          | An Index number for the array contents
'---------------+----------------------------------------------------------------------
'Returns        | The VB String in the array at the given Index
'======================================================================================
Public Property Get Item( _
    ByVal Index As Long _
) As String
    'If there are no items in the array, we're out of bounds whatever the Index is
    If My_Count = 0 Then Err.Raise (9): Exit Property
    'Is the index within bounds?
    If (Index < 1) Or (Index > My_Count) Then Err.Raise (9): Exit Property
    'Return the VB String from our array
    Let Item = My_Data(Index)
End Property

'Length : Get the length of a String in the array
'======================================================================================
'Index          | An Index number for the array contents
'---------------+----------------------------------------------------------------------
'Returns        | The length in Characters (not Bytes) of the Indexed String
'======================================================================================
Public Property Get Length( _
    ByVal Index As Long _
) As Long
    'If there are no items in the array, we're out of bounds whatever the Index is
    If My_Count = 0 Then Err.Raise (9): Exit Property
    'Is the index within bounds?
    If (Index < 1) Or (Index > My_Count) Then Err.Raise (9): Exit Property
    'Return the length (in characters) of the String in the array
    Let Length = Len(My_Data(Index))
End Property

'LengthBytes : Get the length of a String in the array, in Bytes
'======================================================================================
'Index          | An Index number for the array contents
'---------------+----------------------------------------------------------------------
'Returns        | The length in Bytes (not Characters) of the Indexed String
'======================================================================================
Public Property Get LengthBytes( _
    ByVal Index As Long _
) As Long
    'Two Bytes to each Character
    Let LengthBytes = Me.Length(Index) * 2
End Property

'Pointer : Get access to a String in the array
'======================================================================================
'Index          | An Index number for the array contents
'---------------+----------------------------------------------------------------------
'Returns        | The memory location of the String contents
'======================================================================================
Public Property Get Pointer( _
    ByVal Index As Long _
) As Long
    'If there are no items in the array, we're out of bounds whatever the Index is
    If My_Count = 0 Then Err.Raise (9): Exit Property
    'Is the index within bounds?
    If (Index < 1) Or (Index > My_Count) Then Err.Raise (9): Exit Property
    'Return the String pointer -- be careful with this!
    Let Pointer = StrPtr(My_Data(Index))
End Property

'/// PRIVATE INTERFACE ////////////////////////////////////////////////////////////////

'Increase : Increases the size of the array
'======================================================================================
'Returns        | The new index number
'======================================================================================
Private Function Increase() As Long
    'The next index will be filled
    Let My_Count = My_Count + 1
    Let Increase = My_Count
    'Increase the array bounds
    ReDim Preserve My_Data(1 To Increase) As String
End Function
