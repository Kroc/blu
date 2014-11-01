VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bluString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================================
'blu : A Modern Metro-esque graphical toolkit; Copyright (C) Kroc Camen, 2013-14
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS :: bluString

'A replacement for VB6's terribly slow String handling of any kind. _
 A buffer is used to minimise the number of times the reserved memory has to be _
 expanded and methods for most of VB6's String functions are provided to work _
 directly on the String bytes

'/// API //////////////////////////////////////////////////////////////////////////////

'Copy raw memory from one place to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366535%28v=vs.85%29.aspx>
Private Declare Sub api_RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef DestinationPointer As Any, _
    ByRef SourcePointer As Any, _
    ByVal Length As Long _
)

'Fill memory with zeroes _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366920%28v=vs.85%29.aspx>
Private Declare Sub api_RtlZeroMemory Lib "kernel32" Alias "RtlZeroMemory" ( _
    ByRef DestinationPointer As Any, _
    ByVal Length As Long _
)

'Cast String to Array: _
 --------------------------------------------------------------------------------------
'By manipulating some pointers we can cast a VB String directly into an array without _
 having to make a copy of the string, nor waste memory on a separate array. _
 This is obviously very fast, especially for long strings
'<vbforums.com/showthread.php?729385>

'This is VB6's internal structure used for VB arrays _
 <msdn.microsoft.com/en-us/library/ms221482(v=vs.85).aspx>
Private Type SAFEARRAY1D
    cDims       As Integer              'The count of dimensions
    fFeatures   As Integer              'Flags used by the SafeArray
    cbElements  As Long                 'The size of an array element
    cLocks      As Long                 'Lock count
    pvData      As Long                 'Pointer to the data
    cElements   As Long                 'The number of elements in the dimension
    lLbound     As Long                 'The lower bound of the dimension
End Type

'- An array that is allocated on the stack
Private Const FADF_AUTO      As Integer = &H1
'- An array that may not be resized or reallocated
Private Const FADF_FIXEDSIZE As Integer = &H10

'Undocumented VB6 API to get at the array pointer
Private Declare Function msvbvm60_VarPtr Lib "msvbvm60.dll" Alias "VarPtr" ( _
    ByRef ArrayVar() As Any _
) As Long

'Undocumented VB6 API to write 4-bytes (Long) of memory
Private Declare Sub msvbvm60_PutMem4 Lib "msvbvm60.dll" Alias "PutMem4" ( _
    ByVal Addr As Long, _
    ByVal NewVal As Long _
)

Private CastStrSA1D As SAFEARRAY1D      'A dummy VB6 array header
Private CastStrPtr  As Long             'Pointer to our dummy array
Private CastStrArr() As Integer         'Empty array that will be hacked

'/// CLASS STORAGE ////////////////////////////////////////////////////////////////////

'This is where we'll store our "string" as an array of Unicode (UCS-2) points
Private Data() As Integer
'The buffer and content-length will be a 1-based count of characters (not bytes); _
 UCS-2 uses two bytes per character
Private My_Buffer As Long: Private Const DEFAULT_BUFFER As Long = 256
Private My_Length As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'Prepare our dummy VB6 String header that will be used as _
     a means to cast VB Strings directly into arrays
    With CastStrSA1D
        Let .cDims = 1
        Let .fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        Let .cbElements = 2&
        Let .cLocks = 1&
        Let .lLbound = 0&
    End With
    
    'Get the pointer to the undimmed array that will act as zombie host to _
     any Strings that we want to use as arrays
    Let CastStrPtr = msvbvm60_VarPtr(CastStrArr())
    'Swap the String headers over
    Call msvbvm60_PutMem4(CastStrPtr, VarPtr(CastStrSA1D))
    
    'NOTE: The buffer is not initialised to begin with _
          (therefore not wasting memory for a string that is not yet populated, _
           but more so that the reserved memory does not have to be reallocated _
           if a String larger than the default buffer is created first)
End Sub

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Free up the reserved memory
    Erase Data
    
    'Clean up our array hacking (for casting a String to an array) _
     otherwise VB will crash when it frees up the memory
    Call msvbvm60_PutMem4(CastStrPtr, 0&)
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Append : Add another string to the end of the current string _
 ======================================================================================
Public Function Append( _
    ByRef InputString As String _
) As bluString
    'Appending a null-string will do nothing
    If InputString = vbNullString Or LenB(InputString) = 0 Then Exit Function
    'The current length will be the starting position of the appended string
    Dim Start As Long: Let Start = My_Length
    'Increase the length accordingly
    Let Me.Length = My_Length + Len(InputString)
    'Copy the new string in
    Call api_RtlMoveMemory(Data(Start), ByVal InputString, LenB(InputString))
    
    'Allow method chaining
    Set Append = Me
End Function

'Buffer : Number of characters to reserve in memory for the string _
 ======================================================================================
Public Property Get Buffer() As Long: Let Buffer = My_Buffer: End Property
Public Property Let Buffer( _
    ByRef NumberOfChars As Long _
)
    'Store the new desired buffer length (1-based)
    Let My_Buffer = NumberOfChars
    'If the buffer is cut shorter than the current content, we need to adjust it too
    If My_Buffer < My_Length Then Let My_Length = My_Buffer
    'Resize the data store. Note that though `My_Buffer` is 1-based we abuse this _
     to provide room for a null-terminator at the end of the string
    ReDim Preserve Data(0 To My_Buffer) As Integer
    'Ensure it's null-terminated!
    Let Data(My_Buffer) = 0
End Property

'Char : The Unicode (UCS-2) of the character at the provided character index _
 ======================================================================================
Public Property Get Char( _
    ByRef Index As Long _
) As Integer
    'NOTE: Though `My_Length` is 1-based, but `Index` is 0-based, we allow the user _
           to access the spare element at the end as it will be the null terminator, _
           which the end-user may wish to check for (blind-looping, for example)
           
    'If a character index is provided outside of the length of the string, error
    If Index > My_Length Then Err.Raise (9): Exit Property
    'Return the character value in question. The user can use `And &HFF` to get ASCII
    Let Char = Data(Index)
End Property
Public Property Let Char( _
    ByRef Index As Long, ByRef Char As Integer _
)
    'If a character index is provided outside of the length of the string, error. _
     (note that, compared to the getter, you may not write to the null terminator)
    If Index >= My_Length Then Err.Raise (9): Exit Property
    'Set the character value at that point
    Let Data(Index) = Char
End Property

'Clear : Wipe the stored string, fill the buffer with nulls _
 ======================================================================================
Public Function Clear() As bluString
    'You can't clear a null-string any more than it is!
    If My_Length > 0 Then
        'This does not change the size of the buffer, should it have increased. _
         If you intend to reuse a large string for a small string, _
         create a new bluString instance instead
        Call api_RtlZeroMemory(Data(0), My_Length * 2)
        'Mark the string as being null-length
        Let My_Length = 0
    End If
    'Allow method chaining
    Set Clear = Me
End Function

'Clone : Duplicate our string into another bluString instance _
 ======================================================================================
Public Function Clone() As bluString
    'Instantiate the return object
    Set Clone = New bluString
    'Set the destination size
    Let Clone.Length = My_Length
    'Copy the text across (if there is any)
    If My_Length > 0 Then
        Call api_RtlMoveMemory(ByVal Clone.Pointer, Data(0), My_Length * 2)
    End If
End Function

'Compare : A fast method to check if one string equals another _
 ======================================================================================
Public Function Compare( _
    ByRef Comparison As String _
) As Boolean
    '
End Function

'Format : Format the text according to a pattern _
 ======================================================================================
Public Function Format() As bluString
    'TODO
    
    'Allow method chaining
    Set Format = Me
End Function

'Join : Appends a bluString to the end of this one _
 ======================================================================================
Public Function Join( _
    ByRef InputString As bluString _
) As bluString
    'TODO
    
    'Allow method chaining
    Set Join = Me
End Function

'LCase : Convert our string to lowercase _
 ======================================================================================
'TODO: We'll need a hefty look up table of UCS-2 case conversion
Public Function LCase() As bluString
    'TODO
    
    'Allow method chaining
    Set LCase = Me
End Function

'Left : Cut the string short to a given number of characters _
 ======================================================================================
Public Function Left( _
    ByRef NumberOfCharacters As Long _
) As bluString
    'You can't increase the length with this!
    If NumberOfCharacters > My_Length Then Err.Raise (9): Exit Function
    'Do the reduction
    Let Me.Length = NumberOfCharacters
    
    'Allow method chaining
    Set Left = Me
End Function

'Length : Length of the actual stored text (1-based), not the buffer _
 ======================================================================================
Public Property Get Length() As Long: Let Length = My_Length: End Property
Public Property Let Length( _
    ByRef NewLength As Long _
)   'Store the new desired content length (1-based)
    Let My_Length = NewLength
    'If the buffer has not been initialised, start with the default length
    If My_Buffer = 0 Then Let My_Buffer = DEFAULT_BUFFER
    
TestBuffer:
    'If this increases the data too large
    If My_Length > My_Buffer Then
        'Double the length of the buffer
        Let My_Buffer = My_Buffer * 2
        'Re-check if this is large enough
        GoTo TestBuffer
    End If
    'Apply the new buffer length _
     (this will do the actual memory allocation)
    Let Me.Buffer = My_Buffer
    
    'WARNING: Increasing the content-length via this property, instead of by-proxy _
     via methods such as `Append` may leave 'junk' data in the increased space!
    Let Data(My_Length) = 0
End Property

'Mid : Reduce the string to a sub-string portion _
 ======================================================================================
Public Function Mid() As bluString
    'TODO
    
    'Allow method chaining
    Set Mid = Me
End Function

'Pointer : Allow access to the "string" pointer for external copying _
 ======================================================================================
Public Property Get Pointer() As Long
    'If a null string, return a null-string pointer
    If My_Length = 0 Then
        Let Pointer = VarPtr(vbNullString)
    Else
        Let Pointer = VarPtr(Data(0))
    End If
End Property

'Prepend : Add our string to the end of the one provided _
 ======================================================================================
'And no, "prefix" is not the right word for this!
Public Function Prepend( _
    ByRef InputString As String _
) As bluString
    'Prepending a null-string will do nothing
    If InputString = vbNullString Or LenB(InputString) = 0 Then Exit Function
    'The current length will be the new starting position of our string
    Dim Start As Long: Let Start = My_Length
    'Increase the length accordingly
    Let Me.Length = My_Length + Len(InputString)
    'Shift our string up
    Call api_RtlMoveMemory(Data(Start), Data(0), My_Length * 2)
    'Copy the new string in
    Call api_RtlMoveMemory(Data(0), ByVal InputString, LenB(InputString))
    
    'Allow method chaining
    Set Prepend = Me
End Function

'ReadFile : Import our string from a file _
 ======================================================================================
Public Function ReadFile( _
    ByRef FilePath As String _
) As bluString
    'TODO
    
    'Allow method chaining
    Set ReadFile = Me
End Function

'Right : Keep only the given right-most number of characters _
 ======================================================================================
Public Function Right() As bluString
    'TODO
    
    'Allow method chaining
    Set Right = Me
End Function

'Replace : Search for a sub-string and replace it with another _
 ======================================================================================
Public Function Replace( _
    ByRef SearchFor As String, _
    ByRef ReplaceWith As String _
) As bluString
    'TODO
    
    'Allow method chaining
    Set Replace = Me
End Function

'Text : Get/Let the contents as native VB strings _
 ======================================================================================
Public Property Get Text() As String
    'If ours is a null-string, return a null-string
    If My_Length = 0 Then
        Let Text = vbNullString
    Else
        'Pad the return String to the right size
        Let Text = Space$(My_Length)
        'Take the string, and attach it to our dummy array header
        With CastStrSA1D
            Let .pvData = StrPtr(Text)
            Let .cElements = LenB(Text)
        End With
        'Copy our String into the return field
        Call api_RtlMoveMemory(CastStrArr(0), Data(0), My_Length * 2)
    End If
End Property
Public Property Let Text( _
    ByRef NewString As String _
)   'If a null-string is provided then set ours to a null string
    If NewString = vbNullString Or LenB(NewString) = 0 Then
        Call Me.Clear
    Else
        'Set the new length of our String. The buffer will be increased appropriately
        Let Me.Length = Len(NewString)
        'Copy the String data over
        Call api_RtlMoveMemory(Data(0), ByVal NewString, LenB(NewString))
    End If
End Property

'Trim : Strip whitespace characters from the beginning and end _
 ======================================================================================
'TODO: We may want a look-up table of all whitespace chars in UCS-2
Public Function Trim() As bluString
    'TODO
    
    'Allow method chaining
    Set Trim = Me
End Function

'TrimLeft : Strip whitespace characters from the beginning _
 ======================================================================================
Public Function TrimLeft() As bluString
    'TODO
    
    'Allow method chaining
    Set TrimLeft = Me
End Function

'TrimRight : Strip whitespace characters from the end _
 ======================================================================================
Public Function TrimRight() As bluString
    'TODO
    
    'Allow method chaining
    Set TrimRight = Me
End Function

'UCase : Convert our string to UPPERCASE _
 ======================================================================================
'TODO: We'll need a hefty look up table of UCS-2 case conversion
Public Function UCase()
    'TODO
    
    'Allow method chaining
    Set TrimRight = Me
End Function

'WriteFile : Export our string to a file _
 ======================================================================================
Public Function WriteFile( _
    ByRef FilePath As String _
) As bluString
    'TODO
    
    'Allow method chaining
    Set WriteFile = Me
End Function
