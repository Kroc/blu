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

'Special features of this class:

' * Unicode (UCS-2) support

' * Null-strings are supported, so no memory is allocated until content is added _
    (other such string classes always seem to allocate their default buffer)

' * Method-chaining allowing for some nice OOP behaviour; _
    `bluString.Clone.Left(12).Append("...")`

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

'Undocumented VB6 API to write 4-bytes (Long) of memory
Private Declare Sub api_PutMem4 Lib "msvbvm60" Alias "PutMem4" ( _
    ByVal Addr As Long, _
    ByVal NewVal As Long _
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
Private Declare Function api_VarPtr Lib "msvbvm60" Alias "VarPtr" ( _
    ByRef ArrayVar() As Any _
) As Long

Private CastStrSA1D As SAFEARRAY1D      'A dummy VB6 array header
Private CastStrPtr  As Long             'Pointer to our dummy array
Private CastStrArr() As Integer         'Empty array that will be hacked

'Windows Version: _
 --------------------------------------------------------------------------------------

'In VB6 True is -1 and False is 0, but in the Win32 API it's 1 for True
Private Enum BOOL
    API_TRUE = 1
    API_FALSE = 0
End Enum

'Structure for obtaining the Windows version _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms724834%28v=vs.85%29.aspx>
Private Type OSVERSIONINFO
    SizeOfMe As Long
    MajorVersion As Long
    MinorVersion As Long
    BuildNumber As Long
    PlatformID As Long
    ServicePack As String * 128
End Type

'Get the Windows version _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms724451%28v=vs.85%29.aspx>
Private Declare Function api_GetVersionEx Lib "kernel32" Alias "GetVersionExA" ( _
    ByRef VersionInfo As OSVERSIONINFO _
) As BOOL

'Locale Mapping (for comparison and case conversion): _
 --------------------------------------------------------------------------------------
'This page helped with working out correct methods of using Unicode API calls _
 <www.xtremevbtalk.com/showthread.php?t=68956>
'This page helped with translating the MSDN documentation into VB6; _
 <www.ex-designz.net/apidetail.asp?api_id=383>

'Get the Locale Identifier (LCID) of this app _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318127(v=vs.85).aspx>
'This is used for Windows XP support as Vista+ use Locale Name strings
Private Declare Function api_GetThreadLocale Lib "kernel32" Alias "GetThreadLocale" ( _
) As Long

'Unicode & Locale-aware case conversion (amongst other things) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318700(v=vs.85).aspx>
'Private Declare Function api_LCMapString Lib "kernel32" Alias "LCMapStringW" ( _
') As Long

'Unicode & Locale-aware string comparison (Windows XP) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd317759(v=vs.85).aspx>
Private Declare Function api_CompareString Lib "kernel32" Alias "CompareStringW" ( _
    ByVal LocaleID As Long, _
    ByVal CompareFlags As CompareString_Flags, _
    ByVal String1Pointer As Long, _
    ByVal String1Length As Long, _
    ByVal String2Pointer As Long, _
    ByVal String2Length As Long _
) As CSTR_

'Unicode & Locale-aware string comparison (Windows Vista+) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd317761(v=vs.85).aspx>
Private Declare Function api_CompareStringEx Lib "kernel32" Alias "CompareStringEx" ( _
    ByVal LocaleNamePointer As Long, _
    ByVal CompareFlags As CompareString_Flags, _
    ByVal String1Pointer As Long, _
    ByVal String1Length As Long, _
    ByVal String2Pointer As Long, _
    ByVal String2Length As Long, _
    ByVal VersionInformationPointer As Long, _
    ByVal ReservedPointer As Long, _
    ByVal Param As Long _
) As CSTR_

Private Enum CompareString_Flags
    'Ignore case, as linguistically appropriate
    LINGUISTIC_IGNORECASE = &H10&
    'Ignore nonspacing characters, as linguistically appropriate
    LINGUISTIC_IGNOREDIACRITIC = &H20&
    'Legacy ignore case -- less accurate than LINGUISTIC_IGNORECASE
    NORM_IGNORECASE = &H1&
    'Treat equivilent Japanese Hiragana and Katakana as the same
    NORM_IGNOREKANATYPE = &H10000
    'Legacy version of `LINGUISTIC_IGNOREDIACRITIC`
    NORM_IGNORENONSPACE = &H2&
    'Ignore symbols and punctuation
    NORM_IGNORESYMBOLS = &H4&
    'Ignore the difference between half-width and full-width characters
    NORM_IGNOREWIDTH = &H20000
    'Use linguistic rules for casing, instead of file system rules (default)
    NORM_LINGUISTIC_CASING = &H8000000
    'Treat punctuation the same as symbols
    SORT_STRINGSORT = &H1000
End Enum

'Comparison API return values; _
 You can subtract 2 from these to get C-style values
Private Enum CSTR_
    CSTR_LESS_THAN = 1
    CSTR_EQUAL = 2
    CSTR_GREATER_THAN = 3
End Enum

'/// CLASS STORAGE ////////////////////////////////////////////////////////////////////

'We'll cache the Windows version here as it's very slow to do this every time we _
 call methods that use alternative APIs for Vista+
Private WindowsVersion As Single

'This is where we'll store our "string" as an array of Unicode (UCS-2) points
Private Data() As Integer
'The buffer and content-length will be a 1-based count of characters (not bytes); _
 UCS-2 uses two bytes per character
Private My_Buffer As Long: Private Const DEFAULT_BUFFER As Long = 128
Private My_Length As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'Get the Windows version as Windows Vista and above use alternative APIs for _
     locale-dependent functions (such as `Compare`)
    Dim VersionInfo As OSVERSIONINFO
    Let VersionInfo.SizeOfMe = Len(VersionInfo)
    'NOTE: If the app is in compatibility mode, this will return the compatible _
     Windows version, not the actual version; but that's fine with me
    'Windows 8.1 and above also cause this API to always return as Windows 8.0 unless _
     the app opts-in to correct version info via manifest, but this isn't an issue to _
     bluString as we are only concerned with Windows Vista+
    If api_GetVersionEx(VersionInfo) = API_TRUE Then
        Let WindowsVersion = _
            Val(VersionInfo.MajorVersion & "." & VersionInfo.MinorVersion)
    End If
    
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
    Let CastStrPtr = api_VarPtr(CastStrArr())
    'Backup the pointer for the dummy header
    Call api_PutMem4(CastStrPtr, VarPtr(CastStrSA1D))
    
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
    Call api_PutMem4(CastStrPtr, 0&)
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'ANSI : Provides an ANSI version (1-byte per character) of the String _
 ======================================================================================
Public Function ANSI() As String
    'Use this function to get strings for ANSI Windows APIs
    'TODO
End Function

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

'Char : The Unicode (UCS-2) point of the character at the provided character index _
 ======================================================================================
Public Property Get Char( _
    ByRef Index As Long _
) As Long
    'NOTE: Though `My_Length` is 1-based, but `Index` is 0-based, we allow the user _
           to access the spare element at the end as it will be the null terminator, _
           which the end-user may wish to check for (blind-looping, for example)
           
    'If a character index is provided outside of the length of the string, error
    If Index > My_Length Then Err.Raise (9): Exit Property
    'Return the character value in question. The user can use `And &HFF` to get ASCII
    Let Char = Data(Index)
End Property
Public Property Let Char( _
    ByRef Index As Long, ByRef Char As Long _
)
    'If a character index is provided outside of the length of the string, error. _
     (note that, compared to the getter, you may not write to the null terminator)
    If Index >= My_Length Then Err.Raise (9): Exit Property
    'VB6 can only do signed Integers, so Unicode points 32'768+ will error unless _
     we convert them to their signed equivilent to stuff them into two bytes
    If Char >= 32768 Then
        Let Data(Index) = (Char And &HFFFF&) - 65536
    Else
        Let Data(Index) = (Char And &HFFFF&)
    End If
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
             ByRef Comparison As String, _
    Optional ByRef CompareMethod As VBA.VbCompareMethod = vbTextCompare _
) As Boolean
    'The APIs used default to binary search, with a flag for ignoring case
    Dim Flags As CompareString_Flags
    If CompareMethod = vbTextCompare Then Let Flags = LINGUISTIC_IGNORECASE
    
    'Check for Windows Vista and above:
    If WindowsVersion >= 6 Then
        'NOTE: The use of null on the first parameter causes Windows to use the _
         user's current locale, so we don't need to manually fetch the locale name
        Let Compare = (api_CompareStringEx( _
            0&, Flags, _
            StrPtr(Comparison), -1, _
            VarPtr(Data(0)), -1, _
            0&, 0&, 0& _
        ) = CSTR_EQUAL)
    Else
        'Fall back to the XP-compatible API
        Let Compare = (api_CompareString( _
            api_GetThreadLocale(), Flags, _
            StrPtr(Comparison), -1, _
            VarPtr(Data(0)), -1 _
        ) = CSTR_EQUAL)
    End If
End Function

'Format : Format the text according to a pattern _
 ======================================================================================
Public Function Format() As bluString
    'TODO: we need to decide what form this is going to take. _
     a `printf` impl.? use VB's format? Or go for separate methods like `formatNumber`?

    'Allow method chaining
    Set Format = Me
End Function

'Insert : Insert some text in the middle of the existing string _
 ======================================================================================
Public Function Insert() As bluString
    'TODO
    
    'Allow method chaining
    Set Insert = Me
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

'Remove : Cut some text out of the middle of the existing string _
 ======================================================================================
Public Function Remove() As bluString
    'TODO
    
    'Allow method chaining
    Set Remove = Me
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

'Right : Keep only the given right-most number of characters _
 ======================================================================================
Public Function Right() As bluString
    'TODO
    
    'Allow method chaining
    Set Right = Me
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
        Call api_RtlMoveMemory(Data(0), ByVal StrPtr(NewString), LenB(NewString))
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

'/// PRIVATE INTERFACE ////////////////////////////////////////////////////////////////

