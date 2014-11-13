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

' * Unicode (UCS-2) throughout

' * Null-strings are supported, so no memory is allocated until content is added _
    (other such string classes always seem to allocate their default buffer)

' * Method-chaining allowing for some nice OOP behaviour; _
    `bluString.Clone.Left(12).Append("...")`
    
' * Ridiculously fast! _
    Half a million string comparisons per second? No problem

'/// API //////////////////////////////////////////////////////////////////////////////

'Copy raw memory from one place to another _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366535%28v=vs.85%29.aspx>
Private Declare Sub api_RtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal DestinationPointer As Long, _
    ByVal SourcePointer As Long, _
    ByVal Length As Long _
)

'Fill memory with zeroes _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa366920%28v=vs.85%29.aspx>
Private Declare Sub api_RtlZeroMemory Lib "kernel32" Alias "RtlZeroMemory" ( _
    ByVal DestinationPointer As Long, _
    ByVal Length As Long _
)

'Undocumented VB6 API to write 4-bytes (Long) of memory
Private Declare Sub api_PutMem4 Lib "msvbvm60" Alias "PutMem4" ( _
    ByVal DestinationPointer As Long, _
    ByVal Value As Long _
)

'Cast String to Array: _
 --------------------------------------------------------------------------------------
'By manipulating some pointers we can cast a VB String directly into an array without _
 having to make a copy of the string, nor waste memory on a separate array. _
 This is obviously very fast, especially for long strings
'<vbforums.com/showthread.php?729385>

'This is VB6's internal structure used for arrays _
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

'Locale Mapping (for case conversion and string comparison): _
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

'Unicode & Locale-aware case conversion (Windows XP) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318700(v=vs.85).aspx>
Private Declare Function api_LCMapString Lib "kernel32" Alias "LCMapStringW" ( _
    ByVal LocaleID As Long, _
    ByVal MapFlags As LCMAP, _
    ByVal SourceStringPointer As Long, _
    ByVal SourceStringLength As Long, _
    ByVal DestinationStringPointer As Long, _
    ByVal DestinationStringLength As Long _
) As Long

'Unicode & Locale-aware case conversion (Windows Vista+) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318702(v=vs.85).aspx>
Private Declare Function api_LCMapStringEx Lib "kernel32" Alias "LCMapStringEx" ( _
    ByVal LocaleNamePointer As Long, _
    ByVal MapFlags As LCMAP, _
    ByVal SourceStringPointer As Long, _
    ByVal SourceStringLength As Long, _
    ByVal DestinationStringPointer As Long, _
    ByVal DestinationStringLength As Long, _
    ByVal VersionInformationPointer As Long, _
    ByVal ReservedPointer As Long, _
    ByVal SortHandle As Long _
) As Long

Private Enum LCMAP
    LCMAP_LINGUISTIC_CASING = &H1000000 'Manage accents &c. in lower/upper casing
    LCMAP_LOWERCASE = &H100&            'Convert to lower case
    LCMAP_UPPERCASE = &H200&            'Convert to UPPER CASE
    LCMAP_SORTKEY = &H400&              'Generate a sort key
    LCMAP_BYTEREV = &H800&              'Reverse the bytes
    LCMAP_HIRAGANA = &H100000           'Change Japanese Katakana to Hiragana
    LCMAP_KATAKANA = &H200000           'Change Japanese Hiragana to Katakana
    LCMAP_HALFWIDTH = &H400000          'Use standard (half-width) characters
    LCMAP_FULLWIDTH = &H800000          'Use wide-characters (i.e. Latin in Chinese)
    'Change traditional Chinese to simplified Chinese and vice-versa
    LCMAP_SIMPLIFIED_CHINESE = &H2000000
    LCMAP_TRADITIONAL_CHINESE = &H4000000
End Enum

'--------------------------------------------------------------------------------------

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

'Unicode character properties: _
 --------------------------------------------------------------------------------------
'Returns an array of Unicode character properties for a string _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318119(v=vs.85).aspx>
Private Declare Function api_GetStringTypeW Lib "kernel32" Alias "GetStringTypeW" ( _
    ByVal InfoType As CT, _
    ByVal SourceStringPointer As Long, _
    ByVal SourceStringLength As Long, _
    ByVal CharTypePointer As Long _
) As BOOL

Private Enum CT
    CT_CTYPE1 = &H1                     'Character type info (unicode properties)
    CT_CTYPE2 = &H2                     'Bi-directional layout information
    CT_CTYPE3 = &H3                     'Text-processing information
End Enum

Public Enum C1
    C1_UPPER = 2 ^ 0                    'Uppercase
    C1_LOWER = 2 ^ 1                    'Lowercase
    C1_DIGIT = 2 ^ 2                    'Decimal digit
    C1_SPACE = 2 ^ 3                    'Space characters
    C1_PUNCT = 2 ^ 4                    'Punctuation
    C1_CNTRL = 2 ^ 5                    'Control characters
    C1_BLANK = 2 ^ 6                    'Blank characters
    C1_XDIGIT = 2 ^ 7                   'Hexadecimal digits
    C1_ALPHA = 2 ^ 8                    'Any linguistic character
    C1_DEFINED = 2 ^ 9                  'Defined, but not one of the other C1_* types
End Enum

'String Searching: _
 --------------------------------------------------------------------------------------
'Find the first occurance of a string in another [case-sensitive] _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb773436%28v=vs.85%29.aspx>
Private Declare Function api_StrStr Lib "shlwapi" Alias "StrStrW" ( _
    ByVal HaystackPointer As Long, _
    ByVal NeedlePointer As Long _
) As Long

'Find the first occurance of a string in another [case-insensitive] _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb773439%28v=vs.85%29.aspx>
Private Declare Function api_StrStrI Lib "shlwapi" Alias "StrStrIW" ( _
    ByVal HaystackPointer As Long, _
    ByVal NeedlePointer As Long _
) As Long

'/// CLASS STORAGE ////////////////////////////////////////////////////////////////////

'We'll cache the Windows version here as it's very slow to do this every time we _
 call methods that use alternative APIs for Vista+
Private WindowsVersion As Single

'This is where we'll store our "string" as an array of Unicode (UCS-2) points
Private Data() As Integer
'The buffer and content-length will be a 1-based count of characters (not bytes); _
 UCS-2 uses two bytes per character
Private My_Buffer As Long
Private My_Length As Long
'A cache of `My_Length * 2` as it's used in a lot of places
Private My_LengthB As Long

'You can set the buffer to any size (and if you know how long your string is going to _
 be beforehand, you should), but if you don't leave room for growth then the memory _
 will be reallocated once the string overflows the buffer and this is very slow. _
 The buffer automatically doubles in size each time it fills, under the assumption _
 that big data has bigger changes. Whilst not necessary, these constants are designed _
 to make the user think about just how much memory they're using and what the _
 likelihood is that they will overflow
Public Enum BUFFER_SIZE
    'These are measured in power-of-2 characters, not bytes
     '(so "1K" = 1024 characters = 2048 Bytes)
    BUFFER_1 = 2 ^ 0:       BUFFER_1K = 2 ^ 10:     BUFFER_1M = 2 ^ 20
    BUFFER_2 = 2 ^ 1:       BUFFER_2K = 2 ^ 11:     BUFFER_2M = 2 ^ 21
    BUFFER_4 = 2 ^ 2:       BUFFER_4K = 2 ^ 12:     BUFFER_4M = 2 ^ 22
    BUFFER_8 = 2 ^ 3:       BUFFER_8K = 2 ^ 13:     BUFFER_8M = 2 ^ 23
    BUFFER_16 = 2 ^ 4:      BUFFER_16K = 2 ^ 14:    BUFFER_16M = 2 ^ 24
    BUFFER_32 = 2 ^ 5:      BUFFER_32K = 2 ^ 15:    BUFFER_32M = 2 ^ 25
    BUFFER_64 = 2 ^ 6:      BUFFER_64K = 2 ^ 16:    BUFFER_64M = 2 ^ 26
    BUFFER_128 = 2 ^ 7:     BUFFER_128K = 2 ^ 17:   BUFFER_128M = 2 ^ 27
    BUFFER_256 = 2 ^ 8:     BUFFER_256K = 2 ^ 18:   BUFFER_256M = 2 ^ 28
    BUFFER_512 = 2 ^ 9:     BUFFER_512K = 2 ^ 19:   BUFFER_512M = 2 ^ 29
    
    'The maximum number of characters allowed in a string is 1,073,741,824. This is
     'because Unicode (UCS-2) uses 2-bytes per character and since VB6 uses signed
     'Longs, it goes negative above 2 billion, limiting us to 1 billion byte pairs.
    BUFFER_1G = 2 ^ 30:     BUFFER_MAX = BUFFER_1G
    
    'Without the buffer pre-set, it will default to 128 chars which is enough room
     'for growth for most string handling and not too memory intensive (256 bytes)
    DEFAULT_BUFFER = BUFFER_128
End Enum

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize _
 ======================================================================================
Private Sub Class_Initialize()
    'Get the Windows version as Windows Vista and above use alternative APIs for _
     locale-dependent methods (such as `Compare`)
    Dim VersionInfo As OSVERSIONINFO
    Let VersionInfo.SizeOfMe = Len(VersionInfo)
    'NOTE: If the app is in compatibility mode, this will return the compatible _
     Windows version, not the actual version. Windows 8.1 and above also cause this _
     API to always return as Windows 8.0 unless the app opts-in to correct version _
     info via manifest, but this isn't an issue to bluString as we are only concerned _
     with testing for Windows Vista+
    If api_GetVersionEx(VersionInfo) = API_TRUE Then
        '`Val`, rather than `CSng` is used here due to locale differences; _
         `CSng("6.1")` will fail in locales that use comma as decimal separator
        Let WindowsVersion = Val( _
            VersionInfo.MajorVersion & "." & VersionInfo.MinorVersion _
        )
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

'Append : Add another string to the end of the current string _
 ======================================================================================
'Text           | The String to be appended
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Append( _
    ByRef Text As String _
) As bluString
    'Allow method chaining
    Set Append = Me
    'Appending a null-string will do nothing
    If LenB(Text) = 0 Then Exit Function
    
    'The current length will be the starting position of the appended string
    Dim Start As Long: Let Start = My_Length
    'Increase the length accordingly
    Let Me.Length = My_Length + Len(Text)
    'Copy the new string in
    Call UnsafeCopyString(Text, Start)
End Function

'Buffer : Number of characters to reserve in memory for the string _
 ======================================================================================
Public Property Get Buffer() As BUFFER_SIZE: Let Buffer = My_Buffer: End Property
Public Property Let Buffer( _
    ByRef Length As BUFFER_SIZE _
)
    'The buffer cannot be set to zero or a negative number. _
     We won't bother throwing an error as it won't affect behaviour
    If Length <= 0 Then
        Debug.Print "[bluString] WARNING: Cannot set Buffer to 0 or negative number"
        Exit Property
    'Also you cannot go over the maximum buffer size
    ElseIf Length > BUFFER_MAX Then
        'Raise VB's "Overflow" error
        Err.Raise (6): Exit Property
    End If
    
    'Store the new desired buffer length (1-based)
    Let My_Buffer = Length
    'If the buffer is cut shorter than the current content, we need to adjust it too
    If My_Buffer < My_Length Then
        Let My_Length = My_Buffer
        Let My_LengthB = My_Length * 2
    End If
    'Resize the data store. Note that though `My_Buffer` is 1-based we abuse this _
     to provide room for a null-terminator at the end of the string
    ReDim Preserve Data(0 To My_Buffer) As Integer
    'Ensure it's null-terminated!
    Let Data(My_Buffer) = 0
End Property

'Char : The Unicode (UCS-2) point of the character at the provided character index _
 ======================================================================================
Public Property Get Char( _
    ByVal Index As Long _
) As Long
    'NOTE: Though `My_Length` is 1-based, but `Index` is 0-based, we allow the user _
           to access the spare element at the end as it will be the null terminator, _
           which the end-user may wish to check for (blind-looping, for example)
           
    'If a character index is provided outside of the length of the string, error _
     (this also prevents using an index bigger than the maximum buffer size as it's _
      not possible to set `My_Length` above `BUFFER_MAX`)
    If Index > My_Length Then Err.Raise (9): Exit Property
           
    'If the position is negative, get the character code 'x' number of characters _
     before the end of our string. That is, -1 will be the second to last character
    If Index < 0 Then
        'The negative position cannot be more than the length of our string
        If Abs(Index) > My_Length Then Err.Raise (9): Exit Property
        'Calculate the new position from the end of our string
        Let Char = (Data(My_Length + Index) And &HFFFF&)
    Else
        'Return the character value in question _
         The user can use `And &HFF&` to get ASCII
        Let Char = (Data(Index) And &HFFFF&)
    End If
End Property
'--------------------------------------------------------------------------------------
Public Property Let Char( _
    ByVal Index As Long, _
    ByVal Char As Long _
)
    'If a character index is provided outside of the length of the string, error. _
     (note that, compared to the getter, you may not write to the null terminator)
    If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Property
    'Validate the character code. It could be negative if it came from a VB _
     signed-Integer, so this is converted to the positive equivilent
    If Not ValidateChar(Char) Then Err.Raise (9): Exit Property
    
    'VB6 can only do signed Integers, so Unicode points 32'768+ will error unless _
     we convert them to their signed equivilent to stuff them into two bytes
    If Char >= 32768 _
        Then Let Data(Index) = (Char And &HFFFF&) - 65536 _
        Else Let Data(Index) = (Char And &HFFFF&)
End Property

'CharInsert : Inserts a single Unicode character code _
 ======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the _
                | Index is calculated from the end of the string working backwards
'Char           | A Unicode character code (0 to 65535)
'Returns        | This class object (method chaining)
'======================================================================================
Public Function CharInsert( _
    ByVal Index As Long, _
    ByVal Char As Long _
) As bluString
    'Allow method chaining
    Set CharInsert = Me
    
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Function
    'Validate the character code. It could be negative if it came from a VB _
     signed-Integer, so this is converted to the positive equivilent
    If Not ValidateChar(Char) Then Err.Raise (9): Exit Function
    
    'Increase the string length by 1 to make room
    Let Me.Length = My_Length + 1
    'Shift the upper portion of our string up
    Call UnsafeCopy(Index, Index + 1, My_Length - Index - 1)
    'VB6 can only do signed Integers, so Unicode points 32'768+ will error unless _
         we convert them to their signed equivilent to stuff them into two bytes
    If Char >= 32768 _
        Then Let Data(Index) = (Char And &HFFFF&) - 65536 _
        Else Let Data(Index) = (Char And &HFFFF&)
End Function

'CharProperties : Returns the Unicode properties for a character in the string _
 ======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the _
                | Index is calculated from the end of the string working backwards
'Returns        | A bit battern that describes the Unicode properties of the character _
                | as defined in the C1 enum
'======================================================================================
Public Function CharProperties( _
    ByVal Index As Long _
) As C1
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Function
    'Use the API to fill the return value with the character properties
    Call api_GetStringTypeW( _
        CT_CTYPE1, UnsafePointer(Index), 1, _
        VarPtr(CharProperties) _
    )
End Function

'CharPull : Removes the last character of the string and returns the character code _
 ======================================================================================
'Returns        | A Unicode character point (0 to 65535)
'======================================================================================
Public Function CharPull() As Long
    'If there are no characters, just return null
    If My_Length = 0 Then Exit Function
    'Return the character value in question _
     The user can use `And &HFF&` to get ASCII
    Let CharPull = (Data(My_Length - 1) And &HFFFF&)
    'Reduce the size of the string
    Let Me.Length = My_Length - 1
End Function

'CharPush : Appends a single character code to the end of the string _
 ======================================================================================
'Char           | Unicode character point (0 to 65535)
'Returns        | This class object (method chaining)
'======================================================================================
Public Function CharPush( _
    ByVal Char As Long _
) As bluString
    'Allow method chaining
    Set CharPush = Me
    'Validate the character code. It could be negative if it came from a VB _
     signed-Integer, so this is converted to the positive equivilent
    If Not ValidateChar(Char) Then Err.Raise (9): Exit Function
    
    'Increase the length of our string by one character _
     (the buffer will manage any necessary growth)
    Let Me.Length = My_Length + 1
    'Set the character code of the last character
    'VB6 can only do signed Integers, so Unicode points 32'768+ will error unless _
         we convert them to their signed equivilent to stuff them into two bytes
    If Char >= 32768 _
        Then Let Data(My_Length - 1) = (Char And &HFFFF&) - 65536 _
        Else Let Data(My_Length - 1) = (Char And &HFFFF&)
End Function

'CharRemove : Removes a single character from the string and returns its code _
 ======================================================================================
'[Index]        | Position within string (0-based). Can be negative, in which case the _
                | Index is calculated from the end of the string working backwards
'Returns        | A Unicode character point (0 to 65535)
'======================================================================================
Public Function CharRemove( _
    Optional ByVal Index As Long = 0 _
) As Long
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Function
    
    'Return the character value
    Let CharRemove = Data(Index)
    
    'If removing just the last character, simply shorten the string
    If Index = My_Length - 1 Then
        Let Me.Length = My_Length - 1
    Else
        'Shift the remainder of the string down a character
        Call UnsafeCopy(Index + 1, Index, My_Length - Index - 1)
        'Reduce the overal length of the string by one
        Let Me.Length = My_Length - 1
    End If
End Function

'Clear : Wipe the stored string, fill the buffer with nulls _
 ======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Clear() As bluString
    'Allow method chaining
    Set Clear = Me
    'You can't clear a null-string any more than it is!
    If My_Length > 0 Then
        'This does not change the size of the buffer, should it have increased. _
         If you intend to reuse a large string for a small string, create a new _
         bluString instance instead or call the `ReduceBuffer` method
        Call api_RtlZeroMemory(UnsafePointer, My_LengthB)
        'Mark the string as being null-length
        Let My_Length = 0: Let My_LengthB = 0
    End If
End Function

'Clone : Duplicate our string (or subportion) into another bluString instance _
 ======================================================================================
'[Index]        | Position within string (0-based). Can be negative, in which case the _
                | Index is calculated from the end of the string working backwards
'[Length]       | Number of characters to copy, can be 0 which will imply the _
                | remainder of the string, from the Index
'Returns        | A new bluString instance
'======================================================================================
Public Function Clone( _
    Optional ByVal Index As Long = 0, _
    Optional ByVal Length As Long = 0 _
) As bluString
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Function
    'Desired length cannot be negative
    If Length < 0 Then Err.Raise (9): Exit Function
    'The optional length parameter defaults to zero, _
     representative of the full string
    If Length = 0 Then Let Length = My_Length - Index
    'The length cannot be longer than the remainder of the string
    If (My_Length - Index - Length) < 0 Then Err.Raise (9): Exit Function
    
    'Instantiate the return object
    Set Clone = New bluString
    'Set the destination size
    Let Clone.Length = Length
    'Copy the text across (if there is any)
    If My_Length > 0 Then Call api_RtlMoveMemory( _
        Clone.Pointer(Index), UnsafePointer, Length * 2 _
    )
End Function

'Equals : A fast method to check if one string equals another _
 ======================================================================================
'Text           | VB String to compare our string against
'CompareMethod  | Whether to use case-sensitive or case-insensitive comparison
'Returns        | True if strings are equal, False if not
'======================================================================================
Public Function Equals( _
             ByRef Text As String, _
    Optional ByRef CompareMethod As VBA.VbCompareMethod = vbTextCompare _
) As Boolean
    'Handle null strings: _
     Two null strings are considered equal, otherwise result is inequal
    If My_Length = 0 Then
        Let Equals = (LenB(Text) = 0)
        Exit Function
    ElseIf LenB(Text) = 0 Then
        Let Equals = (My_Length = 0)
        Exit Function
    End If
    
    'The APIs used default to binary search, with a flag for ignoring case
    Dim Flags As CompareString_Flags
    If CompareMethod = vbTextCompare Then Let Flags = LINGUISTIC_IGNORECASE
    
    'Check for Windows Vista and above:
    If WindowsVersion >= 6 Then
        'NOTE: The use of null on the first parameter causes Windows to use the _
         user 's current locale, so we don't need to manually fetch the locale name
        Let Equals = (api_CompareStringEx( _
            0&, Flags, _
            StrPtr(Text), -1, _
            UnsafePointer, -1, _
            0&, 0&, 0& _
        ) = CSTR_EQUAL)
    Else
        'Fall back to the XP-compatible API
        Let Equals = (api_CompareString( _
            api_GetThreadLocale(), Flags, _
            StrPtr(Text), -1, _
            UnsafePointer, -1 _
        ) = CSTR_EQUAL)
    End If
End Function

'Format : Replace markers in the text with the values provided _
 ======================================================================================
'Arguments ...  |
'Returns        | A new bluString instance, formatted according to our string
'======================================================================================
Public Function Format( _
    ParamArray Params() As Variant _
) As bluString
    'Instantiate the return object. We clone this string first since if there are _
     no markers in the string, this function will still return the original string
    Set Format = Me.Clone
    
    Dim First As Long, Last As Long
    Let First = LBound(Params): Let Last = UBound(Params)
    
    'If no parameters are provided!
    If Last < First Then Exit Function
    
    'For non-String parameters, we will do the String conversion, but we will need _
     somewhere to store them so that we can get at their pointer
    'There are APIs for doing String conversion <vbforums.com/showthread.php?736407> _
     but CStr is probably using this behind the scenes anyway. Also, rolling your _
     own String conversion is actually super difficult _
     <serpentine.com/blog/2011/06/29/here-be-dragons>
    Dim Strings() As String
    ReDim Strings(First To Last) As String
    'For some types we won't need to convert to String and can use the already _
     existing pointer (for example, working with bluStrings)
    Dim Pointers() As Long
    ReDim Pointers(First To Last) As Long
    
    'Walk the params provided
    Dim Index As Long
    For Index = First To Last
        'Check the type of the parameter (Long, String, Double &c.)
        Dim ParamType As VBA.VbVarType
        Let ParamType = VarType(Params(Index))
        
        Select Case ParamType
            Case vbString
                'We don't want to duplicate existing Strings, we just need the _
                 existing pointers. (`StrPtr` will peek correctly into the Variant)
                 Let Pointers(Index) = StrPtr(Params(Index))
                 
            Case vbObject
                'If this is a bluString, use its string without cloning it. For _
                 other objects, we can't cast to String. Raise "Type Mismatch" error
                If TypeOf Param Is bluString _
                    Then Let Pointers(Index) = Param.Pointer _
                    Else Err.Raise (13): GoTo Finish
            
            'These types can be converted to a String
            Case vbBoolean, vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, _
                 vbEmpty, vbError, vbInteger, vbLong, vbSingle
                Let Strings(Index) = CStr(Params(Index))
                Let Pointers(Index) = StrPtr(Strings(Index))
                
            Case Else
                'All other data types can't be co-erced
                Err.Raise (13): GoTo Finish
        End Select
    Next Index
    
    '----------------------------------------------------------------------------------
    
    'TODO: Manual search and replace (using pointers)
    
    '----------------------------------------------------------------------------------
Finish:
    'And clean up
    Erase Strings, Pointers
End Function

'Insert : Insert some text in the middle of the existing string _
 ======================================================================================
'Text           | VB String to insert into our string
'Index          | Position within our string (0-based). Can be negative, in which case _
                | the Index is calculated from the end of the string working backwards
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Insert( _
    ByRef Text As String, _
    ByVal Index As Long _
) As bluString
    'Allow method chaining
    Set Insert = Me
    'Inserting a null-string won't do anything
    If LenB(Text) = 0 Then Exit Function
    
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Function
    
    'Increase the length of our string to fit the inserted string _
     (the buffer will be increased automatically if necessary)
    Let Me.Length = My_Length + Len(Text)
    'Shift the upper portion of our string up to make room for the inserted string
    Call UnsafeCopy(Index, Index + Len(Text), My_Length - Index)
    'Insert the new string
    Call UnsafeCopyString(Text, Index)
End Function

'IsNull _
 ======================================================================================
'Returns        | True if our string is Null (length = 0), False otherwise
'======================================================================================
Public Property Get IsNull() As Boolean
    'This is intended for readability, since you can obviously just check Length = 0
    Let IsNull = (My_Length = 0)
End Property

'Join : Appends another bluString to the end of this one _
 ======================================================================================
'InputString    | Another bluString instance to take text from
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Join( _
    ByRef InputString As bluString _
) As bluString
    'Allow method chaining
    Set Join = Me
    'Appending a null-string will do nothing
    If InputString.Length = 0 Then Exit Function
    
    'The current length will be the starting position of the appended string
    Dim Start As Long: Let Start = My_Length
    'Increase the length accordingly
    Let Me.Length = My_Length + InputString.Length
    'Copy the new string in
    Call api_RtlMoveMemory( _
        UnsafePointer(Start), InputString.Pointer, InputString.Length * 2 _
    )
End Function

'LCase : Convert our string to lowercase _
 ======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function LCase() As bluString
    'Allow method chaining
    Set LCase = Me
    'Nullstring won't change!
    If My_Length = 0 Then Exit Function
    
    'Check for Windows Vista and above:
    If WindowsVersion >= 6 Then
        'NOTE: The use of null on the first parameter causes Windows to use the _
         user's current locale, so we don't need to manually fetch the locale name
        Call api_LCMapStringEx( _
            0&, LCMAP_LOWERCASE Or LCMAP_LINGUISTIC_CASING, _
            UnsafePointer, My_LengthB, _
            UnsafePointer, My_LengthB, _
            0&, 0&, 0& _
        )
    Else
        'Use the Windows XP compatible API
        Call api_LCMapString( _
            api_GetThreadLocale(), LCMAP_LOWERCASE Or LCMAP_LINGUISTIC_CASING, _
            UnsafePointer, My_LengthB, _
            UnsafePointer, My_LengthB _
        )
    End If
End Function

'Left : Cut the string short to a given number of characters _
 ======================================================================================
'Length         | Number of characters to keep of the string. If negative, this will _
                | be calculated backwards from the end of the string
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Left( _
    ByRef Length As Long _
) As bluString
    'Allow method chaining
    Set Left = Me
    
    'Shortcut if trimming to zero
    If Length = 0 Then Call Me.Clear: Exit Function
    'Trimming to the same length does nothing!
    If Length = My_Length Then Exit Function
    'You can't increase the length with this!
    If Length > My_Length Then Err.Raise (9): Exit Function
    
    'If the position is negative, work out 'x' number of characters before the end of _
     our string. That is, -1 will be the second to last character
    If Length < 0 Then
        'The negative position cannot be more than the length of our string
        If Abs(Length) > My_Length Then Err.Raise (9): Exit Function
        'Calculate the new position from the end of our string
        Let Me.Length = My_Length + Length
    Else
        'Do the reduction
        Let Me.Length = Length
    End If
End Function

'Length : Length of the actual stored text (1-based), not the buffer _
 ======================================================================================
Public Property Get Length() As Long: Let Length = My_Length: End Property
Public Property Let Length( _
    ByRef NewLength As Long _
)
    'Cannot obviously set the length to a negative number
    If NewLength < 0 Then Err.Raise (9): Exit Property
    'Can't set the length greater than the maximum allowed size _
     (VB Error #6 is "Overflow")
    If NewLength > BUFFER_MAX Then Err.Raise (6): Exit Property

    'Store the new desired content length (1-based)
    Let My_Length = NewLength: Let My_LengthB = My_Length * 2
    'If the buffer has not been initialised, start with the default length
    If My_Buffer = 0 Then Let My_Buffer = DEFAULT_BUFFER
    
TestBuffer:
    'If this increases the data too large...
    If My_Length > My_Buffer Then
        'If the buffer is already at maximum then we've hit the absolute limit _
         (VB Error #14 is "Out of String Space")
        If My_Buffer = BUFFER_MAX Then Err.Raise (14): Exit Property
        'In the case of custom buffer sizes, doubling the buffer might go over the _
         limit, but there may still be enough room for the new length if we just stop _
         at the buffer limit
        If (My_Buffer * 2) > BUFFER_MAX Then
            Let My_Buffer = BUFFER_MAX
        Else
            'Double the length of the buffer
            Let My_Buffer = My_Buffer * 2
        End If
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
'Index          | Position within string (0-based). Can be negative, in which case _
                | the Index is calculated from the end of the string working backwards
'[Length]       | Number of characters to keep, can be 0 which will imply the _
                | remainder of the string, from the Index
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Mid( _
             ByVal Index As Long, _
    Optional ByVal Length As Long = 0 _
) As bluString
    'Allow method chaining
    Set Mid = Me
    
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Function
    'Desired length cannot be negative
    If Length < 0 Then Err.Raise (9): Exit Function
    'The optional length parameter defaults to zero, _
     representative of the full string
    If Length = 0 Then Let Length = My_Length - Index
    'The length cannot be longer than the remainder of the string
    If (My_Length - Index - Length) < 0 Then Err.Raise (9): Exit Function
    
    'Copy the sub-string to the beginning of our string
    Call UnsafeCopy(Index, 0, Length)
    'Reduce the string to the new sub-string length
    Let Me.Length = Length
End Function

''Normalize : Fix inconsistent use of Unicode encoding _
' ======================================================================================
'Public Function Normalize() As bluString
'    'Allow method chaining
'    Set Normalize = Me
'   'NOTE: Vista+ only _
     http://msdn.microsoft.com/en-gb/library/windows/desktop/dd319093(v=vs.85).aspx
'End Function

'Pointer : Allow access to the "string" pointer for external copying _
 ======================================================================================
'[Index]        | Position within our string (0-based)
'Returns        | Memory pointer to the character index within our string
'======================================================================================
Public Property Get Pointer( _
    Optional ByVal Index As Long = 0 _
) As Long
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Property
    
    'If a null string, return a null pointer (for potential API use)
    If My_Length = 0 _
        Then Let Pointer = 0 _
        Else Let Pointer = UnsafePointer(Index)
End Property

'Prepend : Add our string to the end of the one provided _
 ======================================================================================
'Text           | VB String to prepend to our string
'Returns        | This class object (method chaining)
'======================================================================================
'And no, "prefix" is not the right word for this!
Public Function Prepend( _
    ByRef Text As String _
) As bluString
    'Allow method chaining
    Set Prepend = Me
    'Prepending a null-string will do nothing
    If LenB(Text) = 0 Then Exit Function
    
    'The length of the new String will be the new starting position of our string
    Dim Start As Long: Let Start = Len(Text)
    'Increase the length accordingly
    Let Me.Length = My_Length + Start
    'Shift our string up
    Call UnsafeCopy(0, Start, My_Length)
    'Copy the new string in
    Call UnsafeCopyString(Text, 0)
End Function

''ReadFile : Import our string from a file _
' ======================================================================================
'Public Function ReadFile( _
'    ByRef FilePath As String _
') As bluString
'    'Allow method chaining
'    Set ReadFile = Me
'    'TODO
'   'http://stackoverflow.com/questions/15809081/encoding-of-text-files-in-vb-6-0
'End Function

'ReduceBuffer : Free some memory by reducing the buffer to the string length _
 ======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
'Of course you could just use `bluString.Buffer = bluString.Length`, but this is _
 intended to be more obvious and safer for potential future changes / expansion.
'It's strongly advised NOT to expand the string after you've reduced the buffer
Public Function ReduceBuffer() As bluString
    'Allow method chaining
    Set ReduceBuffer = Me
    'Reduce the buffer down to the size of the string
    Let Me.Length = My_Buffer
End Function

'Remove : Cut some text out of the middle of the existing string _
 ======================================================================================
'Index          | Position within our string (0-based). Can be negative, in which case _
                | the Index is calculated from the end of the string working backwards
'[Length]       | Number of characters to remove, 0 = remainder of string from Index
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Remove( _
             ByVal Index As Long, _
    Optional ByVal Length As Long = 0 _
) As bluString
    'Allow method chaining
    Set Remove = Me
    
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Function
    
    'If you end up removing the whole string, make it a null-string
    If (Index = 0) And (Length = My_Length - 1) Then Call Me.Clear: Exit Function
    
    'Desired length cannot be negative
    If Length < 0 Then Err.Raise (9): Exit Function
    'The optional length parameter defaults to zero, _
     representative of the full string
    If Length = 0 Then Let Length = My_Length - Index
    'The length cannot be longer than the remainder of the string
    If My_Length - Index - Length < 0 Then Err.Raise (9): Exit Function
    
    'Shift the remainder down into the removed area
    Call UnsafeCopy(Index + Length, Index, Length)
    'Update the string length
    Let Me.Length = My_Length - Length
End Function

'Replace : Search for a sub-string and replace it with another _
 ======================================================================================
'SearchFor      | Text to find within our string
'[ReplaceWith]  | What to replace the found text with, can be null-string
'[ReturnCount]  | If provided, your own variable will be set with the number of _
                | replacements done
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Replace( _
             ByRef SearchFor As String, _
    Optional ByRef ReplaceWith As String = vbNullString, _
    Optional ByRef ReturnCount As Long = 0 _
) As bluString
    'Allow method chaining
    Set Replace = Me
    'Ensure that the count will return 0 if nothing is found
    Let ReturnCount = 0
    
    'You can't search in a null string
    If My_Length = 0 Then Exit Function
    'What you're searching for cannot be longer than our string
    If Len(SearchFor) > My_Length Then Exit Function
    
    'If the replacement string is shorter than the search string, then we'll need _
     to remove parts of our string and vice-versa. Work out the impact the _
     replacement will have on the length of our string
    Dim Difference As Long
    Let Difference = Len(ReplaceWith) - Len(SearchFor)
    
    'The API returns a pointer in the string, not an index, _
     we'll use this to start searching from where we left off!
    Dim StartPointer As Long
    Let StartPointer = UnsafePointer()
    
    'Some caches before we start looping
    Dim StrPtr_SearchFor As Long:   Let StrPtr_SearchFor = StrPtr(SearchFor)
    Dim StrPtr_ReplaceWith As Long: Let StrPtr_ReplaceWith = StrPtr(ReplaceWith)
    Dim LenB_SearchFor As Long:     Let LenB_SearchFor = LenB(SearchFor)
    Dim LenB_ReplaceWith As Long:   Let LenB_ReplaceWith = LenB(ReplaceWith)
    Dim LastPointer As Long:        Let LastPointer = UnsafePointer() + My_LengthB
    
Continue:
    'Search for the sub-string
    Dim ResultPointer As Long
    Let ResultPointer = api_StrStr(StartPointer, StrPtr_SearchFor)
    
    'If the returned pointer is null, the sub-string was not found
    If ResultPointer = 0 Then Exit Function
    
    'If both the search string and the replacement are the same length, _
     we don't need to move parts of our string about
    If Difference = 0 Then
        'Replace the found string
        Call api_RtlMoveMemory( _
            ResultPointer, StrPtr_ReplaceWith, LenB_ReplaceWith _
        )
    Else
        'If the replacement string is longer than the search string, _
         our string will get longer overall
        If Difference > 0 Then Let Me.Length = My_Length + Difference
        
        'Calculate the remainder portion of our string that will be moved
        Dim Remainder As Long
        Let Remainder = LastPointer - (StartPointer + LenB_SearchFor)
        
        'Shift up or down the remainder portion of our string
        Call api_RtlMoveMemory( _
            ResultPointer + LenB_ReplaceWith, _
            ResultPointer + LenB_SearchFor, _
            Remainder _
        )
        'If the replacement string is shorter than the search string, _
         our string will get shorter overall
        If Difference < 0 Then Let Me.Length = My_Length + Difference
        
        'Copy the replacement string in
        Call api_RtlMoveMemory( _
            ResultPointer, StrPtr_ReplaceWith, LenB_ReplaceWith _
        )
    End If
    
    'We don't want an infinite loop, so move our search pointer over the replacement
    Let StartPointer = StartPointer + LenB_ReplaceWith
    GoTo Continue
End Function

'Right : Keep only the given right-most number of characters _
 ======================================================================================
'Length         | Number of characters to keep working backwards from the right
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Right( _
    ByRef Length As Long _
) As bluString
    'Allow method chaining
    Set Right = Me
    
    'Cannot trim longer than the string!
    If Length > My_Length Then Err.Raise (9): Exit Function
    'Keeping all characters does nothing!
    If Length = My_Length Then Exit Function
    'Removing all characters creates a null-string!
    If Length = 0 Then Call Me.Clear: Exit Function
    'Negative numbers are not supported, just use `Left` to keep x number of chars
    If Length < 0 Then Err.Raise (9): Exit Function
    'Shift the rightmost portion of the string down the beginning
    Call UnsafeCopy(My_Length - Length, 0, Length)
    'Cut the string short. The buffer will be left at its old length, _
     if you know the string won't expand any more you can free the memory up _
     by calling the `ReduceBuffer` method
    Let Me.Length = Length
End Function

'Search : Find a string within our string _
 ======================================================================================
'Public Function Search() As Long
'    '
'End Function

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
        Call api_RtlMoveMemory(VarPtr(CastStrArr(0)), UnsafePointer, My_LengthB)
    End If
End Property
'--------------------------------------------------------------------------------------
Public Property Let Text( _
    ByRef NewString As String _
)   'If a null-string is provided then set ours to a null string
    If LenB(NewString) = 0 Then
        Call Me.Clear
    Else
        'Set the new length of our String. The buffer will be increased appropriately
        Let Me.Length = Len(NewString)
        'Copy the String data over
        Call api_RtlMoveMemory(UnsafePointer, StrPtr(NewString), LenB(NewString))
    End If
End Property

''Trim : Strip whitespace characters from the beginning and end _
' ======================================================================================
''TODO: We may want a look-up table of all whitespace chars in UCS-2
'Public Function Trim() As bluString
'    'Allow method chaining
'    Set Trim = Me
'    'TODO
'End Function

''TrimLeft : Strip whitespace characters from the beginning _
' ======================================================================================
'Public Function TrimLeft() As bluString
'    'Allow method chaining
'    Set TrimLeft = Me
'    'TODO
'End Function

''TrimRight : Strip whitespace characters from the end _
' ======================================================================================
'Public Function TrimRight() As bluString
'    'Allow method chaining
'    Set TrimRight = Me
'    'TODO
'End Function

'UCase : Convert our string to UPPERCASE _
 ======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function UCase() As bluString
    'Allow method chaining
    Set UCase = Me
    'Nullstring won't change!
    If My_Length = 0 Then Exit Function
    
    'Check for Windows Vista and above:
    If WindowsVersion >= 6 Then
        'NOTE: The use of null on the first parameter causes Windows to use the _
         user's current locale, so we don't need to manually fetch the locale name
        Call api_LCMapStringEx( _
            0&, LCMAP_UPPERCASE Or LCMAP_LINGUISTIC_CASING, _
            UnsafePointer, My_LengthB, _
            UnsafePointer, My_LengthB, _
            0&, 0&, 0& _
        )
    Else
        'Use the Windows XP compatible API
        Call api_LCMapString( _
            api_GetThreadLocale(), LCMAP_UPPERCASE Or LCMAP_LINGUISTIC_CASING, _
            UnsafePointer, My_LengthB, _
            UnsafePointer, My_LengthB _
        )
    End If
End Function

''WriteFile : Export our string to a file _
' ======================================================================================
'Public Function WriteFile( _
'    ByRef FilePath As String _
') As bluString
'    'Allow method chaining
'    Set WriteFile = Me
'    'TODO
'End Function

'/// PRIVATE INTERFACE ////////////////////////////////////////////////////////////////

'UnsafePointer : Gets a pointer to the string, without validation -- internal use _
 ======================================================================================
'[Index]        | Position within our string (0-based)
'Returns        | Memory pointer to the character index within our string
'======================================================================================
Private Property Get UnsafePointer( _
    Optional ByRef Index As Long = 0 _
) As Long
    'Return a pointer to the given character index
    Let UnsafePointer = VarPtr(Data(Index))
End Property

'UnsafeCopy : Move a portion of our string to another place in our string _
 ======================================================================================
'IndexFrom      | Position within string (0-based) to begin copying from
'IndexTo        | As above, but where to copy to
'Length         | Number of characters (not Bytes) to copy
'======================================================================================
Private Sub UnsafeCopy( _
    ByRef IndexFrom As Long, _
    ByRef IndexTo As Long, _
    ByRef Length As Long _
)   'No validation! internal use ONLY!
    Call api_RtlMoveMemory( _
        UnsafePointer(IndexTo), UnsafePointer(IndexFrom), Length * 2 _
    )
End Sub

'UnsafeCopyString : Copies a VB String into our string _
 ======================================================================================
'Text           | VB String to copy into our string
'Index          | Position within our string (0-based) to begin copying
'======================================================================================
Private Sub UnsafeCopyString( _
    ByRef Text As String, _
    ByRef Index As Long _
)   'No validation! internal use ONLY!
    '(This will overwrite anything already in place, so it's assumed you've already _
      handled resizing the string and copying things out)
    Call api_RtlMoveMemory( _
        UnsafePointer(Index), StrPtr(Text), LenB(Text) _
    )
End Sub

'ValidateChar : Validates a Unicode character code _
 ======================================================================================
'Char           | A Unicode character code (0 to 65535) or a negative number _
                | representing the signed-Integer equivilent of codes 32768 to 65535
'Returns        | False for an invalid Char, True for a valid one
'======================================================================================
Private Function ValidateChar( _
    ByRef Char As Long _
) As Boolean
    'If the character code is negative, it might be from a signed-Integer
    If Char < 0 Then
        'If it's in range, convert it to a positive number
        If Char >= -32768 Then
            Let Char = Char + 65535
            Let ValidateChar = True
        End If
    Else
        'Maximum of 65535
        If Char < 65536 Then Let ValidateChar = True
    End If
End Function

'ValidateIndex : Validates an index for the string _
 ======================================================================================
'Index          | Position within string (0-based)
'Returns        | False for an invalid Index, True for a valid one
'======================================================================================
Private Function ValidateIndex( _
    ByRef Index As Long _
) As Boolean
    'Index cannot be less than zero
    If (Index >= 0) Then
        'Index cannot be beyond the end of the string _
        (note that the index is 0-based and the length is 1-based)
        If (Index < My_Length) Then Let ValidateIndex = True
    End If
End Function

'ValidateIndexAllowNegative : A negative index gives 'x' number of chars from the end _
 ======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the _
                | Index is updated to be from the end of the string working backwards
'Returns        | False for an invalid Index, True for a valid one
'======================================================================================
Private Function ValidateIndexAllowNegative( _
    ByRef Index As Long _
) As Boolean
    'Is the Index negative?
    If Index < 0 Then
        'The negative position cannot be more than the length of our string
        If Abs(Index) < My_Length Then
            'Update the Index to work backwards from the end of the string
            Let Index = My_Length + Index
            Let ValidateIndexAllowNegative = True
        End If
    Else
        'Index cannot be beyond the end of the string _
        (note that the index is 0-based and the length is 1-based)
        If (Index < My_Length) Then Let ValidateIndexAllowNegative = True
    End If
End Function
