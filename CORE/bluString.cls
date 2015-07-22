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
'blu : A Modern Metro-esque graphical toolkit; Copyright (C) Kroc Camen, 2013-15
'Licenced under a Creative Commons 3.0 Attribution Licence
'--You may use and modify this code how you see fit as long as you give credit
'======================================================================================
'CLASS :: bluString

'Dependencies       WIN32.tlb, blu > blu_Sys, blu_Strings, blu_FileSystem
'Last Updated       20-JUL-15
'Last Update        *   Added `CharArrayAccess/Unaccess` methods
'                   *   Renamed `Data` array to `Chars`

'Always get the latest version of this file at <github.com/Kroc/blu>. _
 Feedback can be sent to <kroc@camendesign.com>

'--------------------------------------------------------------------------------------

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

'Comparing Strings:
'--------------------------------------------------------------------------------------

'Binary string comparison (case-sensitive) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd317762(v=vs.85).aspx>
Private Declare Function api_CompareStringOrdinal Lib "kernel32" Alias "CompareStringOrdinalW" ( _
    ByVal String1Pointer As Long, _
    ByVal String1Length As Long, _
    ByVal String2Pointer As Long, _
    ByVal String2Length As Long, _
    ByVal IgnoreCase As BOOL _
) As CSTR_

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
    ByVal VersionInfoPointer As Long, _
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

'Unicode character properties:
'--------------------------------------------------------------------------------------
'Returns an array of Unicode character properties for a string _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318119(v=vs.85).aspx>
Private Declare Function api_GetStringType Lib "kernel32" Alias "GetStringTypeW" ( _
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
    
    'Shorthand for "alpha-numeric"
    C1_ALPHANUM = C1_ALPHA Or C1_DIGIT
    'All kinds of blank characters you would want to strip off the ends
    C1_WHITESPACE = C1_SPACE Or C1_BLANK Or C1_CNTRL
    'Visible ("Printable") characters, this includes spaces, tabs &c.
    C1_VISIBLE = C1_SPACE Or C1_PUNCT Or C1_BLANK Or C1_ALPHANUM
End Enum

'String Searching & Replacing:
'--------------------------------------------------------------------------------------
'Find the first occurance of a string in another [case-sensitive] _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb773436(v=vs.85).aspx>
Private Declare Function api_StrStr Lib "shlwapi" Alias "StrStrW" ( _
    ByVal HaystackPointer As Long, _
    ByVal NeedlePointer As Long _
) As Long

'Find the first occurance of a string in another [case-insensitive] _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb773439(v=vs.85).aspx>
Private Declare Function api_StrStrI Lib "shlwapi" Alias "StrStrIW" ( _
    ByVal HaystackPointer As Long, _
    ByVal NeedlePointer As Long _
) As Long

'This type is used during the `Format` method to convert Variants into Strings. _
 For Variants that are already Strings we just use the pointer, but for types that _
 need converting, we can store the String conversion here to get at a safe pointer
Private Type bluString_FormatData
    Text As String                      'Optional; use the pointer of existing strings
    Pointer As Long                     '`StrPtr` of the String, not necessarily above
    ByteCount As Long
End Type

'Counts the length of a string, looking for the first Null _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms647492(v=vs.85).aspx>
Private Declare Function api_lstrlen Lib "kernel32" Alias "lstrlenW" ( _
    ByVal lpString As Long _
) As Long

'/// CLASS ////////////////////////////////////////////////////////////////////////////

'In order to provide external reference to this array as a whole, we need to _
 understand how VB stores Arrays. `VarPtr` typically provides the address of a VB _
 variable in memory. This doesn't work on Arrays due to the way `VarPtr` is declared; _
 `ArrPtr` is provided by the WIN32 TLB and allows you to get the address of an Array.

'However, the Array you are looking for is another castle; the contents of the _
 variable's memory address is simply another memory address. If you follow that you _
 get either "0", for an undefined Array, or *another* address, where the actual Array _
 structure is located

' +-----------+
' | Variable  |     Value is a memory address to:
' +-----------+
'       |
'       v
' +-----------+
' |  Pointer  |     Value is 0 for an unitialised Array, otherwise
' +-----------+     the address of which Array this Variable is bound to:
'       |
'       v
' +-----------+
' | SAFEARRAY |     A SAFEARRAY structure that describes the Array
' +-----------+

'Why so much indirection? It allows VB to clearly differentiate between Arrays that _
 are unintialised (have no dimensions or data) and those that are. It also allows _
 one Variable to change what Array it's pointing at ("Pointer" above); and lastly, _
 we can have multiple Variables point to the same Array! Should any one Variable _
 be destroyed only that Variable will be cleared, not the Array's data

'By manipulating an unitialised Array ("Pointer"=0), we can return an Array from _
 a function that points to our internal string data array, allowing external access _
 to our data, without copying it. However, we need to prevent the user from redimming _
 the array or erasing it entirely (making it uninitalised)

'To do that we add a lock to our array; allowing read/write access to the data, but _
 preventing the user from redimming it. When the external reference goes out of scope, _
 that variable will be cleared, but the lock on our data will prevent it being cleared _
 too! VB will not clear this lock for us, so we need to do this manually via a _
 method that breaks the array reference previously given

'--------------------------------------------------------------------------------------

'This is where we store our "string" as an array of Unicode (UCS-2) points
Private Chars() As Integer

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
    
    'When working with file paths, this is the internal limit for short-paths
    MAX_PATH = 260
    
    'Without the buffer pre-set, it will default to 128 chars which is enough room
     'for growth for most string handling and not too memory intensive (256 bytes)
    DEFAULT_BUFFER = BUFFER_128
End Enum

'Set to 1 once the memory has been reserved, _
 this will allow us to get the pointer to an empty string (for API usage)
Private Reserved As Long

'--------------------------------------------------------------------------------------

'The marker used for formatting strings. As a constant here, it saves VB allocating _
 and deallocating a temporary String everytime we mention it
Private Const FORMAT_MARKER As String = "{$}"

'Case-sensitivity in methods such as `Format` and `LCase`/`UCase` is controlled _
 across the class by a property
Private My_CaseSensitive As Long

'Stuff for doing CRCs:
'--------------------------------------------------------------------------------------
'The CRC function of this class is from "Calculating CRC32 With VB" by Steve McMahon _
 <www.vbaccelerator.com/home/VB/Code/Libraries/CRC32/article.asp> which was _
 derived from from the VB.NET CRC32 algorithm by Paul Caton

'This is the official polynomial used by CRC32 in PKZip. _
 Often the polynomial is shown reversed (04C11DB7)
Private Const CRC_POLYNOMIAL As Long = &HEDB88320

'A lookup table needs to be built for calculating the CRC, but we will delay this _
 until needed as it takes time and adds 1 KB to the memory usage
Private CRCTable() As Long
Private CRCInitialized As Long

'The calculated CRC is cached here, we don't want to have to do the calculation every _
 time the property is accessed (as long as the string hasn't changed)
Private My_CRC As Long
'A flag to mark when the string has changed and the CRC is no longer valid
Private CRCValid As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'NOTE: The buffer is not initialised to begin with _
       (therefore not wasting memory for a string that is not yet populated, _
        but more so that the reserved memory does not have to be reallocated _
        if a String larger than the default buffer is created first)
           
'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    'Free up the reserved memory
    Call Me.Destroy
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Append : Add another string to the end of the current string
'======================================================================================
'Text           | The String to be appended
'---------------¦----------------------------------------------------------------------
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
    Call my_SetLength(My_Length + Len(Text))
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
    
    'Copy the new string in
    Call Win32.MemoryCopyBetween( _
          LengthBytes:=LenB(Text), _
        OutputPointer:=VarPtr(Chars(Start)), _
        SourcePointer:=StrPtr(Text) _
    )
End Function

'Buffer : Number of characters to reserve in memory for the string
'======================================================================================
Public Property Get Buffer() As BUFFER_SIZE
    Let Buffer = My_Buffer
End Property
Public Property Let Buffer( _
    ByRef BufferSize As BUFFER_SIZE _
)
    'No change in buffer?
    If BufferSize = My_Buffer Then Exit Property
    
    'The buffer cannot be set to zero or a negative number. _
     We won't bother throwing an error as it won't affect behaviour
    If BufferSize <= 0 Then
        Debug.Print "[bluString] WARNING: Cannot set Buffer to 0 or negative number"
        Exit Property
        
    'Also you cannot go over the maximum buffer size
    ElseIf BufferSize > BUFFER_MAX Then
        'Raise VB's "Overflow" error
        Err.Raise (6): Exit Property
    End If
    
    'Store the new desired buffer length (1-based)
    Let My_Buffer = BufferSize
    'If the buffer is cut shorter than the current content, we need to adjust it too
    If My_Buffer < My_Length Then
        Let My_Length = My_Buffer
        Let My_LengthB = My_Length * 2
        'Changes to our string will invalidate the CRC
        Let CRCValid = 0
    End If
    
    'Resize the data store. Note that though `My_Buffer` is 1-based we abuse this _
     to provide room for a null-terminator at the end of the string
    'If `CharArrayAccess` has been used the array will be locked and this will fail; _
     the user must release the lock with `CharArrayUnaccess`
    ReDim Preserve Chars(0 To My_Buffer) As Integer
    'Ensure it's null-terminated!
    Let Chars(My_Buffer) = 0
    'Note that the memory has been reserved
    Let Reserved = 1
End Property

'CaseSensitive : Turn case-sensitivity on or off
'======================================================================================
Public Property Get CaseSensitive() As Boolean
    Let CaseSensitive = My_CaseSensitive
End Property
Public Property Let CaseSensitive(ByRef Mode As Boolean)
    Let My_CaseSensitive = Mode
End Property

'Char : The Unicode (UCS-2) point of the character at the provided character index
'======================================================================================
Public Property Get Char( _
    ByVal Index As Long _
) As UCS2
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
        Let Char = (Chars(My_Length + Index) And &HFFFF&)
    Else
        'Return the character value in question _
         The user can use `And &HFF&` to get ASCII
        Let Char = (Chars(Index) And &HFFFF&)
    End If
End Property
Public Property Let Char( _
    ByVal Index As Long, _
    ByVal Char As UCS2 _
)
    'If a character index is provided outside of the length of the string, error. _
     (note that, compared to the getter, you may not write to the null terminator)
    If Not my_ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Property
    'Validate the character code. It could be negative if it came from a VB _
     signed-Integer, so this is converted to the positive equivilent
    If Not my_ValidateChar(Char) Then Err.Raise (9): Exit Property
    
    'VB6 can only do signed Integers, so Unicode points 32'768+ will error unless _
     we convert them to their signed equivilent to stuff them into two bytes
    If Char >= 32768 _
        Then Let Chars(Index) = (Char And &HFFFF&) - 65536 _
        Else Let Chars(Index) = (Char And &HFFFF&)
    
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Property

'CharArrayAccess : Provide an array of the character codes in our string
'======================================================================================
'Returns        | An Integer array diretly referencing our internal string storage.
'               | The array given is locked (cannot be redimmed) and must be unlocked
'               | (call `CharArrayUnaccess`) before this class can be destroyed
'======================================================================================
Public Function CharArrayAccess( _
) As Integer()
    'Set the return array to share the same data as our string's character array _
     (VB cannot return an `Any()` array so this API call is type-specific)
    Let CharArrayAccess = Win32.ArrayRef_Integer(Chars)
    
    'Increase the lock count of the array so that the user cannot redim it. _
     This lock will have to be manually removed before the class can be destroyed
    Call Win32.ArrayLockAdd(Win32.ArrPtr(Chars))
End Function

'CharArrayUnaccess :
'======================================================================================
'CharArray()    | The array previously handed out by `CharArrayAccess`; it will be set
'               | to an uninitalised state and our string will be unlocked
'======================================================================================
Public Sub CharArrayUnaccess( _
    ByRef CharArray() As Integer _
)
    'Is this our array? (Don't let the user unlock any old array!)
    If Win32.ArrPtr(CharArray) <> Win32.ArrPtr(Chars) Then
        'ERROR: "Invalid Procedure Call or Argument"
        Call Err.Raise(5, , "bluString.CharArrayUnaccess: The array passed is not the same as was retrieved from `CharArrayAccess`")
        Exit Sub
    End If
    
    'Remove one from the lock count. If the user has been doing weird things with _
     the array they've been given, it might have some locks remaining
    Call Win32.ArrayLockRemove(Win32.ArrPtr(Chars))
    
    'Detach the array variable from its array structure (`SAFEARRAY`). _
     The passed array will become uninitialised
    Let MemLong(Win32.VarPtr_Array(Chars)) = 0
End Sub

'CharInsert : Inserts a single Unicode character code
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'Char           | A Unicode character code (0 to 65535)
'---------------¦----------------------------------------------------------------------
'Returns        | This class object (method chaining)
'======================================================================================
Public Function CharInsert( _
    ByVal Index As Long, _
    ByVal Char As UCS2 _
) As bluString
    'Allow method chaining
    Set CharInsert = Me
    
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not my_ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Function
    'Validate the character code. It could be negative if it came from a VB _
     signed-Integer, so this is converted to the positive equivilent
    If Not my_ValidateChar(Char) Then Err.Raise (9): Exit Function
    
    'Increase the string length by 1 to make room _
     (this will also invalidate the cached CRC)
    Call my_SetLength(My_Length + 1)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
    
    'Shift the upper portion of our string up
    Call Win32.MemoryCopyWithin( _
        OutputPointer:=VarPtr(Chars(Index + 1)), _
        SourcePointer:=VarPtr(Chars(Index)), _
          LengthBytes:=(My_Length - Index - 1) * 2 _
    )
    
    'VB6 can only do signed Integers, so Unicode points 32'768+ will error unless _
         we convert them to their signed equivilent to stuff them into two bytes
    If Char >= 32768 _
        Then Let Chars(Index) = (Char And &HFFFF&) - 65536 _
        Else Let Chars(Index) = (Char And &HFFFF&)
End Function

'CharIsAlpha : Returns true if the character at the index is a letter
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'[IsASCII]      | If True, only A-Z will validate, accented characters etc. won't
'---------------¦----------------------------------------------------------------------
'Returns        | True if the character is alpha (a-z, but also non-latin characters)
'======================================================================================
Public Property Get CharIsAlpha( _
             ByVal Index As Long, _
    Optional ByRef IsASCII As Boolean = False _
) As Boolean
    'Whilst Unicode recognises a wider range of "alpha" characters, _
     it is sometimes necessary to limit ourselves to just A-Z
    If IsASCII = 0 Then
        'Note that this uses Unicode properties, so covers far more than just A-Z, _
         including accented letters and non-latin alpha letters such as "ß"
        Let CharIsAlpha = (Me.CharProperties(Index) And C1_ALPHA) <> 0
    Else
        'Check the character index is within range, _
         a negative Index will work backwards from the end of the string
        If Not my_ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Property
        'Is the character code within A-Z? (upper and lower case)
        If Chars(Index) >= 97 And Chars(Index) <= 122 Then Let CharIsAlpha = True: Exit Property
        If Chars(Index) >= 65 And Chars(Index) <= 90 Then Let CharIsAlpha = True
    End If
End Property

'CharIsAlphaNumeric : Returns true if the character at the index is a letter or number
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'[IsASCII]      | If True, only A-Z&0-9 will validate, accented characters etc. won't
'---------------¦----------------------------------------------------------------------
'Returns        | True if the character is A-Z, 0-9
'======================================================================================
Public Property Get CharIsAlphaNumeric( _
             ByVal Index As Long, _
    Optional ByRef IsASCII As Boolean = False _
) As Boolean
    'Whilst Unicode recognises a wider range of "alpha-numeric" characters, _
     it is necessary to sometimes limit ourselves to just A-Z & 0-9
    If IsASCII = 0 Then
        'Note that this uses Unicode properties, so covers far more than just A-Z & _
         0-9, including accented letters and non-latin alpha letters such as "ß"
        Let CharIsAlphaNumeric = (Me.CharProperties(Index) And C1_ALPHANUM) <> 0
    Else
        'Check the character index is within range, _
         a negative Index will work backwards from the end of the string
        If Not my_ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Property
        'Is the character code within A-Z? (upper and lower case)
        If Chars(Index) >= 97 And Chars(Index) <= 122 Then Let CharIsAlphaNumeric = True: Exit Property
        If Chars(Index) >= 65 And Chars(Index) <= 90 Then Let CharIsAlphaNumeric = True: Exit Property
        'Is the character code within 0-9?
        If Chars(Index) >= 48 And Chars(Index) <= 57 Then Let CharIsAlphaNumeric = True
    End If
End Property

'CharIsHex : Returns true if the character at the index is a hexadecimal digit
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'---------------¦----------------------------------------------------------------------
'Returns        | True if the character is a hexadecimal digit (0-9 & A-F)
'======================================================================================
Public Property Get CharIsHex( _
    ByVal Index As Long _
) As Boolean
    'We could check ASCII codes, but the API can do this for us
    Let CharIsHex = (Me.CharProperties(Index) And C1_XDIGIT) <> 0
End Property

'CharIsNumeric : Returns true if the character at the index is a number
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'[IsASCII]      | If True, only 0-9 will validate, non-latin/wide numbers etc. won't
'---------------¦----------------------------------------------------------------------
'Returns        | True if the character is a number
'======================================================================================
Public Property Get CharIsNumeric( _
             ByVal Index As Long, _
    Optional ByRef IsASCII As Boolean = False _
) As Boolean
    'Whilst Unicode recognises a wider range of "numeric" characters, _
     it is necessary to sometimes limit ourselves to just 0-9
    If IsASCII = 0 Then
        'Note that this uses Unicode properties, so covers far more than just 0-9, _
         including non-Latin script numerals and "wide" characters
        Let CharIsNumeric = (Me.CharProperties(Index) And C1_DIGIT) <> 0
    Else
        'Is the character code within 0-9?
        If Chars(Index) >= 48 And Chars(Index) <= 57 Then Let CharIsNumeric = True
    End If
End Property

'CharIsLower : Returns true if the character at the index is lower-case
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'---------------¦----------------------------------------------------------------------
'Returns        | True if the character is lower-case
'======================================================================================
Public Property Get CharIsLower( _
    ByVal Index As Long _
) As Boolean
    'Note that this uses Unicode properties, so covers far more than just A-Z
    Let CharIsLower = (Me.CharProperties(Index) And C1_LOWER) <> 0
End Property

'CharIsSpace : Returns true if the character at the index is white-space or invisible
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'---------------¦----------------------------------------------------------------------
'Returns        | True if the character is invisible or white-space
'======================================================================================
Public Property Get CharIsSpace( _
    ByVal Index As Long _
) As Boolean
    'Note that this uses Unicode properties, so covers far more than just ASCII space
    Let CharIsSpace = (Me.CharProperties(Index) And C1_WHITESPACE) <> 0
End Property

'CharIsUpper : Returns true if the character at the index is upper-case
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'---------------¦----------------------------------------------------------------------
'Returns        | True if the character is upper-case
'======================================================================================
Public Property Get CharIsUpper( _
    ByVal Index As Long _
) As Boolean
    'Note that this uses Unicode properties, so covers far more than just A-Z
    Let CharIsUpper = (Me.CharProperties(Index) And C1_UPPER) <> 0
End Property

'CharIsVisible : Returns true if the character at the index is 'visible'
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'---------------¦----------------------------------------------------------------------
'Returns        | True if the character is visible
'======================================================================================
Public Property Get CharIsVisible( _
    ByVal Index As Long _
) As Boolean
    'An 'invisible' character is either a Control Code (such as Null), _
     or an undefined Unicode point. Zero-width space, whilst technically 'invisible' _
     in a proportional font still counts as 'visible' as it's a defined character
    Let CharIsVisible = (Me.CharProperties(Index) And C1_VISIBLE) <> 0
End Property

'CharProperties : Returns the Unicode properties for a character in the string
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'---------------¦----------------------------------------------------------------------
'Returns        | A bit battern that describes the Unicode properties of the character
'               | as defined in the C1 enum
'======================================================================================
Public Property Get CharProperties( _
    ByVal Index As Long _
) As C1
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not my_ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Property
    'Use the API to fill the return value with the character properties
    Call api_GetStringType( _
                   InfoType:=CT_CTYPE1, _
        SourceStringPointer:=VarPtr(Chars(Index)), _
         SourceStringLength:=1, _
            CharTypePointer:=VarPtr(CharProperties) _
    )
End Property

'CharPull : Removes the last character of the string and returns the character code
'======================================================================================
'Returns        | A Unicode character point (0 to 65535)
'======================================================================================
Public Function CharPull() As UCS2
    'If there are no characters, just return null
    If My_Length = 0 Then Exit Function
    'Return the character value in question _
     The user can use `And &HFF&` to get ASCII
    Let CharPull = (Chars(My_Length - 1) And &HFFFF&)
    'Reduce the size of the string
    Call my_SetLength(My_Length - 1)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'CharPush : Appends a single character code to the end of the string
'======================================================================================
'Char           | Unicode character point (0 to 65535)
'---------------¦----------------------------------------------------------------------
'Returns        | This class object (method chaining)
'======================================================================================
Public Function CharPush( _
    ByVal Char As UCS2 _
) As bluString
    'Allow method chaining
    Set CharPush = Me
    'Validate the character code. It could be negative if it came from a VB _
     signed-Integer, so this is converted to the positive equivilent
    If Not my_ValidateChar(Char) Then Err.Raise (9): Exit Function
    
    'Increase the length of our string by one character _
     (the buffer will manage any necessary growth)
    Call my_SetLength(My_Length + 1)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
    
    'Set the character code of the last character
    'VB6 can only do signed Integers, so Unicode points 32'768+ will error unless _
         we convert them to their signed equivilent to stuff them into two bytes
    If Char >= 32768 _
        Then Let Chars(My_Length - 1) = (Char And &HFFFF&) - 65536 _
        Else Let Chars(My_Length - 1) = (Char And &HFFFF&)
End Function

'CharPushInt : Appends a character code to the end of our string with less validation
'======================================================================================
'Char           | Unicode character point as signed Integer (-32768 to 32767)
'---------------¦----------------------------------------------------------------------
'Returns        | This class object (method chaining)
'======================================================================================
Public Function CharPushInt( _
    ByVal Char As Integer _
) As bluString
    'Allow method chaining
    Set CharPushInt = Me
    
    'Increase the length of our string by one character _
     (the buffer will manage any necessary growth)
    Call my_SetLength(My_Length + 1)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
    
    'Set the character code of the last character
    Let Chars(My_Length - 1) = Char
End Function

'CharRemove : Removes a single character from the string and returns its code
'======================================================================================
'[Index]        | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'---------------¦----------------------------------------------------------------------
'Returns        | A Unicode character point (0 to 65535)
'======================================================================================
Public Function CharRemove( _
    Optional ByVal Index As Long = 0 _
) As UCS2
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not my_ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Function
    
    'Return the character value
    Let CharRemove = Chars(Index)
    
    'If removing just the last character, we won't need to shift the data down
    If Index < (My_Length - 1) Then
        'Shift the remainder of the string down a character
        Call my_UnsafeCopy(Index + 1, Index, My_Length - Index - 1)
    End If
    
    'Reduce the overal length of the string by one
    Call my_SetLength(My_Length - 1)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

''CharCodeFromLong : Converts a Long character code to a signed Integer
''======================================================================================
''CharCode       |
''---------------¦----------------------------------------------------------------------
''Returns        | VB6's typical signed-integer character code (-32'768 to 32'767)
''======================================================================================
'Public Function CharCodeFromLong( _
'    ByRef CharCode As Long _
') As Integer
'
'End Function

'Clear : Go back to an empty string, keeps the buffer size
'======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Clear() As bluString
    'Allow method chaining
    Set Clear = Me
    'Mark the string as being null-length
    Call my_SetLength(0)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'Clone : Duplicate our string (or subportion) into another bluString instance
'======================================================================================
'[Index]        | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'[Length]       | Number of characters to copy, can be 0 which will imply the
'               | remainder of the string, from the Index
'---------------¦----------------------------------------------------------------------
'Returns        | A new bluString instance
'======================================================================================
Public Function Clone( _
    Optional ByVal Index As Long = 0, _
    Optional ByVal Length As Long = 0 _
) As bluString
    'Instantiate the return object
    Set Clone = New bluString
    
    'Validate the range: _
     * A negative Index will be recalculated backwards from the end of the string _
     * A Length of zero will default to the remainder of the string, from the Index
    If Not my_ValidateRange(Index, Length) Then Err.Raise (9): Exit Function
    
    'Set the destination size
    Let Clone.Length = Length
    'Copy the text across (if there is any)
    If My_Length > 0 Then Call Win32.MemoryCopyBetween( _
          LengthBytes:=Length * 2, _
        OutputPointer:=Clone.Pointer(Index), _
        SourcePointer:=VarPtr(Chars(0)) _
    )
End Function

'CRC : Do a Cyclic-Redundancy-Check on our string
'======================================================================================
'Returns        | CRC value of our string
'======================================================================================
Public Property Get CRC() As Long
    Dim i As Long, Value As Long
    
    'Has the CRC lookup table been built?
    If CRCInitialized = 0 Then
        'Note that this is going to add 1 KB to the class's memory usage, _
         and it takes time to build the lookup table, try to reuse an existing _
         bluString to do many CRCs in quick succession
        ReDim CRCTable(0 To 255) As Long
        'Create the CRC look-up table _
         (I will not pretend to understand how this works other than it goes through _
          shifting and flipping bits, make of that what you will)
        For i = 0 To 255
            Let Value = i
            Dim ii As Long
            For ii = 8 To 1 Step -1
                If (Value And 1) Then
                    Let Value = ((Value And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                    Let Value = Value Xor CRC_POLYNOMIAL
                Else
                    Let Value = ((Value And &HFFFFFFFE) \ 2&) And &H7FFFFFFF
                End If
            Next
            Let CRCTable(i) = Value
        Next
        'Don't repeat this process
        Let CRCInitialized = 1
    End If
    
    '----------------------------------------------------------------------------------
    
    'Has the CRC already been calculated?
    If CRCValid = 0 Then
        'Calculate the CRC by twiddling the bits
        Let My_CRC = &HFFFFFFFF
        For i = 0 To My_Length - 1
            'Our string is stored as Integers, so we need to do the lower and _
             upper bytes separately
            Let Value = (My_CRC And &HFF) Xor (Chars(i) And &HFF)
            Let My_CRC = ((My_CRC And &HFFFFFF00) \ &H100) And 16777215
            Let My_CRC = My_CRC Xor CRCTable(Value)
            'And the upper byte:
            Let Value = (My_CRC And &HFF) Xor ((Chars(i) And &HFF00&) \ &H100)
            Let My_CRC = ((My_CRC And &HFFFFFF00) \ &H100) And 16777215
            Let My_CRC = My_CRC Xor CRCTable(Value)
        Next
        'Cache the calculated CRC
        Let My_CRC = Not (My_CRC)
        Let CRCValid = 1
    End If
    
    'Return the cached value
    Let CRC = My_CRC
End Property

'Destroy : Sets the string to null and frees all memory
'======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Destroy() As bluString
    'Allow method chaining
    Set Destroy = Me
    
    'Nullify the string
    Let My_Length = 0: Let My_LengthB = 0
    
    'Deallocate the memory
    Erase Chars
    'Reset the buffer length so that any new content will create a new, _
     appropriately-sized buffer
    Let My_Buffer = 0
    'Note that no memory is reserved; `Pointer` will return 0!
    Let Reserved = 0
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'Equals : A fast method to check if one string equals another
'======================================================================================
'Text           | VB String to compare our string against
'---------------¦----------------------------------------------------------------------
'Returns        | True if strings are equal, False if not
'======================================================================================
Public Function Equals( _
    ByRef Text As String _
) As Boolean
    'Any length difference is immediately invalid!
    If LenB(Text) <> My_LengthB Then Exit Function
    
    'Handle null strings: _
     Two null strings are considered equal, otherwise result is inequal
    If My_Length = 0 Then
        Let Equals = (LenB(Text) = 0)
        Exit Function
    ElseIf LenB(Text) = 0 Then
        Let Equals = (My_Length = 0)
        Exit Function
    End If
    
    'Are we doing a case-insensitive comparison?
    '..................................................................................
    If My_CaseSensitive = 0 Then
        'Check for Windows XP; this is faster to do because it's a zero comparison
        If blu.sys.WindowsAPILevel = XP_OR_LOWER Then
            'Fall back to the XP-compatible API
            Let Equals = (api_CompareString( _
                      LocaleID:=Win32.ThreadLocale, _
                  CompareFlags:=NORM_IGNORECASE, _
                String1Pointer:=StrPtr(Text), String1Length:=-1, _
                String2Pointer:=VarPtr(Chars(0)), String2Length:=-1 _
            ) = CSTR_EQUAL)
        Else
            'NOTE: The use of null on the first parameter causes Windows to use the _
             user's current locale, so we don't need to manually fetch the locale name
            Let Equals = (api_CompareStringEx( _
                 LocaleNamePointer:=0&, _
                      CompareFlags:=LINGUISTIC_IGNORECASE, _
                    String1Pointer:=StrPtr(Text), String1Length:=-1, _
                    String2Pointer:=VarPtr(Chars(0)), String2Length:=-1, _
                VersionInfoPointer:=0&, _
                   ReservedPointer:=0&, _
                             Param:=0& _
            ) = CSTR_EQUAL)
        End If
    
    'Case-sensitive (binary) comparison
    '..................................................................................
    Else
        Let Equals = (api_CompareStringOrdinal( _
            String1Pointer:=StrPtr(Text), String1Length:=-1, _
            String2Pointer:=VarPtr(Chars(0)), String2Length:=-1, _
                IgnoreCase:=BOOL_FALSE _
        ) = CSTR_EQUAL)
    End If
End Function

'Format : Replace markers in the text with the values provided
'======================================================================================
'Params...      | An array of values (strings, numbers &c.) to insert
'---------------¦----------------------------------------------------------------------
'Returns        | A new bluString instance, formatted according to our string
'======================================================================================
Public Function Format( _
    ParamArray Params() As Variant _
) As bluString
    'Instantiate the return object. We clone this string first since if there are _
     no markers in the string, this function will still return the original string
    Set Format = Me.Clone
    
    'Locate the bounds of the list of parameters
    Dim First As Long, Last As Long
    Let First = LBound(Params): Let Last = UBound(Params)
    'If no parameters are provided!
    If Last < First Then Exit Function
    
    'For non-String parameters, we will do the String conversion, but we will need _
     somewhere to store them so that we can get at their pointers. The FormatData _
     array provides an easy means of managing the text, length and pointer for each.
    Dim FormatData() As bluString_FormatData
    ReDim FormatData(0) As bluString_FormatData
    
    'Walk the params provided, converting them to Strings.
    Dim Param As Long
    For Param = First To Last
        Call my_FormatVariant(Params(Param), FormatData)
    Next Param
    
    '----------------------------------------------------------------------------------
    
    'Speed up references to the format marker we'll be searching for
    Dim MarkerLengthB As Long:  Let MarkerLengthB = LenB(FORMAT_MARKER)
    Dim MarkerPointer As Long:  Let MarkerPointer = StrPtr(FORMAT_MARKER)
    
    'Begin with our constructed string array, which may be longer than the input _
     params due to Variant arrays. As we find each marker, we'll move to the string. _
     It's possible there will be too many or too few markers
    Let Param = 1
    
    'The API returns a pointer in the string, not an index. Since replacements could _
     extend the string such that it gets reallocated we cannot rely on a pointer _
     returned from a search to be valid after the replacement so we need to calculate _
     the character index from the pointer to be able to continue where we left off _
     -- replacements could contain the search term, leading to an infinite loop!
    Dim StartPointer As Long
    Let StartPointer = Format.Pointer()
    
Continue:
    'If the replacement string is shorter than the marker, then we'll need _
     to remove parts of our string and vice-versa. Work out the impact the _
     replacement will have on the length of our string
    Dim Difference As Long
    Let Difference = FormatData(Param).ByteCount - LenB(FORMAT_MARKER)
    
    'Search for the format marker
    Dim ResultPointer As Long
    Let ResultPointer = api_StrStr( _
        HaystackPointer:=StartPointer, _
          NeedlePointer:=MarkerPointer _
    )
    
    'If the returned pointer is null, the marker was not found
    If ResultPointer = 0 Then GoTo Finish
    
    'From the result pointer, calculate the byte offset index in the string _
     (the pointer may become invalid after string replacement!)
    Dim IndexB As Long
    Let IndexB = ResultPointer - Format.Pointer()
    
    'If both the marker and the replacement are the same length, _
     we don't need to move parts of the string about
    If Difference = 0 Then
        'Replace the marker
        Call Win32.MemoryCopyBetween( _
              LengthBytes:=FormatData(Param).ByteCount, _
            OutputPointer:=ResultPointer, _
            SourcePointer:=FormatData(Param).Pointer _
        )
    Else
        'Calculate the remainder portion of our string that will be moved
        'Calculate the remainder portion of our string that will be moved
        Dim Remainder As Long
        Let Remainder = Format.LengthBytes - (IndexB + MarkerLengthB)
        
        'If the replacement string is longer than the search string, _
         our string will get longer overall
        'WARNING: If the string increases beyond the buffer, _
         it will be reallocated, making `ResultPointer` invalid!
        If Difference > 0 Then
            Let Format.Length = (Format.LengthBytes + Difference) \ 2
            'Get the potentially new string pointer
            Let StartPointer = Format.Pointer()
            Let ResultPointer = StartPointer + IndexB
        End If
        
        'Shift up or down the remainder portion of our string
        Call Win32.MemoryCopyWithin( _
            OutputPointer:=ResultPointer + FormatData(Param).ByteCount, _
            SourcePointer:=ResultPointer + MarkerLengthB, _
              LengthBytes:=Remainder _
        )
        'If the replacement string is shorter than the search string, _
         our string will get shorter overall
        'WARNING: If the string increases beyond the buffer, _
         it will be reallocated, making `StartPointer` invalid!
        If Difference < 0 Then
            Let Format.Length = (Format.LengthBytes + Difference) \ 2
            'Get the potentially new string pointer
            Let StartPointer = Format.Pointer()
            Let ResultPointer = StartPointer + IndexB
        End If
        
        'Copy the replacement string in
        Call Win32.MemoryCopyBetween( _
              LengthBytes:=FormatData(Param).ByteCount, _
            OutputPointer:=ResultPointer, _
            SourcePointer:=FormatData(Param).Pointer _
        )
    End If
    
    'We don't want an infinite loop, so move our search pointer over the replacement
    Let StartPointer = ResultPointer + FormatData(Param).ByteCount
    'Move onto the next parameter
    Let Param = Param + 1: If Param > UBound(FormatData) Then GoTo Finish
    'Proceed with the next search
    GoTo Continue
    
    '----------------------------------------------------------------------------------
Finish:
    'And clean up
    Erase FormatData
End Function

'Insert : Insert some text in the middle of the existing string
'======================================================================================
'Text           | VB String to insert into our string
'Index          | Position within our string (0-based). Can be negative, in which case
'               | the Index is calculated from the end of the string working backwards
'---------------¦----------------------------------------------------------------------
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
    If Not my_ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Function
    
    'Increase the length of our string to fit the inserted string _
     (the buffer will be increased automatically if necessary)
    Call my_SetLength(My_Length + Len(Text))
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
    
    'Shift the upper portion of our string up to make room for the inserted string
    Call my_UnsafeCopy(Index, Index + Len(Text), My_Length - Index)
    'Insert the new string
    Call Win32.MemoryCopyBetween( _
          LengthBytes:=LenB(Text), _
        OutputPointer:=VarPtr(Chars(Index)), _
        SourcePointer:=StrPtr(Text) _
    )
End Function

'IsNull
'======================================================================================
'Returns        | True if our string is Null (length = 0), False otherwise
'======================================================================================
Public Property Get IsNull() As Boolean
    'This is intended for readability, since you can obviously just check Length = 0
    Let IsNull = (My_Length = 0)
End Property

'Join : Appends another bluString to the end of this one
'======================================================================================
'bluStringObj   | Another bluString instance to take text from
'---------------¦----------------------------------------------------------------------
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Join( _
    ByRef bluStringObj As bluString _
) As bluString
    'Allow method chaining
    Set Join = Me
    'Appending a null-string will do nothing
    If bluStringObj.Length = 0 Then Exit Function
    
    'The current length will be the starting position of the appended string
    Dim Start As Long: Let Start = My_Length
    
    'Increase the length accordingly
     Call my_SetLength(My_Length + bluStringObj.Length)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
    
    'Copy the new string in
    Call Win32.MemoryCopyBetween( _
          LengthBytes:=bluStringObj.LengthBytes, _
        OutputPointer:=VarPtr(Chars(Start)), _
        SourcePointer:=bluStringObj.Pointer _
    )
End Function

'LCase : Convert our string to lowercase
'======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function LCase() As bluString
    'Allow method chaining
    Set LCase = Me
    'Nullstring won't change!
    If My_Length = 0 Then Exit Function
    
    'Check for Windows XP; this is faster to do because it's a zero comparison
    If blu.sys.WindowsAPILevel = XP_OR_LOWER Then
        'Use the Windows XP compatible API
        Call Win32.TextConvertCase_NT5( _
                    LocaleID:=Win32.ThreadLocale, _
                       Flags:=vbCaseLower Or vbCaseLinguistic, _
         SourceStringPointer:=VarPtr(Chars(0)), _
          SourceStringLength:=My_LengthB, _
         ResultStringPointer:=VarPtr(Chars(0)), _
          ResultStringLength:=My_LengthB _
        )
    Else
        'NOTE: The use of null on the first parameter causes Windows to use the _
         user's current locale, so we don't need to manually fetch the locale name
        Call Win32.TextConvertCase_NT6( _
             LocaleNamePointer:=0&, _
                          Flags:=vbCaseLower Or vbCaseLinguistic, _
           SourceStringPointer:=VarPtr(Chars(0)), _
            SourceStringLength:=My_LengthB, _
           ResultStringPointer:=VarPtr(Chars(0)), _
            ResultStringLength:=My_LengthB, _
          Reserved_VersionInfo:=0&, Reserved:=0&, Reserved_SortHandle:=0& _
        )
    End If
    
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'Left : Cut the string short to a given number of characters
'======================================================================================
'Length         | Number of characters to keep of the string. If negative, this will
'               | be calculated backwards from the end of the string
'[FillChar]     | If the new Length is longer than our string, the gap will be
'               | filled with the given character, which defaults to Space
'---------------¦----------------------------------------------------------------------
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Left( _
             ByRef Length As Long, _
    Optional ByRef FillChar As UCS2 = SPACE_ _
) As bluString
    'Allow method chaining
    Set Left = Me
    
    'Shortcut if trimming to zero
    If Length = 0 Then Call Me.Destroy: Exit Function
    'Trimming to the same length does nothing!
    If Length = My_Length Then Exit Function
    'Increasing the length of our string will fill the gap, _
     you can use this to left-align strings and numbers
    If Length > My_Length Then
        'Remember the current length, where the gap will begin from
        Dim OldLength As Long
        Let OldLength = My_Length
        'Increase the string space
        Let Me.Length = Length
        'Fill the new space with the character
        Dim i As Long
        For i = OldLength To My_Length - 1: Let Chars(i) = FillChar: Next i
    
    'If the position is negative, work out 'x' number of characters before the end of _
     our string. That is, -1 will be the second to last character
    ElseIf Length < 0 Then
        'The negative position cannot be more than the length of our string
        If Abs(Length) > My_Length Then Err.Raise (9): Exit Function
        'Calculate the new position from the end of our string
        Call my_SetLength(My_Length + Length)
    Else
        'Do the reduction
        Call my_SetLength(Length)
    End If
    
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'Length : Length of the actual stored text (1-based), not the buffer
'======================================================================================
Public Property Get Length() As Long
    Let Length = My_Length
End Property
Public Property Let Length( _
    ByRef NewLength As Long _
)
    'A non-change in length can be skipped
    If My_Length = NewLength Then Exit Property
    
    'Cannot obviously set the length to a negative number
    If NewLength < 0 Then Err.Raise (9): Exit Property
    'Can't set the length greater than the maximum allowed size _
     (VB Error #6 is "Overflow")
    If NewLength > BUFFER_MAX Then Err.Raise (6): Exit Property

    'Store the new desired content length (1-based)
    Call my_SetLength(NewLength)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Property

'LengthBytes : Length of the content in bytes, cannot be set
'======================================================================================
Public Property Get LengthBytes() As Long
    Let LengthBytes = My_LengthB
End Property

'Mid : Reduce the string to a sub-string portion
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case
'               | the Index is calculated from the end of the string working backwards
'[Length]       | Number of characters to keep, can be 0 which will imply the
'               | remainder of the string, from the Index
'---------------¦----------------------------------------------------------------------
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Mid( _
             ByVal Index As Long, _
    Optional ByVal Length As Long = 0 _
) As bluString
    'Allow method chaining
    Set Mid = Me
    
    'Validate the range: _
     * A negative Index will be recalculated backwards from the end of the string _
     * A Length of zero will default to the remainder of the string, from the Index
    If Not my_ValidateRange(Index, Length) Then Err.Raise (9): Exit Function
    
    'Copy the sub-string to the beginning of our string
    Call my_UnsafeCopy(Index, 0, Length)
    
    'Reduce the string to the new sub-string length
    Call my_SetLength(Length)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'Overwrite : Overwrite a portion of our string with text
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case
'               | the Index is calculated from the end of the string working backwards
'Text           | New text to overwrite our string at the given Index
'---------------¦----------------------------------------------------------------------
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Overwrite( _
    ByVal Index As Long, _
    ByRef Text As String _
) As bluString
    'Allow method chaining
    Set Overwrite = Me
    
    'Null-string won't change anything!
    If Len(Text) = 0 Then Exit Function
    
    'Validate the range: _
     * A negative Index will work backwards from the end of our string _
     * If the length of the new text overhangs our string, it will be clipped
    Dim Length As Long: Let Length = Len(Text)
    If Not my_ValidateRange(Index, Length) Then Err.Raise (9): Exit Function
    
    'Now just copy the new text over the top of our string
    Call Win32.MemoryCopyBetween( _
          LengthBytes:=Length * 2, _
        OutputPointer:=VarPtr(Chars(Index)), _
        SourcePointer:=StrPtr(Text) _
    )
    
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'Pointer : Allow access to the "string" pointer for external copying
'======================================================================================
'[Index]        | Position within our string (0-based)
'---------------¦----------------------------------------------------------------------
'Returns        | Memory pointer to the character index within our string
'======================================================================================
Public Property Get Pointer( _
    Optional ByVal Index As Long = 0 _
) As Long
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not my_ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Property
    
    'If the buffer hasn't been reserved yet, return a null-pointer
    If Reserved = 0 _
        Then Let Pointer = 0 _
        Else Let Pointer = VarPtr(Chars(Index))
End Property

'Prepend : Add our string to the end of the one provided
'======================================================================================
'Text           | VB String to prepend to our string
'---------------¦----------------------------------------------------------------------
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
    Call my_SetLength(My_Length + Start)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
    
    'Shift our string up
    Call my_UnsafeCopy(0, Start, My_Length)
    'Copy the new string in
    Call Win32.MemoryCopyBetween( _
          LengthBytes:=LenB(Text), _
        OutputPointer:=VarPtr(Chars(0)), _
        SourcePointer:=StrPtr(Text) _
    )
End Function

'Remove : Cut some text out of the middle of the existing string
'======================================================================================
'Index          | Position within our string (0-based). Can be negative, in which case
'               | the Index is calculated from the end of the string working backwards
'[Length]       | Number of characters to remove, 0 = remainder of string from Index
'---------------¦----------------------------------------------------------------------
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Remove( _
             ByVal Index As Long, _
    Optional ByVal Length As Long = 0 _
) As bluString
    'Allow method chaining
    Set Remove = Me
    
    'Validate the range: _
     * A negative Index will be recalculated backwards from the end of the string _
     * A Length of zero will default to the remainder of the string, from the Index
    If Not my_ValidateRange(Index, Length) Then Err.Raise (9): Exit Function
    
    'If you end up removing the whole string, make it a null-string
    If Index = 0 Then
        If Length = My_Length - 1 Then Call Me.Destroy: Exit Function
    End If
    
    'Shift the remainder down into the removed area
    Call my_UnsafeCopy(Index + Length, Index, Length)
    'Update the string length
    Call my_SetLength(My_Length - Length)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'Replace : Search for a sub-string and replace it with another
'======================================================================================
'SearchFor      | Text to find within our string
'[ReplaceWith]  | What to replace the found text with, can be null-string
'[ReturnCount]  | If provided, your own variable will be set with the number of
'               | replacements done
'---------------¦----------------------------------------------------------------------
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Replace( _
             ByRef SearchFor As String, _
    Optional ByRef ReplaceWith As String = vbNullString, _
    Optional ByRef ReturnCount As Long = 0 _
) As bluString
    'Allow method chaining
    Set Replace = Me
    'Ensure that the count will return 0 if nothing is found _
     (especially if the variable passed in already had a value)
    Let ReturnCount = 0
    
    'You can't search in a null-string!
    If My_Length = 0 Then Exit Function
    'You can't search for a null-string!
    If Len(SearchFor) = 0 Then Exit Function
    'What you're searching for cannot be longer than our string
    If Len(SearchFor) > My_Length Then Exit Function
    
    'If the replacement string is shorter than the search string, then we'll need _
     to remove parts of our string and vice-versa. Work out the impact the _
     replacement will have on the length of our string
    Dim Difference As Long
    Let Difference = Len(ReplaceWith) - Len(SearchFor)
    
    'The API returns a pointer in the string, not an index. Since replacements could _
     extend the string such that it gets reallocated we cannot rely on a pointer _
     returned from a search to be valid after the replacement so we need to calculate _
     the character index from the pointer to be able to continue where we left off _
     -- replacements could contain the search term, leading to an infinite loop!
    Dim StartPointer As Long
    Let StartPointer = VarPtr(Chars(0))
    
    'Some caches before we start looping
    Dim StrPtr_SearchFor As Long:   Let StrPtr_SearchFor = StrPtr(SearchFor)
    Dim StrPtr_ReplaceWith As Long: Let StrPtr_ReplaceWith = StrPtr(ReplaceWith)
    Dim LenB_SearchFor As Long:     Let LenB_SearchFor = LenB(SearchFor)
    Dim LenB_ReplaceWith As Long:   Let LenB_ReplaceWith = LenB(ReplaceWith)
    
Continue:
    '----------------------------------------------------------------------------------
    'Search for the sub-string: _
     (case-sensitivity is handled by the class `CaseSensitive` property)
    Dim ResultPointer As Long
    If My_CaseSensitive = 0 _
        Then Let ResultPointer = api_StrStr(StartPointer, StrPtr_SearchFor) _
        Else Let ResultPointer = api_StrStrI(StartPointer, StrPtr_SearchFor)
    
    'If the returned pointer is null, the sub-string was not found
    If ResultPointer = 0 Then GoTo Finish
    
    'From the result pointer, calculate the byte offset index in our string _
     (the pointer may become invalid after string replacement!)
    Dim IndexB As Long
    Let IndexB = ResultPointer - VarPtr(Chars(0))
    
    'If both the search string and the replacement are the same length, _
     we don't need to move parts of our string about
    If Difference = 0 Then
        'Replace the found string
        Call Win32.MemoryCopyBetween( _
            OutputPointer:=ResultPointer, _
            SourcePointer:=StrPtr_ReplaceWith, _
              LengthBytes:=LenB_ReplaceWith _
        )
    Else
        'Calculate the remainder portion of our string that will be moved
        Dim Remainder As Long
        Let Remainder = My_LengthB - (IndexB + LenB_SearchFor)
        
        'If the replacement string is longer than the search string, _
         our string will get longer overall
        'WARNING: If the string increases beyond the buffer, _
         it will be reallocated, making `ResultPointer` invalid!
        If Difference > 0 Then
            Call my_SetLength(My_Length + Difference)
            'Get the potentially new string pointer
            Let StartPointer = VarPtr(Chars(0))
            Let ResultPointer = StartPointer + IndexB
        End If
        
        'Shift up or down the remainder portion of our string
        Call Win32.MemoryCopyWithin( _
            OutputPointer:=ResultPointer + LenB_ReplaceWith, _
            SourcePointer:=ResultPointer + LenB_SearchFor, _
              LengthBytes:=Remainder _
        )
        'If the replacement string is shorter than the search string, _
         our string will get shorter overall
        'WARNING: If the string increases beyond the buffer, _
         it will be reallocated, making `StartPointer` invalid!
        If Difference < 0 Then
            Call my_SetLength(My_Length + Difference)
            'Get the potentially new string pointer
            Let StartPointer = VarPtr(Chars(0))
            Let ResultPointer = StartPointer + IndexB
        End If
        
        'Copy the replacement string in
        Call Win32.MemoryCopyBetween( _
            OutputPointer:=ResultPointer, _
            SourcePointer:=StrPtr_ReplaceWith, _
              LengthBytes:=LenB_ReplaceWith _
        )
    End If
    
    'We don't want an infinite loop, so move our search pointer over the replacement
    Let StartPointer = ResultPointer + LenB_ReplaceWith
    'Increase the result count
    Let ReturnCount = ReturnCount + 1
    
    GoTo Continue
    
Finish:
    '----------------------------------------------------------------------------------
    'Changes to our string will invalidate the CRC
    If ReturnCount <> 0 Then Let CRCValid = 0
End Function

'Right : Keep only the given right-most number of characters
'======================================================================================
'Length         | Number of characters to keep working backwards from the right
'---------------¦----------------------------------------------------------------------
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
    If Length = 0 Then Call Me.Destroy: Exit Function
    'Negative numbers are not supported, just use `Left` to keep x number of chars
    If Length < 0 Then Err.Raise (9): Exit Function
    'Shift the rightmost portion of the string down the beginning
    Call my_UnsafeCopy(My_Length - Length, 0, Length)
    
    'Cut the string short. The buffer will be left at its old length, _
     if you know the string won't expand any more you can free the memory up _
     by calling the `ReduceBuffer` method
    Call my_SetLength(Length)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

''SplitLines : Split a string into an array of lines (breaking at new lines)
''=====================================================================================
''Returns        | An array of strings
''=====================================================================================
'Public Function SplitLines( _
') As String()
'    'A null-string won't split!
'    If My_Length = 0 Then Exit Function
'
'    'This will be the number of lines found (1-based)
'    Dim Lines As Long
'    Dim LineBegin As Long
'
'    'Walk the string looking for line-breaks
'    Dim i As Long
'    For i = 0 To My_Length
'        'Look for a line-break. Windows uses CR + LF, Unix just LF
'        If Data(i) = vbLf Then
'            'Add the line
'            Let Lines = Lines + 1
'            ReDim Preserve SplitLines(1 To Lines) As String
'
'            Dim LineLength As Long
'            Let LineLength = i - LineBegin
'
'            'TODO: handle zero-length
'
'            Let LineBegin = i
'        End If
'    Next i
'
'End Function

'SplitWords : Split our string into words (separated by white-space)
'======================================================================================
'[QuotedText]   | Keeps quoted text ("...") together as one word
'---------------¦----------------------------------------------------------------------
'Returns        | An array of strings, each containing a word. If there are no words
'               | in the string, a zero-based single entry array containing nullstring
'               | will be returned. When there are results, the array will be 1-based
'======================================================================================
Public Function SplitWords( _
    Optional ByRef QuotedText As Boolean = False _
) As String()
    'If our string is Null, nothing can be done; we return a single array entry so as _
     not to confuse the caller with an uninitialised array, which are hard to detect
    If My_LengthB = 0 Then
        ReDim SplitWords(0) As String
        Exit Function
    End If
    
    'Cache zero-based length of our string
    Dim Length0 As Long
    Let Length0 = My_Length - 1
    
    'Create a duplicate copy of our string converted to its Unicode properties
    Dim Whitespace() As Integer
    ReDim Whitespace(0 To Length0) As Integer
    
    'Call the Windows API; this will convert each character of our string _
     into a bit-pattern of its Unicode properties (see the `C1` enum)
    Call api_GetStringType( _
                   InfoType:=CT_CTYPE1, _
        SourceStringPointer:=VarPtr(Chars(0)), _
         SourceStringLength:=My_Length, _
            CharTypePointer:=VarPtr(Whitespace(0)) _
    )
    
    '----------------------------------------------------------------------------------
    
    Dim i As Long
    Dim WordBegin As Long, WordEnd As Long, WordLength As Long
    Dim IsWord As Long, IsQuoted As Long, WordCount As Long
    
    'Due to the limitation of VB syntatic sugar, we can't assign values to the return _
     array because VB interprets this as a function call instead. We'll populate a _
     temporary array and then assign it to the return array at the end
    Dim OutArray() As String
    
    'Walk the characters in our string
    For i = 0 To Length0
        'Are we currently processing a word? _
         If not, we will look for the first non-whitespace character
        '..............................................................................
        If IsWord = 0 Then
            'Is this a non-whitespace character? (the start of a word)
            If (Whitespace(i) And C1.C1_WHITESPACE) = 0 Then
                'Are we matching quotes?
                If QuotedText <> 0 Then
                    'Is this a quote?
                    If Chars(i) = UCS2.QUOTATION_MARK Then
                        Let IsQuoted = 1
                    End If
                End If
                
                'Note the start of the word and change the processing mode _
                 for the next character (look for the end of the word)
                Let WordBegin = i
                Let IsWord = 1
            End If
            
        '..............................................................................
        Else
            'Are we matching quotes?
            If IsQuoted <> 0 Then
                If Chars(i) = UCS2.QUOTATION_MARK Then
                    Let IsQuoted = 0
                    Let WordEnd = i + 1
                    GoSub AddWord
                End If
            Else
                'Is this a whitespace character? (end of the word)
                If (Whitespace(i) And C1.C1_WHITESPACE) <> 0 Then
                    Let WordEnd = i
                    GoSub AddWord
                End If
            End If
        End If
    Next i
    
    'We're done with the Unicode properties
    Erase Whitespace
    
    'It is likely that our string will end without whitespace, _
     so the loop will have ended without adding the final word
    If IsWord = 1 Then
        Let WordEnd = i
        GoSub AddWord
    Else
        'Should no words have been found, return nullstring
        If WordCount = 0 Then ReDim OutArray(0) As String
    End If
    
    Let SplitWords = OutArray
    Exit Function
    
AddWord:
    '..................................................................................
    Let IsWord = 0
    
    'Add the word to array
    Let WordCount = WordCount + 1
    ReDim Preserve OutArray(1 To WordCount) As String
    
    Let WordLength = WordEnd - WordBegin
    Let OutArray(WordCount) = Space$(WordLength)
    Call Win32.MemoryCopyBetween( _
        OutputPointer:=StrPtr(OutArray(WordCount)), _
        SourcePointer:=VarPtr(Chars(WordBegin)), _
          LengthBytes:=WordLength * 2 _
    )
    Return
End Function

'Text : Get/Let the contents as native VB strings
'======================================================================================
Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "200"
    'If ours is a null-string, return a null-string
    If My_Length = 0 Then
        Let Text = vbNullString
    Else
        'Pad the return String to the right size
        Let Text = Win32.CreateUninitializedString(ByVal 0&, My_Length)
        'Copy our String into the return field
        Call Win32.MemoryCopyBetween( _
            OutputPointer:=StrPtr(Text), _
            SourcePointer:=VarPtr(Chars(0)), _
              LengthBytes:=My_LengthB _
        )
    End If
End Property
Public Property Let Text( _
    ByRef NewString As String _
)   'If a null-string is provided then set ours to a null string
    If LenB(NewString) = 0 Then
        Call my_SetLength(0)
    Else
        'Set the new length of our string; _
         The buffer will be allocated appropriately
        Call my_SetLength(Len(NewString))
        'Copy the VB String data over
        Call Win32.MemoryCopyBetween( _
            OutputPointer:=VarPtr(Chars(0)), _
            SourcePointer:=StrPtr(NewString), _
              LengthBytes:=LenB(NewString) _
        )
    End If
    'Invalidate the CRC
    Let CRCValid = 0
End Property

'TrimBuffer : Free some memory by reducing the buffer to the string length
'======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function TrimBuffer() As bluString
    'Of course you could just use `bluString.Buffer = bluString.Length`, but this is _
     intended to be more obvious and safer for potential future changes / expansion. _
     It's strongly advised NOT to expand the string after you've reduced the buffer

    'Allow method chaining
    Set TrimBuffer = Me
    'Reduce the buffer down to the size of the string
    Let Me.Buffer = My_Length
End Function

'TrimSpace : Strip whitespace characters from the beginning and end
'======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function TrimSpace( _
) As bluString
    'Allow method chaining
    Set TrimSpace = Me
    'This is just a macro
    Call TrimSpaceLeft
    Call TrimSpaceRight
End Function

'TrimSpaceLeft : Strip whitespace characters from the beginning
'======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function TrimSpaceLeft( _
) As bluString
    'Allow method chaining
    Set TrimSpaceLeft = Me
    'Null string won't trim!
    If My_Length = 0 Then Exit Function
    
Continue:
    'Walk along the left of the string looking for whitespace/blank characters
    Dim Index As Long
    Dim Props As C1
    Let Props = Me.CharProperties(Index)
    
    'If the character is whitespace, or not a defined Unicode point, then ...
    If ((Props And C1_WHITESPACE) <> 0) Or ((Props And C1_DEFINED) = 0) Then
        'Move along to the next char
        Let Index = Index + 1
        'Have we hit the end of our string?
        If Index = My_Length Then
            'The whole string is whitespace -- nullify it
            Call Me.Destroy
            Exit Function
        End If
        'Check the next character
        GoTo Continue
    End If
    
    'The first non-space character has been found, remove this portion of the string
    Call my_UnsafeCopy(Index, 0, My_Length - Index)
    'Update the string length
    Call my_SetLength(My_Length - Index)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'TrimSpaceRight : Strip whitespace characters from the end
'======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function TrimSpaceRight( _
) As bluString
    'Allow method chaining
    Set TrimSpaceRight = Me
    'Null string won't trim!
    If My_Length = 0 Then Exit Function
    
    Dim Index As Long
    Let Index = My_Length - 1
    
Continue:
    'Walk backwards along the string looking for whitespace/blank characters
    Dim Props As C1
    Let Props = Me.CharProperties(Index)
    
    'If the character is whitespace, or not a defined Unicode point, then ...
    If ((Props And C1_WHITESPACE) <> 0) Or ((Props And C1_DEFINED) = 0) Then
        'Move along to the next char
        Let Index = Index - 1
        'Have we hit the beginning of our string?
        If Index = -1 Then
            'The whole string is whitespace -- nullify it
            Call Me.Destroy
            Exit Function
        End If
        'Check the next character
        GoTo Continue
    End If
    
    'The last non-space character has been found, trim the string here
    Call my_SetLength(Index + 1)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'TrimToNull : Set our string length according to the first null found in the buffer
'======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function TrimToNull( _
) As bluString
    'You must make sure that you have set the buffer size big enough when using API _
     calls (for example `MAX_PATH`) otherwise Windows might overrun the buffer
    
    'Allow method chaining
    Set TrimToNull = Me
    
    'We now clip our string to fit the length given by the first Null found
    Call my_SetLength( _
        api_lstrlen(VarPtr(Chars(0))) _
    )
End Function

'Truncate : Line-break our string to a maximum given length, adding ellipsis if needed
'======================================================================================
'MaxLength      | Maximum allowable length of our string. This length is guaranteed,
'               | as the ellipsis will be placed within this length and not appended
'[Ellipsis]     | A string to include at the end if our string is clipped short
'---------------¦----------------------------------------------------------------------
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Truncate( _
             ByVal MaxLength As Long, _
    Optional ByRef Ellipsis As String = "…" _
) As bluString
    'Allow method chaining
    Set Truncate = Me
    
    'Validation:
    '----------------------------------------------------------------------------------
    'Truncating to zero will simply return nullstring
    If MaxLength = 0 Then Call Me.Destroy: Exit Function
    'Negative length not allowed
    If MaxLength < 0 Then Call Err.Raise(9): Exit Function
    
    'Remove any whitespace from the end of our string! _
     We don't want our ellipsis hanging in space
    'TODO: This duplicates a lot of what we're doing in this method, _
     so we should do this manually here?
    Call Me.TrimSpaceRight
    
    'You can't truncate a null-string
    If My_Length = 0 Then Exit Function
    'If our string is shorther than the given truncation, no work needs to be done
    If My_Length <= MaxLength Then Exit Function
    
    'Cache length of the ellipsis provided; _
     when it's longer than one character, we'll need to factor this into truncation
    Dim EllipsisLength As Long
    Let EllipsisLength = Len(Ellipsis)
    
    'If the ellipsis is longer than the maximum length, we can't possibly truncate! _
     The only viable solution is to clip the string by ignoring the ellipsis
    If EllipsisLength > MaxLength Then
        Let EllipsisLength = 0
    'If the ellipsis is the same length as the maximum, then our string is inevitably _
     truncated and we can just replace our string with the ellipsis
    ElseIf EllipsisLength = MaxLength Then
        'Replace our string with the ellipsis _
         (this also clears the CRC)
        Call my_SetLength(EllipsisLength)
        Call Win32.MemoryCopyBetween( _
            OutputPointer:=VarPtr(Chars(0)), _
            SourcePointer:=StrPtr(Ellipsis), _
              LengthBytes:=LenB(Ellipsis) _
        )
        Exit Function
    End If
    
    'Beyond this point, an ellipsis will inevitably be inserted; this needs to be _
     within the maximum length given so that the user can rely on the string length _
     afterwards, therefore assume the ellipsis on the end
    Let MaxLength = MaxLength - EllipsisLength
    
    'Look for white-space:
    '----------------------------------------------------------------------------------
    
    'In general we want to walk backwards from the truncation point looking for the _
     first space to word-break, however we want to fill the truncated string as much _
     as possible and don't want to word-break near the beginning of our string due to _
     a long non-breaking 'word' overlapping the truncation point. For that reason we _
     will always truncate mid-word at the given point if we can't find a space 15 _
     letters behind the given truncation point (and ellipsis)
     
    'For speed, we'll fetch up to 15 Unicode character properties at once. _
     If our truncation length is shorter than that, don't overflow the buffer!
    Dim PropsLength As Long
    If MaxLength < 15 Then Let PropsLength = MaxLength Else Let PropsLength = 15
    Dim PropsIndex As Long
    Let PropsIndex = MaxLength - PropsLength
    
    'Call the Windows API to fetch the Unicode character properties
    'TODO: Handle API errors
    Dim CharProps(0 To 14) As Integer
    Call api_GetStringType( _
                   InfoType:=CT_CTYPE1, _
        SourceStringPointer:=VarPtr(Chars(PropsIndex)), _
         SourceStringLength:=PropsLength, _
            CharTypePointer:=VarPtr(CharProps(0)) _
    )
    
    'Now walk backwards from the truncation point
    Dim i As Long
    Let i = (PropsLength - 1)
    Do While i >= 0
        'If the character is whitespace, or not a defined Unicode point, then ...
        If ((CharProps(i) And C1_WHITESPACE) <> 0) _
        Or ((CharProps(i) And C1_DEFINED) = 0) _
        Then
            'Whitespace has been found, so truncate here
            Call my_SetLength(PropsIndex + i + EllipsisLength)
            'Changes to our string will invalidate the CRC
            Let CRCValid = 0
            'Insert the ellipsis at the end
            If EllipsisLength <> 0 Then Call Win32.MemoryCopyBetween( _
                OutputPointer:=VarPtr(Chars(PropsIndex + i)), _
                SourcePointer:=StrPtr(Ellipsis), _
                  LengthBytes:=LenB(Ellipsis) _
            )
            Exit Function
        End If
        'Keep searching...
        Let i = i - 1
    Loop
    
    'Force truncation:
    '----------------------------------------------------------------------------------
    'Truncate our string at the given point, including the ellipsis
    Call my_SetLength(MaxLength + EllipsisLength)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
    
    'Insert the ellipsis at the end
    If EllipsisLength <> 0 Then Call Win32.MemoryCopyBetween( _
        OutputPointer:=VarPtr(Chars(MaxLength)), _
        SourcePointer:=StrPtr(Ellipsis), _
          LengthBytes:=LenB(Ellipsis) _
    )
End Function

'UCase : Convert our string to UPPERCASE
'======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function UCase() As bluString
    'Allow method chaining
    Set UCase = Me
    'Nullstring won't change!
    If My_Length = 0 Then Exit Function
    
    'Check for Windows XP; this is faster to do because it's a zero comparison
    If blu.sys.WindowsAPILevel = XP_OR_LOWER Then
        'Use the Windows XP compatible API
        Call Win32.TextConvertCase_NT5( _
                    LocaleID:=Win32.ThreadLocale(), _
                       Flags:=vbCaseUpper Or vbCaseLinguistic, _
         SourceStringPointer:=VarPtr(Chars(0)), _
          SourceStringLength:=My_LengthB, _
         ResultStringPointer:=VarPtr(Chars(0)), _
          ResultStringLength:=My_LengthB _
        )
    Else
        'NOTE: The use of null on the first parameter causes Windows to use the _
         user's current locale, so we don't need to manually fetch the locale name
        Call Win32.TextConvertCase_NT6( _
             LocaleNamePointer:=0&, _
                         Flags:=vbCaseUpper Or vbCaseLinguistic, _
           SourceStringPointer:=VarPtr(Chars(0)), _
            SourceStringLength:=My_LengthB, _
           ResultStringPointer:=VarPtr(Chars(0)), _
            ResultStringLength:=My_LengthB, _
          Reserved_VersionInfo:=0&, Reserved:=0&, Reserved_SortHandle:=0& _
        )
    End If
    
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'Wrap : Append & Prepend our string, e.g. for adding quotes around text
'======================================================================================
'Before         | String to prepend to our string
'[After]        | String to append to our string. If ommitted, defaults to Before
'---------------¦----------------------------------------------------------------------
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Wrap( _
             ByRef Before As String, _
    Optional ByRef After As String = vbNullString _
) As bluString
    'Allow method chaining
    Set Wrap = Me
    
    'If the second param is ommitted, use the same string for both ends
    If Len(After) = 0 Then Let After = Before
    
    'Caching
    Dim Len_Before As Long, Len_After As Long
    Let Len_Before = Len(Before)
    Let Len_After = Len(After)
    
    'If both before and after are null, nothing will happen!
    If (Len_Before = 0) And (Len_After = 0) Then Exit Function
    
    'Adjust the length of our string to account for the added before and after
    Dim OldLength As Long
    Let OldLength = My_Length
    
    Call my_SetLength(Len_Before + My_Length + Len_After)
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
    
    'Shift our string up to account for the before _
     (assuming Before is not null-string)
    If Len_Before > 0 Then Call my_UnsafeCopy(0, Len_Before, OldLength)
    'Copy in the new before and after
    If Len_Before > 0 Then Call my_UnsafeCopyString(Before, 0)
    If Len_After > 0 Then Call my_UnsafeCopyString(After, Len_Before + OldLength)
End Function

'/// PRIVATE INTERFACE ////////////////////////////////////////////////////////////////

'FormatVariant : Convert Variants to Strings, including recursive array walking
'======================================================================================
'AddVar         | Some value of any type, including arrays
'AddArray       | The array of FormatData you want to put the converted Strings into
'======================================================================================
Private Sub my_FormatVariant( _
    ByRef AddVar As Variant, _
    ByRef AddArray() As bluString_FormatData _
)
    'We don't want to calculate array bounds on every loop
    Dim First As Long, Last As Long
    
    'Check the type of the parameter (Long, String, Double &c.)
    Dim AddVarType As VBA.VbVarType
    Let AddVarType = VarType(AddVar)
    
    'We'll prepare the format data (String, Pointer, Length) here _
     before copying it to the return array
    Dim DataToAdd As bluString_FormatData
    
    Dim i As Long
    Select Case AddVarType
        'We can just use the pointer of an existing String
        Case vbString
            '..........................................................................
            Let DataToAdd.Pointer = StrPtr(AddVar)
            Let DataToAdd.ByteCount = LenB(AddVar)
            GoSub AddPointer
            
        'These types can be converted to a String; there are APIs for doing String _
         conversion <vbforums.com/showthread.php?736407> but `CStr` is probably _
         already using this behind the scenes. Also, rolling your own is actually _
         super difficult: <serpentine.com/blog/2011/06/29/here-be-dragons>
        Case vbBoolean, vbByte, vbCurrency, vbDate, vbDecimal, _
             vbDouble, vbEmpty, vbError, vbInteger, vbLong, vbSingle
            '..........................................................................
            Let DataToAdd.Text = CStr(AddVar)
            GoSub AddString
            
        Case vbObject
            '..........................................................................
            'Use the contents of a bluString
            If TypeOf AddVar Is bluString Then
                Let DataToAdd.Pointer = AddVar.Pointer
                Let DataToAdd.ByteCount = AddVar.LengthBytes
                GoSub AddPointer
            Else
                'Other objects can't be co-erced
                Err.Raise (13)
            End If
            
        'Note that `vbArray` is combined with the data type inside the array, _
         including Variant too, i.e. `vbArray + vbVariant`
        Case Is >= vbArray
            '..........................................................................
            'Arrays need to be walked, pushing each value as its own parameter
            Let First = LBound(AddVar): Let Last = UBound(AddVar)
            For i = First To Last: Call my_FormatVariant(AddVar(i), AddArray): Next i
            
        Case Else
            '..........................................................................
            'All other data types can't be co-erced
            Err.Raise (13)
    End Select
    Exit Sub
    
    '----------------------------------------------------------------------------------
AddPointer:
    Let DataToAdd.Text = vbNullString
    GoTo AddData
    
AddString:
    With DataToAdd
        Let .Pointer = StrPtr(.Text)
        Let .ByteCount = LenB(.Text)
    End With
    
AddData:
    Dim Count As Long
    Let Count = UBound(AddArray) + 1
    ReDim Preserve AddArray(0 To Count) As bluString_FormatData
    LSet AddArray(Count) = DataToAdd
    Return
End Sub

'SetLength : An internal method to set the string length without repeated validation
'======================================================================================
'NewLength      | What to set our string length to
'======================================================================================
Private Sub my_SetLength( _
    ByRef NewLength As Long _
)   'NO VALIDATION! INTERNAL USE ONLY!
    
    'Store the new desired content length (1-based)
    Let My_Length = NewLength: Let My_LengthB = My_Length * 2
    
    'If this increases the data too large...
    If My_Length >= My_Buffer Then
        'If the buffer has not been initialised, start with the default length. _
         We skip ahead as we need to test the new buffer size before increasing it
        If My_Buffer = 0 Then Let My_Buffer = DEFAULT_BUFFER: GoTo DefaultBuffer
        
        'Try increasing the buffer...
        Do
            'If the buffer is already at maximum then we've hit the absolute limit _
             (VB Error #14 is "Out of String Space")
            If My_Buffer = BUFFER_MAX Then Err.Raise (14): Exit Sub
            
            'In the case of custom buffer sizes, doubling the buffer might go over the _
             limit, but there may still be enough room for the new length if we just _
             stop at the buffer limit
            If (My_Buffer * 2) > BUFFER_MAX Then
                Let My_Buffer = BUFFER_MAX
            Else
                'Double the length of the buffer
                Let My_Buffer = My_Buffer * 2
            End If
            
DefaultBuffer:
        'Re-check if this is large enough
        Loop While My_Length > My_Buffer
        
        'Resize the data store. Note that though `My_Buffer` is 1-based we abuse this _
         to provide room for a null-terminator at the end of the string
        ReDim Preserve Chars(0 To My_Buffer) As Integer
        'Ensure it's null-terminated!
        Let Chars(My_Buffer) = 0
        'Note that the memory has been reserved
        Let Reserved = 1
    End If
    
    'Ensure that our string is null-terminated!
    Let Chars(My_Length) = 0
End Sub

'UnsafeCopy : Move a portion of our string to another place in our string
'======================================================================================
'IndexFrom      | Position within string (0-based) to begin copying from
'IndexTo        | As above, but where to copy to
'Length         | Number of characters (not Bytes) to copy
'======================================================================================
Private Sub my_UnsafeCopy( _
    ByRef IndexFrom As Long, _
    ByRef IndexTo As Long, _
    ByRef Length As Long _
)   'NO VALIDATION! INTERNAL USE ONLY!
    Call Win32.MemoryCopyWithin( _
        VarPtr(Chars(IndexTo)), VarPtr(Chars(IndexFrom)), Length * 2 _
    )
End Sub

'UnsafeCopyString : Copies a VB String into our string
'======================================================================================
'Text           | VB String to copy into our string
'Index          | Position within our string (0-based) to begin copying
'======================================================================================
Private Sub my_UnsafeCopyString( _
    ByRef Text As String, _
    ByRef Index As Long _
)   'NO VALIDATION! INTERNAL USE ONLY!
    '(This will overwrite anything already in place, so it's assumed you've already _
      handled resizing the string and copying things out)
    Call Win32.MemoryCopyBetween( _
        LenB(Text), VarPtr(Chars(Index)), StrPtr(Text) _
    )
End Sub

'ValidateChar : Validates a Unicode character code
'======================================================================================
'Char           | A Unicode character code (0 to 65535) or a negative number
'               | representing the signed-Integer equivilent of codes 32768 to 65535
'---------------¦----------------------------------------------------------------------
'Returns        | False for an invalid Char, True for a valid one
'======================================================================================
Private Function my_ValidateChar( _
    ByRef Char As UCS2 _
) As Boolean
    'If the character code is negative, it might be from a signed-Integer
    If Char < 0 Then
        'If it's in range, convert it to a positive number
        If Char >= -32768 Then
            Let Char = Char + 65535
            Let my_ValidateChar = True
        End If
    Else
        'Maximum of $FFFF
        If Char < 65536 Then Let my_ValidateChar = True
    End If
End Function

'ValidateIndex : Validates an index for the string
'======================================================================================
'Index          | Position within string (0-based)
'---------------¦----------------------------------------------------------------------
'Returns        | False for an invalid Index, True for a valid one
'======================================================================================
Private Function my_ValidateIndex( _
    ByRef Index As Long _
) As Boolean
    'Index cannot be less than zero
    If (Index >= 0) Then
        'Index cannot be beyond the end of the string _
        (note that the index is 0-based and the length is 1-based)
        If (Index < My_Length) Then Let my_ValidateIndex = True
    End If
End Function

'ValidateIndexAllowNegative : A negative index gives 'x' number of chars from the end
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is updated to be from the end of the string working backwards
'---------------¦----------------------------------------------------------------------
'Returns        | False for an invalid Index, True for a valid one
'======================================================================================
Private Function my_ValidateIndexAllowNegative( _
    ByRef Index As Long _
) As Boolean
    'Allow null string
    If (My_Length = 0) And (Index = 0) Then
        Let my_ValidateIndexAllowNegative = True
        Exit Function
    End If
    
    'Is the Index negative?
    If Index < 0 Then
        'The negative position cannot be more than the length of our string
        If Abs(Index) < My_Length Then
            'Update the Index to work backwards from the end of the string
            Let Index = My_Length + Index
            Let my_ValidateIndexAllowNegative = True
        End If
    Else
        'Index cannot be beyond the end of the string _
        (note that the index is 0-based and the length is 1-based)
        If (My_Length > Index) Then Let my_ValidateIndexAllowNegative = True
    End If
End Function

'ValidateRange : Validate an Index and Length fit within our string
'======================================================================================
'Index          | Position within our string (0-based). Can be negative, in which case
'               | the Index is calculated from the end of the string working backwards
'Length         | Length of the string portion, 0 = remainder of string from Index
'---------------¦----------------------------------------------------------------------
'Returns        | True if the range is valid (within our string), otherwise False
'======================================================================================
Private Function my_ValidateRange( _
    ByRef Index As Long, _
    ByRef Length As Long _
) As Boolean
    'Validate the Index (and convert negatives)
    If my_ValidateIndexAllowNegative(Index) Then
        'If Length is zero, then calculate the remainder of the string
        If Length = 0 Then Let Length = My_Length - Index
        'Length cannot be negative
        If Length >= 0 Then
            'The range cannot hang over the end of our string
            If (Index + Length) > My_Length Then
                'Clip the range to the end of our string
                Let Length = My_Length - Index
            End If
            'The Length must not be longer than the remainder of the string
            If (My_Length - Index - Length) >= 0 Then
                Let my_ValidateRange = True
            End If
        End If
    End If
End Function
