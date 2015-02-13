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

'Dependencies       bluArrayStrings.cls, <bluArray.cls>
'Last Updated       15-JAN-15
'Last Update        `Left` can now pad the string

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

'This enum is provided to make working with Character Codes more readable and obvious. _
 The data for this was taken from <danshort.com/unicode>
Public Enum UCS2
'   Null                        "Start Of Heading"          "Start Of Text"
:   C0_NUL:                     C0_SOH:                     C0_STX
'   "End of TeXt"               "End Of Transmission"       "Enquiry"
:   C0_ETX:                     C0_EOT:                     C0_ENQ
'   "Acknowledge"               "Bell"                      "Backspace"
:   C0_ACK:                     C0_BEL:                     C0_BS
'   "Horizontal Tab"            "Line Feed"                 "Vertical Tab"
:   C0_HT:                      C0_LF:                      C0_VT
'   "Form Feed"                 "Carriage Return"           "Shift Out"
:   C0_FF:                      C0_CR:                      C0_SO
'   "Shift In"                  "Data Link Escape"          "Device Control One"
:   C0_SI:                      C0_DLE:                     C0_DC1
'   "Device Control Two"        "Device Control Three"      "Device Control Four"
:   C0_DC2:                     C0_DC3:                     C0_DC4
'   "Negative Acknowledge"      "Synchronous Idle"          "End of Transmission Block"
:   C0_NAK:                     C0_SYN:                     C0_ETB
'   "Cancel"                    "End of Medium"             "Substitute"
:   C0_CAN:                     C0_EM:                      C0_SUB
'   "Escape"                    "File Separator"            "Group Separator"
:   C0_ESC:                     C0_FS:                      C0_GS
'   "Record Separator"          "Unit Separator"
:   C0_RS:                      C0_US
    
'                               !                           "
:   SPACE_:                     EXCLAMATION_MARK:           QUOTATION_MARK
'   #                           $                           %
:   NUMBER_SIGN:                DOLLAR_SIGN:                PERCENT_SIGN
'   &                           '                           (
:   AMPERSAND:                  APOSTROPHE:                 LEFT_PARENTHESIS
'   )                           *                           +
:   RIGHT_PARENHESIS:           ASTERISK:                   PLUS_SIGN
'   ,                           -                           .
:   COMMA:                      HYPHEN_MINUS:               FULL_STOP
'   /                           0                           1
:   SOLIDUS:                    DIGIT_ZERO:                 DIGIT_ONE
'   2                           3                           4
:   DIGIT_TWO:                  DIGIT_THREE:                DIGIT_FOUR
'   5                           6                           7
:   DIGIT_FIVE:                 DIGIT_SIX:                  DIGIT_SEVEN
'   8                           9                           :
:   DIGIT_EIGHT:                DIGIT_NINE:                 COLON
'   ;                           <                           =
:   SEMICOLON:                  LESS_THAN_SIGN:             EQUALS_SIGN
'   >                           ?                           @
:   GREATER_THAN_SIGN:          QUESTION_MARK:              COMMERCIAL_AT
    
:   LATIN_CAPITAL_LETTER_A:     LATIN_CAPITAL_LETTER_B:     LATIN_CAPITAL_LETTER_C
:   LATIN_CAPITAL_LETTER_D:     LATIN_CAPITAL_LETTER_E:     LATIN_CAPITAL_LETTER_F
:   LATIN_CAPITAL_LETTER_G:     LATIN_CAPITAL_LETTER_H:     LATIN_CAPITAL_LETTER_I
:   LATIN_CAPITAL_LETTER_J:     LATIN_CAPITAL_LETTER_K:     LATIN_CAPITAL_LETTER_L
:   LATIN_CAPITAL_LETTER_M:     LATIN_CAPITAL_LETTER_N:     LATIN_CAPITAL_LETTER_O
:   LATIN_CAPITAL_LETTER_P:     LATIN_CAPITAL_LETTER_Q:     LATIN_CAPITAL_LETTER_R
:   LATIN_CAPITAL_LETTER_S:     LATIN_CAPITAL_LETTER_T:     LATIN_CAPITAL_LETTER_U
:   LATIN_CAPITAL_LETTER_V:     LATIN_CAPITAL_LETTER_W:     LATIN_CAPITAL_LETTER_X
:   LATIN_CAPITAL_LETTER_Y:     LATIN_CAPITAL_LETTER_Z
    
'   [                           \                           ]
:   LEFT_SQUARE_BRACKET:        REVERSE_SOLIDUS:            RIGHT_SQUARE_BRACKET
'   ^                           _                           `
:   CIRCUMFLEX_ACCENT:          LOW_LINE:                   GRAVE_ACCENT
    
:   LATIN_SMALL_LETTER_A:       LATIN_SMALL_LETTER_B:       LATIN_SMALL_LETTER_C
:   LATIN_SMALL_LETTER_D:       LATIN_SMALL_LETTER_E:       LATIN_SMALL_LETTER_F
:   LATIN_SMALL_LETTER_G:       LATIN_SMALL_LETTER_H:       LATIN_SMALL_LETTER_I
:   LATIN_SMALL_LETTER_J:       LATIN_SMALL_LETTER_K:       LATIN_SMALL_LETTER_L
:   LATIN_SMALL_LETTER_M:       LATIN_SMALL_LETTER_N:       LATIN_SMALL_LETTER_O
:   LATIN_SMALL_LETTER_P:       LATIN_SMALL_LETTER_Q:       LATIN_SMALL_LETTER_R
:   LATIN_SMALL_LETTER_S:       LATIN_SMALL_LETTER_T:       LATIN_SMALL_LETTER_U
:   LATIN_SMALL_LETTER_V:       LATIN_SMALL_LETTER_W:       LATIN_SMALL_LETTER_X
:   LATIN_SMALL_LETTER_Y:       LATIN_SMALL_LETTER_Z

'   {                           |                           }
:   LEFT_CURLY_BRACKET:         VERTICAL_LINE:              RIGHT_CURLY_BRACKET
'   ~                           "Delete"
:   TILDE:                      C0_DEL
End Enum

'Memory manipulation: _
 --------------------------------------------------------------------------------------

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
    
    'Shorthand for "alpha-numeric"
    C1_ALPHANUM = C1_ALPHA Or C1_DIGIT
    'All kinds of blank characters you would want to strip off the ends
    C1_WHITESPACE = C1_SPACE Or C1_BLANK Or C1_CNTRL
    'Visible ("Printable") characters, this includes spaces, tabs &c.
    C1_VISIBLE = C1_SPACE Or C1_PUNCT Or C1_BLANK Or C1_ALPHANUM
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

'/// DATA /////////////////////////////////////////////////////////////////////////////

'We'll cache the Windows version here as it's very slow to do this every time _
 we call methods that use alternative APIs for Vista+
Private My_WindowsVersion As Single

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

'The marker used for formatting strings. As a constant here, it saves VB allocating _
 and deallocating a temporary String everytime we mention it
Private Const FORMAT_MARKER As String = "{$}"

'Case-sensitivity in methods such as `Format` and `LCase`/`UCase` is controlled _
 across the class by a property
Private My_CaseSensitive As Long

'Stuff for doing CRCs: _
 --------------------------------------------------------------------------------------
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

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    'NOTE: The buffer is not initialised to begin with _
          (therefore not wasting memory for a string that is not yet populated, _
           but more so that the reserved memory does not have to be reallocated _
           if a String larger than the default buffer is created first)
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    'Free up the reserved memory
    Erase Data, CRCTable
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Append : Add another string to the end of the current string
'======================================================================================
'Text           | The String to be appended
'---------------+----------------------------------------------------------------------
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
    'Increase the length accordingly _
     (this will also invalidate the cached CRC)
    Let Me.Length = My_Length + Len(Text)
    'Copy the new string in
    Call UnsafeCopyString(Text, Start)
End Function

'Buffer : Number of characters to reserve in memory for the string
'======================================================================================
Public Property Get Buffer() As BUFFER_SIZE: Let Buffer = My_Buffer: End Property
Public Property Let Buffer( _
    ByRef Length As BUFFER_SIZE _
)
    'No change in buffer?
    If Length = My_Buffer Then Exit Property
    
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
        'Changes to our string will invalidate the CRC
        Let CRCValid = 0
    End If
    'Resize the data store. Note that though `My_Buffer` is 1-based we abuse this _
     to provide room for a null-terminator at the end of the string
    ReDim Preserve Data(0 To My_Buffer) As Integer
    'Ensure it's null-terminated!
    Let Data(My_Buffer) = 0
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
        Let Char = (Data(My_Length + Index) And &HFFFF&)
    Else
        'Return the character value in question _
         The user can use `And &HFF&` to get ASCII
        Let Char = (Data(Index) And &HFFFF&)
    End If
End Property
Public Property Let Char( _
    ByVal Index As Long, _
    ByVal Char As UCS2 _
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
    
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Property

'CharInsert : Inserts a single Unicode character code
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'Char           | A Unicode character code (0 to 65535)
'---------------+----------------------------------------------------------------------
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
    If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Function
    'Validate the character code. It could be negative if it came from a VB _
     signed-Integer, so this is converted to the positive equivilent
    If Not ValidateChar(Char) Then Err.Raise (9): Exit Function
    
    'Increase the string length by 1 to make room _
     (this will also invalidate the cached CRC)
    Let Me.Length = My_Length + 1
    'Shift the upper portion of our string up
    Call UnsafeCopy(Index, Index + 1, My_Length - Index - 1)
    'VB6 can only do signed Integers, so Unicode points 32'768+ will error unless _
         we convert them to their signed equivilent to stuff them into two bytes
    If Char >= 32768 _
        Then Let Data(Index) = (Char And &HFFFF&) - 65536 _
        Else Let Data(Index) = (Char And &HFFFF&)
End Function

'CharIsAlpha : Returns true if the character at the index is a letter
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'[IsASCII]      | If True, only A-Z will validate, accented characters etc. won't
'---------------+----------------------------------------------------------------------
'Returns        | True if the character is alpha (a-z, but also non-latin characters)
'======================================================================================
Public Property Get CharIsAlpha( _
             ByVal Index As Long, _
    Optional ByRef IsASCII As Boolean = False _
) As Boolean
    'Whilst Unicode recognises a wider range of "alpha" characters, it is necessary _
     to sometimes limit ourselves to just A-Z
    If IsASCII = 0 Then
        'Note that this uses Unicode properties, so covers far more than just A-Z, _
         including accented letters and non-latin alpha letters such as "ß"
        Let CharIsAlpha = (Me.CharProperties(Index) And C1_ALPHA) <> 0
    Else
        'Check the character index is within range, _
         a negative Index will work backwards from the end of the string
        If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Property
        'Is the character code within A-Z? (upper and lower case)
        If Data(Index) >= 97 And Data(Index) <= 122 Then Let CharIsAlpha = True: Exit Property
        If Data(Index) >= 65 And Data(Index) <= 90 Then Let CharIsAlpha = True
    End If
End Property

'CharIsAlphaNumeric : Returns true if the character at the index is a letter or number
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'[IsASCII]      | If True, only A-Z&0-9 will validate, accented characters etc. won't
'---------------+----------------------------------------------------------------------
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
        If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Property
        'Is the character code within A-Z? (upper and lower case)
        If Data(Index) >= 97 And Data(Index) <= 122 Then Let CharIsAlphaNumeric = True: Exit Property
        If Data(Index) >= 65 And Data(Index) <= 90 Then Let CharIsAlphaNumeric = True: Exit Property
        'Is the character code within 0-9?
        If Data(Index) >= 48 And Data(Index) <= 57 Then Let CharIsAlphaNumeric = True
    End If
End Property

'CharIsHex : Returns true if the character at the index is a hexadecimal digit
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'---------------+----------------------------------------------------------------------
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
'---------------+----------------------------------------------------------------------
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
        If Data(Index) >= 48 And Data(Index) <= 57 Then Let CharIsNumeric = True
    End If
End Property

'CharIsLower : Returns true if the character at the index is lower-case
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'---------------+----------------------------------------------------------------------
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
'---------------+----------------------------------------------------------------------
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
'---------------+----------------------------------------------------------------------
'Returns        | True if the character is upper-case
'======================================================================================
Public Property Get CharIsUpper( _
    ByVal Index As Long _
) As Boolean
    'Note that this uses Unicode properties, so covers far more than just A-Z
    Let CharIsUpper = (Me.CharProperties(Index) And C1_UPPER) <> 0
End Property

'CharIsVisible : Returns true if the character at the index is visible
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'---------------+----------------------------------------------------------------------
'Returns        | True if the character is visible
'======================================================================================
Public Property Get CharIsVisible( _
    ByVal Index As Long _
) As Boolean
    'An "invisible" character is either a Control Code (such as Null), _
     or an undefined Unicode point. Zero-width space, whilst technically 'invisible' _
     in a proportional font still counts as "visible" as it's a defined character
    Let CharIsVisible = (Me.CharProperties(Index) And C1_VISIBLE) <> 0
End Property

'CharProperties : Returns the Unicode properties for a character in the string
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'---------------+----------------------------------------------------------------------
'Returns        | A bit battern that describes the Unicode properties of the character
'               | as defined in the C1 enum
'======================================================================================
Public Property Get CharProperties( _
    ByVal Index As Long _
) As C1
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Property
    'Use the API to fill the return value with the character properties
    Call api_GetStringTypeW( _
        CT_CTYPE1, UnsafePointer(Index), 1, _
        VarPtr(CharProperties) _
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
    Let CharPull = (Data(My_Length - 1) And &HFFFF&)
    'Reduce the size of the string _
     (this will also invalidate the cached CRC)
    Let Me.Length = My_Length - 1
End Function

'CharPush : Appends a single character code to the end of the string
'======================================================================================
'Char           | Unicode character point (0 to 65535)
'---------------+----------------------------------------------------------------------
'Returns        | This class object (method chaining)
'======================================================================================
Public Function CharPush( _
    ByVal Char As UCS2 _
) As bluString
    'Allow method chaining
    Set CharPush = Me
    'Validate the character code. It could be negative if it came from a VB _
     signed-Integer, so this is converted to the positive equivilent
    If Not ValidateChar(Char) Then Err.Raise (9): Exit Function
    
    'Increase the length of our string by one character _
     (the buffer will manage any necessary growth. _
      this will also invalidate the cached CRC)
    Let Me.Length = My_Length + 1
    'Set the character code of the last character
    'VB6 can only do signed Integers, so Unicode points 32'768+ will error unless _
         we convert them to their signed equivilent to stuff them into two bytes
    If Char >= 32768 _
        Then Let Data(My_Length - 1) = (Char And &HFFFF&) - 65536 _
        Else Let Data(My_Length - 1) = (Char And &HFFFF&)
End Function

'CharRemove : Removes a single character from the string and returns its code
'======================================================================================
'[Index]        | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'---------------+----------------------------------------------------------------------
'Returns        | A Unicode character point (0 to 65535)
'======================================================================================
Public Function CharRemove( _
    Optional ByVal Index As Long = 0 _
) As UCS2
    'Check the character index is within range, _
     a negative Index will work backwards from the end of the string
    If Not ValidateIndexAllowNegative(Index) Then Err.Raise (9): Exit Function
    
    'Return the character value
    Let CharRemove = Data(Index)
    
    'If removing just the last character, we won't need to shift the data down
    If Index < (My_Length - 1) Then
        'Shift the remainder of the string down a character
        Call UnsafeCopy(Index + 1, Index, My_Length - Index - 1)
    End If
    
    'Reduce the overal length of the string by one _
     (this will also invalidate the cached CRC)
    Let Me.Length = My_Length - 1
End Function

'Clear : Wipe the stored string, fill the buffer with nulls
'======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function Clear() As bluString
    'Note that we clear the buffer, and not just the old string content. _
     This is so that if the length is manually increased, junk data won't remain

    'Allow method chaining
    Set Clear = Me
    'You can't clear a null-string any more than it is!
    If My_Length > 0 Then
        'This does not change the size of the buffer, should it have increased. _
         If you intend to reuse a large string for a small string, create a new _
         bluString instance instead or call the `Destroy` method
        Call api_RtlZeroMemory(UnsafePointer, My_Buffer)
        'Mark the string as being null-length
        Let My_Length = 0: Let My_LengthB = 0
        'Changes to our string will invalidate the CRC
        Let CRCValid = 0
    End If
End Function

'Clone : Duplicate our string (or subportion) into another bluString instance
'======================================================================================
'[Index]        | Position within string (0-based). Can be negative, in which case the
'               | Index is calculated from the end of the string working backwards
'[Length]       | Number of characters to copy, can be 0 which will imply the
'               | remainder of the string, from the Index
'---------------+----------------------------------------------------------------------
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
    If Not ValidateRange(Index, Length) Then Err.Raise (9): Exit Function
    
    'Set the destination size
    Let Clone.Length = Length
    'Copy the text across (if there is any)
    If My_Length > 0 Then Call api_RtlMoveMemory( _
        Clone.Pointer(Index), UnsafePointer, Length * 2 _
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
            Let Value = (My_CRC And &HFF) Xor (Data(i) And &HFF)
            Let My_CRC = ((My_CRC And &HFFFFFF00) \ &H100) And 16777215
            Let My_CRC = My_CRC Xor CRCTable(Value)
            'And the upper byte:
            Let Value = (My_CRC And &HFF) Xor ((Data(i) And &HFF00&) \ &H100)
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
    Erase Data
    'Reset the buffer length so that any new content will create a new, _
     appropriately-sized buffer
    Let My_Buffer = 0
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'Equals : A fast method to check if one string equals another
'======================================================================================
'Text           | VB String to compare our string against
'---------------+----------------------------------------------------------------------
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
    If My_CaseSensitive = 0 Then
        'Case sensitivity flags in the API is different for each operating-system
        Dim Flags As CompareString_Flags
        
        'Check for Windows Vista and above:
        If WindowsVersion >= 6 Then
            'Windows Vista and above have more capable case-comparisons
            If My_CaseSensitive = 0 Then Let Flags = LINGUISTIC_IGNORECASE
            'NOTE: The use of null on the first parameter causes Windows to use the _
             user's current locale, so we don't need to manually fetch the locale name
            Let Equals = (api_CompareStringEx( _
                0&, Flags, _
                StrPtr(Text), -1, _
                UnsafePointer, -1, _
                0&, 0&, 0& _
            ) = CSTR_EQUAL)
        Else
            'This is the simpler case-comparison supported by XP
            If My_CaseSensitive = 0 Then Let Flags = NORM_IGNORECASE
            'Fall back to the XP-compatible API
            Let Equals = (api_CompareString( _
                api_GetThreadLocale(), Flags, _
                StrPtr(Text), -1, _
                UnsafePointer, -1 _
            ) = CSTR_EQUAL)
        End If
    
    'Case-sensitive (binary) comparison
    Else
        Let Equals = (api_CompareStringOrdinal( _
            StrPtr(Text), -1, _
            UnsafePointer, -1, _
            API_FALSE _
        ) = CSTR_EQUAL)
    End If
End Function

'Format : Replace markers in the text with the values provided
'======================================================================================
'Arguments ...  | An array of values (strings, numbers &c.) to insert
'---------------+----------------------------------------------------------------------
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
     somewhere to store them so that we can get at their pointers. The bluArrayString _
     provides an easy means of managing the text, length and pointer for each.
    'There are APIs for doing String conversion <vbforums.com/showthread.php?736407> _
     but `CStr` is probably already using this behind the scenes. Also, rolling your _
     own String conversion is actually super difficult _
     <serpentine.com/blog/2011/06/29/here-be-dragons>
    Dim Strings As bluArrayStrings
    Set Strings = New bluArrayStrings
    Let Strings.AllowDuplicates = True
    
    'Walk the params provided, converting them to Strings.
    Dim Param As Long
    For Param = First To Last
        Call Strings.AddVariant(Params(Param))
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
    Let Difference = Strings.Length(Param) - Len(FORMAT_MARKER)
    
    'Search for the format marker
    Dim ResultPointer As Long
    Let ResultPointer = api_StrStr(StartPointer, MarkerPointer)
    
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
        Call api_RtlMoveMemory( _
            ResultPointer, _
            Strings.Pointer(Param), Strings.LengthBytes(Param) _
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
            Let Format.Length = Format.Length + Difference
            'Get the potentially new string pointer
            Let StartPointer = Format.Pointer()
            Let ResultPointer = StartPointer + IndexB
        End If
        
        'Shift up or down the remainder portion of our string
        Call api_RtlMoveMemory( _
            ResultPointer + Strings.LengthBytes(Param), _
            ResultPointer + MarkerLengthB, _
            Remainder _
        )
        'If the replacement string is shorter than the search string, _
         our string will get shorter overall
        'WARNING: If the string increases beyond the buffer, _
         it will be reallocated, making `StartPointer` invalid!
        If Difference < 0 Then
            Let Format.Length = Format.Length + Difference
            'Get the potentially new string pointer
            Let StartPointer = Format.Pointer()
            Let ResultPointer = StartPointer + IndexB
        End If
        
        'Copy the replacement string in
        Call api_RtlMoveMemory( _
            ResultPointer, Strings.Pointer(Param), Strings.LengthBytes(Param) _
        )
    End If
    
    'We don't want an infinite loop, so move our search pointer over the replacement
    Let StartPointer = ResultPointer + Strings.LengthBytes(Param)
    'Move onto the next parameter
    Let Param = Param + 1: If Param > Strings.Count Then GoTo Finish
    'Proceed with the next search
    GoTo Continue
    
    '----------------------------------------------------------------------------------
Finish:
    'And clean up
    Set Strings = Nothing
End Function

'GetArray : Returns a byte-array copy of our string
'======================================================================================
Public Function GetArray( _
) As Integer()
    Dim ReturnArray() As Integer
    
    'Null-string won't return anything!
    If My_Length = 0 Then Exit Function
    'Size the return array to fit our string content (not the buffer)
    ReDim ReturnArray(0 To My_Length - 1) As Integer
    'Copy our string into the return array
    Call api_RtlMoveMemory(VarPtr(ReturnArray(0)), VarPtr(Data(0)), My_LengthB - 2)
    
    Let GetArray = ReturnArray
End Function

'Insert : Insert some text in the middle of the existing string
'======================================================================================
'Text           | VB String to insert into our string
'Index          | Position within our string (0-based). Can be negative, in which case
'               | the Index is calculated from the end of the string working backwards
'---------------+----------------------------------------------------------------------
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
     (the buffer will be increased automatically if necessary. _
      this will also invalidate the cached CRC)
    Let Me.Length = My_Length + Len(Text)
    'Shift the upper portion of our string up to make room for the inserted string
    Call UnsafeCopy(Index, Index + Len(Text), My_Length - Index)
    'Insert the new string
    Call UnsafeCopyString(Text, Index)
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
'---------------+----------------------------------------------------------------------
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
    'Increase the length accordingly _
     (this will also invalidate the cached CRC)
    Let Me.Length = My_Length + bluStringObj.Length
    'Copy the new string in
    Call api_RtlMoveMemory( _
        UnsafePointer(Start), bluStringObj.Pointer, bluStringObj.Length * 2 _
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
    
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'Left : Cut the string short to a given number of characters
'======================================================================================
'Length         | Number of characters to keep of the string. If negative, this will
'               | be calculated backwards from the end of the string
'[FillChar]     | If the new Length is longer than our string, the gap will be
'               | filled with the given character, which defaults to Space
'---------------+----------------------------------------------------------------------
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
        For i = OldLength To My_Length - 1: Let Data(i) = FillChar: Next i
    
    'If the position is negative, work out 'x' number of characters before the end of _
     our string. That is, -1 will be the second to last character
    ElseIf Length < 0 Then
        'The negative position cannot be more than the length of our string
        If Abs(Length) > My_Length Then Err.Raise (9): Exit Function
        'Calculate the new position from the end of our string _
         (this will also invalidate the cached CRC)
        Let Me.Length = My_Length + Length
    Else
        'Do the reduction _
         (this will also invalidate the cached CRC)
        Let Me.Length = Length
    End If
    
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'Length : Length of the actual stored text (1-based), not the buffer
'======================================================================================
Public Property Get Length() As Long: Let Length = My_Length: End Property
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
    Let My_Length = NewLength: Let My_LengthB = My_Length * 2
    'If the buffer has not been initialised, start with the default length
    Dim NewBuffer As Long
    If My_Buffer = 0 _
        Then Let NewBuffer = DEFAULT_BUFFER _
        Else Let NewBuffer = My_Buffer
    
TestBuffer:
    'If this increases the data too large...
    If My_Length > NewBuffer Then
        'If the buffer is already at maximum then we've hit the absolute limit _
         (VB Error #14 is "Out of String Space")
        If NewBuffer = BUFFER_MAX Then Err.Raise (14): Exit Property
        'In the case of custom buffer sizes, doubling the buffer might go over the _
         limit, but there may still be enough room for the new length if we just stop _
         at the buffer limit
        If (NewBuffer * 2) > BUFFER_MAX Then
            Let NewBuffer = BUFFER_MAX
        Else
            'Double the length of the buffer
            Let NewBuffer = NewBuffer * 2
        End If
        'Re-check if this is large enough
        GoTo TestBuffer
    End If
    'Apply the new buffer length _
     (this will do the actual memory allocation)
    Let Me.Buffer = NewBuffer
    
    'WARNING: Increasing the content-length via this property, instead of by-proxy _
     via methods such as `Append` may leave 'junk' data in the increased space!
    Let Data(My_Length) = 0
    
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Property

'LengthBytes : Length of the content in bytes, cannot be set
'======================================================================================
Public Property Get LengthBytes() As Long: Let LengthBytes = My_LengthB: End Property

'Mid : Reduce the string to a sub-string portion
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case
'               | the Index is calculated from the end of the string working backwards
'[Length]       | Number of characters to keep, can be 0 which will imply the
'               | remainder of the string, from the Index
'---------------+----------------------------------------------------------------------
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
    If Not ValidateRange(Index, Length) Then Err.Raise (9): Exit Function
    
    'Copy the sub-string to the beginning of our string
    Call UnsafeCopy(Index, 0, Length)
    'Reduce the string to the new sub-string length _
     (this will also invalidate the cached CRC)
    Let Me.Length = Length
End Function

'Overwrite : Overwrite a portion of our string with text
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case
'               | the Index is calculated from the end of the string working backwards
'Text           | New text to overwrite our string at the given Index
'---------------+----------------------------------------------------------------------
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
    If Not ValidateRange(Index, Length) Then Err.Raise (9): Exit Function
    
    'Now just copy the new text over the top of our string
    Call api_RtlMoveMemory(UnsafePointer(Index), StrPtr(Text), Length * 2)
End Function

'Pointer : Allow access to the "string" pointer for external copying
'======================================================================================
'[Index]        | Position within our string (0-based)
'---------------+----------------------------------------------------------------------
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

'Prepend : Add our string to the end of the one provided
'======================================================================================
'Text           | VB String to prepend to our string
'---------------+----------------------------------------------------------------------
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
    'Increase the length accordingly _
     (this will also invalidate the cached CRC)
    Let Me.Length = My_Length + Start
    'Shift our string up
    Call UnsafeCopy(0, Start, My_Length)
    'Copy the new string in
    Call UnsafeCopyString(Text, 0)
End Function

'ReadFile : Import our string from a file
'======================================================================================
'FilePath       | Path to the file to read
'---------------+----------------------------------------------------------------------
'Returns        | This class object (method chaining)
'======================================================================================
Public Function ReadFile( _
    ByRef FilePath As String _
) As bluString
    'Allow method chaining
    Set ReadFile = Me
    
    Dim FileNumber As Integer
    Let FileNumber = FreeFile
    
    'At the moment this is the simplest method of reading a file of unknown _
     encoding (ASNI / UTF-16 / UTF-8 &c.), however it wastes memory as it duplicates _
     the file in memory to do the conversion to UTF-16, and we duplicate it again _
     to hold it in a VB6 String before copying it to our bluString. I'd like to _
     rewrite this whole thing to read the file serially and convert-as-we-go. _
     This is a useful reference for manually reading a file encoding: _
     <stackoverflow.com/questions/15809081/encoding-of-text-files-in-vb-6-0>
    Open FilePath For Input Access Read Lock Write As #FileNumber
    'TODO: File not found, file too large and other such errors
    Dim Source As String
    Let Source = StrConv(InputB(LOF(FileNumber), FileNumber), vbUnicode)
    
    'Use our already-existing property for copying a VB6 string into ours
    Let Me.Text = Source
    
Finish:
    Close #FileNumber
    
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'ReduceBuffer : Free some memory by reducing the buffer to the string length
'======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function ReduceBuffer() As bluString
    'Of course you could just use `bluString.Buffer = bluString.Length`, but this is _
     intended to be more obvious and safer for potential future changes / expansion. _
     It's strongly advised NOT to expand the string after you've reduced the buffer

    'Allow method chaining
    Set ReduceBuffer = Me
    'Reduce the buffer down to the size of the string
    Let Me.Buffer = My_Length
End Function

'Remove : Cut some text out of the middle of the existing string
'======================================================================================
'Index          | Position within our string (0-based). Can be negative, in which case
'               | the Index is calculated from the end of the string working backwards
'[Length]       | Number of characters to remove, 0 = remainder of string from Index
'---------------+----------------------------------------------------------------------
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
    If Not ValidateRange(Index, Length) Then Err.Raise (9): Exit Function
    
    'If you end up removing the whole string, make it a null-string
    If (Index = 0) And (Length = My_Length - 1) Then Call Me.Destroy: Exit Function
    
    'Shift the remainder down into the removed area
    Call UnsafeCopy(Index + Length, Index, Length)
    'Update the string length _
     (this will also invalidate the CRC)
    Let Me.Length = My_Length - Length
End Function

'Replace : Search for a sub-string and replace it with another
'======================================================================================
'SearchFor      | Text to find within our string
'[ReplaceWith]  | What to replace the found text with, can be null-string
'[ReturnCount]  | If provided, your own variable will be set with the number of
'               | replacements done
'---------------+----------------------------------------------------------------------
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
    Let StartPointer = UnsafePointer()
    
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
    Let IndexB = ResultPointer - UnsafePointer()
    
    'If both the search string and the replacement are the same length, _
     we don't need to move parts of our string about
    If Difference = 0 Then
        'Replace the found string
        Call api_RtlMoveMemory(ResultPointer, StrPtr_ReplaceWith, LenB_ReplaceWith)
    Else
        'Calculate the remainder portion of our string that will be moved
        Dim Remainder As Long
        Let Remainder = My_LengthB - (IndexB + LenB_SearchFor)
        
        'If the replacement string is longer than the search string, _
         our string will get longer overall
        'WARNING: If the string increases beyond the buffer, _
         it will be reallocated, making `ResultPointer` invalid!
        If Difference > 0 Then
            Let Me.Length = My_Length + Difference
            'Get the potentially new string pointer
            Let StartPointer = UnsafePointer()
            Let ResultPointer = StartPointer + IndexB
        End If
        
        'Shift up or down the remainder portion of our string
        Call api_RtlMoveMemory( _
            ResultPointer + LenB_ReplaceWith, _
            ResultPointer + LenB_SearchFor, _
            Remainder _
        )
        'If the replacement string is shorter than the search string, _
         our string will get shorter overall
        'WARNING: If the string increases beyond the buffer, _
         it will be reallocated, making `StartPointer` invalid!
        If Difference < 0 Then
            Let Me.Length = My_Length + Difference
            'Get the potentially new string pointer
            Let StartPointer = UnsafePointer()
            Let ResultPointer = StartPointer + IndexB
        End If
        
        'Copy the replacement string in
        Call api_RtlMoveMemory( _
            ResultPointer, StrPtr_ReplaceWith, LenB_ReplaceWith _
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
'---------------+----------------------------------------------------------------------
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
    Call UnsafeCopy(My_Length - Length, 0, Length)
    'Cut the string short. The buffer will be left at its old length, _
     if you know the string won't expand any more you can free the memory up _
     by calling the `ReduceBuffer` method
    Let Me.Length = Length
    
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
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
        Let Text = Space$(My_Length)
        'Copy our String into the return field
        Call api_RtlMoveMemory(StrPtr(Text), UnsafePointer, My_LengthB)
    End If
End Property
Public Property Let Text( _
    ByRef NewString As String _
)   'If a null-string is provided then set ours to a null string
    If LenB(NewString) = 0 Then
        Call Me.Destroy
    Else
        'Completely free our existing string
        Call Me.Destroy
        'Set the new length of our string; _
         The buffer will be allocated appropriately
        Let Me.Length = Len(NewString)
        'Copy the VB String data over
        Call api_RtlMoveMemory(UnsafePointer, StrPtr(NewString), LenB(NewString))
    End If
    
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Property

''Trim : Strip whitespace characters from the beginning and end
''======================================================================================
''TODO: We may want a look-up table of all whitespace chars in UCS-2
'Public Function Trim() As bluString
'    'Allow method chaining
'    Set Trim = Me
'    'TODO
'End Function

'TrimLeft : Strip whitespace characters from the beginning
'======================================================================================
'Returns        | This class object (method chaining)
'======================================================================================
Public Function TrimLeft() As bluString
    'Allow method chaining
    Set TrimLeft = Me
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
    Call UnsafeCopy(Index, 0, My_Length - Index)
    'Update the string length _
     (this will also invalidate the CRC)
    Let Me.Length = My_Length - Index
End Function

''TrimRight : Strip whitespace characters from the end
''======================================================================================
'Public Function TrimRight() As bluString
'    'Allow method chaining
'    Set TrimRight = Me
'    'TODO
'End Function

'UCase : Convert our string to UPPERCASE
'======================================================================================
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
    
    'Changes to our string will invalidate the CRC
    Let CRCValid = 0
End Function

'Wrap : Append & Prepend our string, e.g. for adding quotes around text
'======================================================================================
'Before         | String to prepend to our string
'[After]        | String to append to our string. If ommitted, defaults to Before
'---------------+----------------------------------------------------------------------
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
    Let Me.Length = My_Length + Len_Before + Len_After
    'Shift our string up to account for the before _
     (assuming Before is not null-string)
    If Len_Before > 0 Then Call UnsafeCopy(0, Len_Before, OldLength)
    'Copy in the new before and after
    If Len_Before > 0 Then Call UnsafeCopyString(Before, 0)
    If Len_After > 0 Then Call UnsafeCopyString(After, Len_Before + OldLength)
End Function

''WriteFile : Export our string to a file
''======================================================================================
'Public Function WriteFile( _
'    ByRef FilePath As String _
') As bluString
'    'Allow method chaining
'    Set WriteFile = Me
'    'TODO
'End Function

'/// PRIVATE INTERFACE ////////////////////////////////////////////////////////////////

'UnsafePointer : Gets a pointer to the string, without validation -- internal use
'======================================================================================
'[Index]        | Position within our string (0-based)
'---------------+----------------------------------------------------------------------
'Returns        | Memory pointer to the character index within our string
'======================================================================================
Private Property Get UnsafePointer( _
    Optional ByRef Index As Long = 0 _
) As Long
    'Return a pointer to the given character index
    Let UnsafePointer = VarPtr(Data(Index))
End Property

'UnsafeCopy : Move a portion of our string to another place in our string
'======================================================================================
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

'UnsafeCopyString : Copies a VB String into our string
'======================================================================================
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

'ValidateChar : Validates a Unicode character code
'======================================================================================
'Char           | A Unicode character code (0 to 65535) or a negative number
'               | representing the signed-Integer equivilent of codes 32768 to 65535
'---------------+----------------------------------------------------------------------
'Returns        | False for an invalid Char, True for a valid one
'======================================================================================
Private Function ValidateChar( _
    ByRef Char As UCS2 _
) As Boolean
    'If the character code is negative, it might be from a signed-Integer
    If Char < 0 Then
        'If it's in range, convert it to a positive number
        If Char >= -32768 Then
            Let Char = Char + 65535
            Let ValidateChar = True
        End If
    Else
        'Maximum of $FFFF
        If Char < 65536 Then Let ValidateChar = True
    End If
End Function

'ValidateIndex : Validates an index for the string
'======================================================================================
'Index          | Position within string (0-based)
'---------------+----------------------------------------------------------------------
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

'ValidateIndexAllowNegative : A negative index gives 'x' number of chars from the end
'======================================================================================
'Index          | Position within string (0-based). Can be negative, in which case the
'               | Index is updated to be from the end of the string working backwards
'---------------+----------------------------------------------------------------------
'Returns        | False for an invalid Index, True for a valid one
'======================================================================================
Private Function ValidateIndexAllowNegative( _
    ByRef Index As Long _
) As Boolean
    'Allow null string
    If (My_Length = 0) And (Index = 0) Then
        Let ValidateIndexAllowNegative = True
        Exit Function
    End If
    
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
        If (My_Length > Index) Then Let ValidateIndexAllowNegative = True
    End If
End Function

'ValidateRange : Validate an Index and Length fit within our string
'======================================================================================
'Index          | Position within our string (0-based). Can be negative, in which case
'               | the Index is calculated from the end of the string working backwards
'Length         | Length of the string portion, 0 = remainder of string from Index
'---------------+----------------------------------------------------------------------
'Returns        | True if the range is valid (within our string), otherwise False
'======================================================================================
Private Function ValidateRange( _
    ByRef Index As Long, _
    ByRef Length As Long _
) As Boolean
    'Validate the Index (and convert negatives)
    If ValidateIndexAllowNegative(Index) Then
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
                Let ValidateRange = True
            End If
        End If
    End If
End Function

'WindowsVersion
'======================================================================================
'Returns        | Windows Kernel version number
'======================================================================================
Private Property Get WindowsVersion() As Single
    'For speed, we don't want to call the API every time, cache the value. _
     If it's not been cached yet, do so
    If My_WindowsVersion = 0 Then
        'Get the Windows version as Windows Vista and above use alternative APIs for _
         locale-dependent methods (such as `Equals`)
        Dim VersionInfo As OSVERSIONINFO
        Let VersionInfo.SizeOfMe = Len(VersionInfo)
        'NOTE: If the app is in compatibility mode, this will return the compatible _
         Windows version, not the actual version. Windows 8.1 and above also cause _
         this API to always return as Windows 8.0 unless the app opts-in to correct _
         version info via manifest, but this isn't an issue to bluString as we are _
         only concerned with testing for Windows Vista+
        If api_GetVersionEx(VersionInfo) = API_TRUE Then
            '`Val`, rather than `CSng` is used here due to locale differences; _
             `CSng("6.1")` will fail in locales that use comma as decimal separator
            Let My_WindowsVersion = Val( _
                VersionInfo.MajorVersion & "." & VersionInfo.MinorVersion _
            )
        End If
    End If
    'Return the cached value
    Let WindowsVersion = My_WindowsVersion
End Property
