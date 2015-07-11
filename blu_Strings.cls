VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "blu_Strings"
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
'CLASS :: blu_Strings

'Dependencies       blu > blu_Sys
'Last Updated       12-JUL-15
'Last Update        Initial version - INCOMPLETE

'/// API //////////////////////////////////////////////////////////////////////////////

'In VB6 True is -1 and False is 0, but in the Win32 API it's 1 for True
Private Enum BOOL
    API_TRUE = 1
    API_FALSE = 0
End Enum

'Locale Mapping (for case conversion and string comparison):
'--------------------------------------------------------------------------------------
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
    ByVal ResultStringPointer As Long, _
    ByVal ResultStringLength As Long _
) As Long

'Unicode & Locale-aware case conversion (Windows Vista+) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318702(v=vs.85).aspx>
Private Declare Function api_LCMapStringEx Lib "kernel32" Alias "LCMapStringEx" ( _
    ByVal LocaleNamePointer As Long, _
    ByVal MapFlags As LCMAP, _
    ByVal SourceStringPointer As Long, _
    ByVal SourceStringLength As Long, _
    ByVal ResultStringPointer As Long, _
    ByVal ResultStringLength As Long, _
    ByVal VersionInfoPointer As Long, _
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

'/// CLASS ////////////////////////////////////////////////////////////////////////////

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

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'======================================================================================
Private Sub Class_Initialize()
    '
End Sub

'======================================================================================
Private Sub Class_Terminate()
    '
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'CompareText
'======================================================================================
Public Function CompareText( _
    ByRef String1 As String, _
    ByRef String2 As String _
) As Boolean
    'Check for Windows XP:
    If blu.sys.WindowsAPILevel = XP_OR_LOWER Then
        'Fall back to the XP-compatible API
        Let CompareText = (api_CompareString( _
                  LocaleID:=api_GetThreadLocale(), _
              CompareFlags:=NORM_IGNORECASE, _
            String1Pointer:=StrPtr(String1), String1Length:=-1, _
            String2Pointer:=StrPtr(String2), String2Length:=-1 _
        ) = CSTR_EQUAL)
    Else
        'NOTE: The use of null on the first parameter causes Windows to use the _
         user's current locale, so we don't need to manually fetch the locale name
        Let CompareText = (api_CompareStringEx( _
             LocaleNamePointer:=0&, _
                  CompareFlags:=LINGUISTIC_IGNORECASE, _
                String1Pointer:=StrPtr(String1), String1Length:=-1, _
                String2Pointer:=StrPtr(String2), String2Length:=-1, _
            VersionInfoPointer:=0&, _
               ReservedPointer:=0&, _
                         Param:=0& _
        ) = CSTR_EQUAL)
    End If
End Function