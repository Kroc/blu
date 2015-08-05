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

'Last Updated       24-JUL-15
'Last Update        *   Moved APIs to the type library
'                   *   Added `UnicodeProperties` lookup table

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
:   DIGIT_EIGHT:                DIGIT_NINE:                 colon
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
:   tilde:                      C0_DEL
End Enum

Private Const STR_SPC As String = " "
Private Const STR_ELLIPSIS As String = "…"

Private My_UnicodeProperties(-32768 To 65535) As Integer

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    Dim i As Long
    
    'Create the lookup table of Unicode character properties
    For i = -32768 To 32767: Let My_UnicodeProperties(i) = i: Next i
    For i = 32768 To 65535: Let My_UnicodeProperties(i) = i - 65535: Next i
    'Lock the array so that once externally accessed, it can't be redimmed or erased
    Call Win32.ArrayLockAdd(ArrPtr(My_UnicodeProperties))
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    'Unlock the Unicode character properties lookup table; _
     since it's pre-dimmed, `Erase` will only zero it so leave it to drop out of scope
    Call Win32.ArrayLockRemove(ArrPtr(My_UnicodeProperties))
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
        Let CompareText = (Win32.StringCompare_NT5( _
                  LocaleID:=Win32.ThreadLocale, _
              CompareFlags:=w32IgnoreCase, _
            String1Pointer:=StrPtr(String1), String1Length:=-1, _
            String2Pointer:=StrPtr(String2), String2Length:=-1 _
        ) = Win32.StringCompareResult.w32Equal)
    Else
        'NOTE: The use of null on the first parameter causes Windows to use the _
         user's current locale, so we don't need to manually fetch the locale name
        Let CompareText = (Win32.StringCompare_NT6( _
             LocaleNamePointer:=0&, _
                  CompareFlags:=w32IgnoreCaseLinguistic, _
                String1Pointer:=StrPtr(String1), String1Length:=-1, _
                String2Pointer:=StrPtr(String2), String2Length:=-1 _
        ) = Win32.StringCompareResult.w32Equal)
    End If
End Function

''Concatenate : VB6's string concatenation is very slow, this should be faster
''======================================================================================
'Public Function Concatenate( _
'    ParamArray Strings() As Variant _
') As String
'    'If nothing is passed, return Null-string
'    If UBound(Strings) < 0 Then Exit Function
'
'
'End Function

'FormatText : Fit a piece of text to a fixed length with truncate / align / pad
'======================================================================================
'InputText      | The text you want to format
'OutputLength   | The fixed length of the output text that will be created; e.g. you
'               | want to format file names in an ASCII table and the length of the
'               | table cell must always be 32 chars (padded for shorter filenames)
'[Align]        | Alignment of the text within the given fixed size of output
'[Ellipsis]     | Input text longer than the fixed output length will be truncated
'               | using the ellipsis string (i.e. "...") to denote the overflow
'[FillChar]     | The "background" of the output, e.g. pad with spaces
'---------------¦----------------------------------------------------------------------
'Returns        | A new string, of the length specified
'======================================================================================
Public Function FormatText( _
             ByRef InputText As String, _
             ByRef OutputLength As Long, _
    Optional ByRef Align As VBRUN.AlignmentConstants = vbLeftJustify, _
    Optional ByRef Ellipsis As String = STR_ELLIPSIS, _
    Optional ByRef FillChar As String = STR_SPC _
) As String
    'Don't be silly
    If OutputLength < 0 Then Call Err.Raise(9): Exit Function
    
    'Should the input text match the desired output length, nothing needs to be done!
    If Len(InputText) = OutputLength Then
        Let FormatText = InputText
        Exit Function
    End If
    
    Dim OutputLengthBytes As Long
    Let OutputLengthBytes = OutputLength * 2
    
    'Prepate the output buffer (using the TLB-API here is faster than `Space$`!). _
     The string will currently contain garbage from previously used memory
    Let FormatText = Win32.StringCreateUninitialized(ByVal 0&, OutputLength)
    
    'cache a bunch of stuff we're going to need
    Dim OutputPtr As Long
    Let OutputPtr = StrPtr(FormatText)
    
    Dim EllipsisPtr As Long
    Let EllipsisPtr = StrPtr(Ellipsis)
    
    Dim EllipsisLen As Long
    Let EllipsisLen = LenB(Ellipsis)
    
    Dim SourcePtr As Long
    Let SourcePtr = StrPtr(InputText)
    
    Dim SourceLen As Long
    Let SourceLen = LenB(InputText)
    
    'If the input text is too long it needs to be truncated
    '----------------------------------------------------------------------------------
    If SourceLen > OutputLengthBytes Then
        'Left align, truncate on the right
        '..............................................................................
        If Align = vbLeftJustify Then
            'Write the ellipsis on the right hand side
            Call Win32.MemoryCopyWithin( _
                OutputPointer:=OutputPtr + OutputLengthBytes - EllipsisLen, _
                SourcePointer:=EllipsisPtr, _
                  LengthBytes:=EllipsisPtr _
            )
            'Fill the remainder with the input text
            Call Win32.MemoryCopyWithin( _
                OutputPointer:=OutputPtr, _
                SourcePointer:=SourcePtr, _
                  LengthBytes:=OutputLengthBytes - EllipsisLen _
            )
        
        'If the output is to be right-aligned, _
         then the truncation will be on the left
        '..............................................................................
        ElseIf Align = vbRightJustify Then
            'Write the ellipsis on the left hand side
            Call Win32.MemoryCopyBetween( _
                  LengthBytes:=EllipsisLen, _
                OutputPointer:=OutputPtr, _
                SourcePointer:=EllipsisLen _
            )
            'Fill the remainder with the input text (right aligned)
            Call Win32.MemoryCopyBetween( _
                  LengthBytes:=OutputLengthBytes - EllipsisLen, _
                OutputPointer:=(OutputPtr + EllipsisLen), _
                SourcePointer:=(SourcePtr + SourceLen) _
                              - OutputLengthBytes - EllipsisLen _
            )
        
        '..............................................................................
        Else
            'TODO
            
        
        End If
    
    'The input text is shorter than the output length, so we need to align & pad
    '----------------------------------------------------------------------------------
    Else
        'Fill the buffer with the given background character: We use an overlapping _
         memory copy to walk down the string, copying the character before us forward; _
         like laying down track from a moving train! My use of this technique is based _
         upon this post: <stackoverflow.com/a/3213694>. There is very likely no faster _
         way of doing this in VB6
        
        'We prepare the first two characters (32-bits), _
         as the memory copy will walk one Long at a time
        Let Win32.MemInteger(OutputPtr) = AscW(FillChar)
        Let Win32.MemInteger(OutputPtr + 2) = AscW(FillChar)
        
        'We do a memory copy within the same buffer (notice we are using our API that _
         *should* be used for copying between different buffers -- it doesn't check _
         for overlapped copies); this copies the previous four bytes to the next four _
         bytes, all the way down the string
        Call Win32.MemoryCopyBetween( _
              LengthBytes:=OutputLengthBytes - 4, _
            OutputPointer:=OutputPtr + 4, _
            SourcePointer:=OutputPtr _
        )
        
        'Select the text alignment:
        '..............................................................................
        If Align = vbLeftJustify Then
            'Place the text on the left
            Call Win32.MemoryCopyBetween( _
                  LengthBytes:=SourceLen, _
                OutputPointer:=OutputPtr, _
                SourcePointer:=SourcePtr _
            )
            
        '..............................................................................
        ElseIf Align = vbRightJustify Then
            'Place the text on the right
            Call Win32.MemoryCopyBetween( _
                  LengthBytes:=SourceLen, _
                OutputPointer:=OutputPtr + (OutputLengthBytes - SourceLen), _
                SourcePointer:=SourcePtr _
            )
            
        '..............................................................................
        Else
            'TODO
            
        End If
    End If
End Function

'UnicodeProperties
'======================================================================================
Public Property Get UnicodeProperties( _
) As Integer()
    Let UnicodeProperties = My_UnicodeProperties
End Property
