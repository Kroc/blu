Attribute VB_Name = "W32"
Option Explicit
'======================================================================================
'bluW32; Copyright (C) Kroc Camen, 2013-15
'MIT License:
'--------------------------------------------------------------------------------------
'Permission is hereby granted, free of charge, to any person obtaining a copy of this
'software and associated documentation files (the "Software"), to deal in the Software
'without restriction, including without limitation the rights to use, copy, modify,
'merge, publish, distribute, sublicense, and/or sell copies of the Software, and to
'permit persons to whom the Software is furnished to do so, subject to the following
'conditions:
'
'   *   The above copyright notice and this permission notice shall be
'       included in all copies or substantial portions of the Software
'
'   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'   IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
'   FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
'   COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
'   IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
'   WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

'======================================================================================

'MODULE :: W32

'This is a support module for bluW32 (WIN32 Type Library);
'We can't include Windows version-specific APIs in the type library otherwise your _
 compiled app will only work on the highest version of Windows included. This module _
 provides the Windows version-specific APIs so that your app can hapily work on XP+

'======================================================================================

'Locale Mapping (for case conversion and string comparison):
'---------------------------------------------------------------------------------------
'This page helped with working out correct methods of using Unicode API calls _
 <www.xtremevbtalk.com/showthread.php?t=68956>
'This page helped with translating the MSDN documentation into VB6; _
 <www.ex-designz.net/apidetail.asp?api_id=383>

'Get the Locale Identifier (LCID) of this app _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318127(v=vs.85).aspx>
'This is used for Windows XP support as Vista+ use Locale Name strings
Public Declare Function GetThreadLocale Lib "kernel32" ( _
) As Long

'Unicode & Locale-aware case conversion (Windows XP) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318700(v=vs.85).aspx>
Public Declare Function StringConvertCase_NT5 Lib "kernel32" Alias "LCMapStringW" ( _
    ByVal LocaleID As Long, _
    ByVal Flags As StringConvertCase, _
    ByVal SourceStringPointer As Long, _
    ByVal SourceStringLength As Long, _
    ByVal ResultStringPointer As Long, _
    ByVal ResultStringLength As Long _
) As Long

'Unicode & Locale-aware case conversion (Windows Vista+) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd318702(v=vs.85).aspx>
Public Declare Function StringConvertCase_NT6 Lib "kernel32" Alias "LCMapStringEx" ( _
    ByVal LocaleNamePointer As Long, _
    ByVal Flags As StringConvertCase, _
    ByVal SourceStringPointer As Long, _
    ByVal SourceStringLength As Long, _
    ByVal ResultStringPointer As Long, _
    ByVal ResultStringLength As Long, _
    ByVal Reserved_VersionInfo As Long, _
    ByVal Reserved As Long, _
    ByVal Reserved_SortHandle As Long _
) As Long

Public Enum StringConvertCase
    'Use byte reversal. For example, if the application passes in 0x3450 0x4822,
    ' the result is 0x5034 0x2248
    w32CaseReverseBytes = &H800&
    'Use Unicode (wide) characters where applicable.
    ' This flag and `w32CaseHalfWidth` are mutually exclusive
    w32CaseFullWidth = &H800000
    'Use narrow characters where applicable.
    ' This flag and `w32CaseFullWidth` are mutually exclusive
    w32CaseHalfWidth = &H400000
    'Map all katakana characters to hiragana.
    ' This flag and `w32CaseKatakana` are mutually exclusive
    w32CaseHiragana = &H100000
    'Map all hiragana characters to katakana.
    ' This flag and `w32CaseHirigana` are mutually exclusive
    w32CaseKatakana = &H200000
    'Use linguistic rules for casing, instead of file system rules (default)
    ' This flag is valid with `w32CaseLower` or `w32CaseUpper` only
    w32CaseLinguistic = &H1000000
    'For locales and scripts capable of handling uppercase and lowercase,
    ' map all characters to lowercase
    w32CaseLower = &H100&
    'Map traditional Chinese characters to simplified Chinese characters.
    ' This flag and `w32CaseChineseTraditional` are mutually exclusive
    w32CaseChineseSimplified = &H2000000
    'Windows 7: Map all characters to title case,
    ' in which the first letter of each major word is capitalized
    w32CaseTitle_WIN7 = &H300&
    'Map simplified Chinese characters to traditional Chinese characters.
    ' This flag and `w32CaseChineseSimplified` are mutually exclusive
    w32CaseChineseTraditional = &H4000000
    'For locales and scripts capable of handling uppercase and lowercase,
    ' map all characters to uppercase
    w32CaseUpper = &H200&
End Enum

'Comparing Strings:
'---------------------------------------------------------------------------------------

'Unicode & Locale-aware string comparison (Windows XP) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd317759(v=vs.85).aspx>
Public Declare Function StringCompare_NT5 Lib "kernel32" Alias "CompareStringW" ( _
    ByVal LocaleID As Long, _
    ByVal CompareFlags As StringCompareFlags_NT5, _
    ByVal String1Pointer As Long, _
    ByVal String1Length As Long, _
    ByVal String2Pointer As Long, _
    ByVal String2Length As Long _
) As bluW32.StringCompareResult

Public Enum StringCompareFlags_NT5
    w32IgnoreCase = &H1&
    w32IgnoreNonSpace = &H2&
    w32IgnoreSymbols = &H4&
    w32IgnoreWidth = &H20000
    w32IgnoreKana = &H10000
    w32LinguisticCasing = &H8000000
End Enum

'Unicode & Locale-aware string comparison (Windows Vista+) _
 <msdn.microsoft.com/en-us/library/windows/desktop/dd317761(v=vs.85).aspx>
Public Declare Function StringCompare_NT6 Lib "kernel32" Alias "CompareStringEx" ( _
    ByVal LocaleNamePointer As Long, _
    ByVal CompareFlags As StringCompareFlags_NT6, _
    ByVal String1Pointer As Long, _
    ByVal String1Length As Long, _
    ByVal String2Pointer As Long, _
    ByVal String2Length As Long, _
    Optional ByVal VersionInfoPointer As Long = 0, _
    Optional ByVal ReservedPointer As Long = 0, _
    Optional ByVal Param As Long = 0 _
) As bluW32.StringCompareResult

Public Enum StringCompareFlags_NT6
    w32IgnoreCaseLinguistic = &H10&
    w32IgnoreDiacritics = &H20&
End Enum

'TODO: Clean this up

'Unicode character properties:
'--------------------------------------------------------------------------------------
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

