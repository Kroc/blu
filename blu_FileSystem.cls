VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "blu_FileSystem"
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
'CLASS :: bluFileSystem

'Dependencies       None, self-contained - INCOMPLETE
'Last Updated       12-JUL-15
'Last Update        Added a variety of functions

'Provides Unicode-aware file system interaction.

'File paths are passed for each function so that if the file/folder disappears _
 between calls, the state won't bork
 
'With thanks to Tanner Helland for his PhotoDemon pdFSO class which this module is _
 based around, though my own work.

'/// API //////////////////////////////////////////////////////////////////////////////

'In VB6 True is -1 and False is 0, but in the Win32 API it's 1 for True
Private Enum BOOL
    API_TRUE = 1
    API_FALSE = 0
End Enum

'Some of the more modern WIN32 APIs return 0 for success instead of 1, it varies _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa378137(v=vs.85).aspx>
Private Enum HRESULT
    S_OK = 0
    S_FALSE = 1
End Enum

'Path handling:
'--------------------------------------------------------------------------------------
'The internal length limit on path strings
Private Const MAX_PATH As Long = 260

'Fix a path, unfolding any relative elements (e.g. "..", ".") _
 <msdn.microsoft.com/en-us/library/bb773569(v=vs.85).aspx>
Private Declare Function api_PathCanonicalize Lib "shlwapi" Alias "PathCanonicalizeW" ( _
    ByVal ResultStringPointer As Long, _
    ByVal SourceStringPointer As Long _
) As BOOL

'Combine two paths, fixing relative elements (i.e. "..", ".") between them _
 <msdn.microsoft.com/en-us/library/bb773571(v=vs.85).aspx>
Private Declare Function api_PathCombine Lib "shlwapi" Alias "PathCombineW" ( _
    ByVal ResultStringPointer As Long, _
    ByVal Path1StringPointer As Long, _
    ByVal Path2StringPointer As Long _
) As Long

'<msdn.microsoft.com/en-us/library/bb773660(v=vs.85).aspx>
Private Declare Function api_PathIsRelative Lib "shlwapi" Alias "PathIsRelativeW" ( _
    ByVal SourceStringPointer As Long _
) As BOOL

'<msdn.microsoft.com/en-us/library/bb773746(v=vs.85).aspx>
Private Declare Sub api_PathRemoveExtension Lib "shlwapi" Alias "PathRemoveExtensionW" ( _
    ByVal StringPointer As Long _
)

'<msdn.microsoft.com/en-us/library/bb773748(v=vs.85).aspx>
Private Declare Function api_PathRemoveFileSpec Lib "shlwapi" Alias "PathRemoveFileSpecW" ( _
    ByVal StringPointer As Long _
) As BOOL

'Special Folders:
'--------------------------------------------------------------------------------------

'Get the location of a special folder, e.g. "My Documents", "System32" &c. _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb762181(v=vs.85).aspx>
Private Declare Function api_SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathW" ( _
    ByVal hWndOwner As Long, _
    ByVal Folder As CSIDL, _
    ByVal Token As Long, _
    ByVal Flags As SHGFP, _
    ByVal ResultPointer As Long _
) As HRESULT

'Full list with descriptions here: _
 <msdn.microsoft.com/en-us/library/windows/desktop/bb762494(v=vs.85).aspx>
Public Enum CSIDL
    CSIDL_COMMON_VIDEO = &H37&      'All Users\My Video
    CSIDL_APPDATA = &H1A&           'Application data (roaming), intended for app data
    CSIDL_COMMON_APPDATA = &H23&    'Application data shared between all users
                                     'that should persist with the user between machines
    CSIDL_DESKTOP = &H0
    
    CSIDL_LOCAL_APPDATA = &H1C&     'Application data specific to the PC (e.g. cache)
    
    CSIDL_SYSTEM = &H25&            'Windows System, i.e. "System32"
    
    CSIDL_FLAG_CREATE = &H8000&     'OR this with any of the above to create the folder
                                     'if it doesn't exist (e.g. user deleted My Pictures)
End Enum

Private Enum SHGFP
    SHGFP_TYPE_CURRENT = 0          'Get the folder's current path (it may have moved)
    SHGFP_TYPE_DEFAULT = 1          'Get the default path
End Enum

'File attributes:
'--------------------------------------------------------------------------------------

'Get the attributes from a file (also a quick way of testing a file exists) _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa364944(v=vs.85).aspx>
Private Declare Function api_GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" ( _
    ByVal FileNamePointer As Long _
) As FILE_ATTRIBUTE

'<msdn.microsoft.com/en-us/library/windows/desktop/gg258117(v=vs.85).aspx>
Private Enum FILE_ATTRIBUTE
    INVALID_FILE_ATTRIBUTES = &HFFFFFFFF
    
    FILE_ATTRIBUTE_NORMAL = &H80&
    
    FILE_ATTRIBUTE_READONLY = &H1&
    FILE_ATTRIBUTE_HIDDEN = &H2&
    FILE_ATTRIBUTE_SYSTEM = &H4&
    FILE_ATTRIBUTE_DIRECTORY = &H10&
    FILE_ATTRIBUTE_ARCHIVE = &H20&
    
    FILE_ATTRIBUTE_COMPRESSED = &H800&
    FILE_ATTRIBUTE_ENCRYPTED = &H4000&
    
    FILE_ATTRIBUTE_TEMPORARY = &H100&
    FILE_ATTRIBUTE_SPARSE_FILE = &H200&
    FILE_ATTRIBUTE_REPARSE_POINT = &H400&
    FILE_ATTRIBUTE_OFFLINE = &H1000&
    
    FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000&
    FILE_ATTRIBUTE_DEVICE = &H40&       'Reserved
    FILE_ATTRIBUTE_VIRTUAL = &H10000    'Reserved
End Enum

Private Const ERROR_SHARING_VIOLATION As Long = 32

'File Loading & Saving:
'--------------------------------------------------------------------------------------

'Open a file for reading/writing _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa363858(v=vs.85).aspx>
Private Declare Function api_CreateFile Lib "kernel32" Alias "CreateFileW" ( _
    ByVal FileNamePointer As Long, _
    ByVal DesiredAccess As GENERIC, _
    ByVal ShareMode As FILE_SHARE, _
    ByVal SecurityAttributesPointer As Long, _
    ByVal CreationDisposition As CREATE, _
    ByVal FlagsAndAttributes As FILE_FLAGS, _
    ByVal TemplateFileHandle As Long _
) As Long

'Simple file access flags _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa446632(v=vs.85).aspx>
Private Enum GENERIC
    GENERIC_ALL = &H10000000
    GENERIC_READ = &H80000000
    GENERIC_WRITE = &H40000000
    GENERIC_EXECUTE = &H20000000
End Enum

Private Enum FILE_SHARE
    FILE_SHARE_READ = 1                 'Allow others to read the file
    FILE_SHARE_WRITE = 2                'Allow others to write to the file
    FILE_SHARE_DELETE = 4               'Allow others to delete/rename the file
End Enum

Private Enum CREATE
    CREATE_NEW = 1                      'Create a file only if it doesn't exist
    CREATE_ALWAYS = 2                   'Create or overwrite
    OPEN_EXISTING = 3                   'Open a file only if it already exists
    TRUNCATE_EXISTING = 5               'If file exists, reduce it to zero bytes
End Enum

Private Enum FILE_FLAGS
    'The file does not have other attributes set
    'This attribute is valid only if used alone
    FILE_ATTRIBUTE_NORMAL = &H80&
    
    FILE_ATTRIBUTE_ARCHIVE = &H20&      'Mark the file as archived
    FILE_ATTRIBUTE_HIDDEN = &H2&        'Hide the file
    FILE_ATTRIBUTE_READONLY = &H1&      'Make the file read-only
    FILE_ATTRIBUTE_SYSTEM = &H4&        'Make the file a System file
    FILE_ATTRIBUTE_TEMPORARY = &H100&   'The file is intended for temporary usage
    
    'Hint to caching that the file will be accessed randomly
    FILE_FLAG_RANDOM_ACCESS = &H10000000
    'Hint to caching that the file be read/written from start to finish
    FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
End Enum

'Get the size of a file in bytes, including ones over 4 GB _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa364957(v=vs.85).aspx>
Private Declare Function api_GetFileSizeEx Lib "kernel32" Alias "GetFileSizeEx" ( _
    ByVal FileHandle As Long, _
    ByRef FileSize As Currency _
) As BOOL

'Maximum file size we can load = (2 ^ 31) - 1
Private Const FILE_MAX As Long = 2147483647

'Read the contents of a file _
 <msdn.microsoft.com/en-us/library/windows/desktop/aa365467(v=vs.85).aspx>
Private Declare Function api_ReadFile Lib "kernel32" Alias "ReadFile" ( _
    ByVal FileHandle As Long, _
    ByVal BufferPointer As Long, _
    ByVal NumberOfBytesToRead As Long, _
    ByRef NumberOfBytesRead As Long, _
    ByVal OverlappedPointer As Long _
) As BOOL

'Close an opened file (or other) handle _
 <msdn.microsoft.com/en-us/library/windows/desktop/ms724211(v=vs.85).aspx>
Private Declare Function api_CloseHandle Lib "kernel32" Alias "CloseHandle" ( _
    ByVal Handle As Long _
) As BOOL

'/// CLASS ////////////////////////////////////////////////////////////////////////////

'A lot of methods here rely upon a Null-terminated string buffer of `MAX_PATH` length; _
 a bluString makes this easy for us to work with
Private APIString As bluString

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    Set APIString = New bluString
    Let APIString.Buffer = MAX_PATH
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    Set APIString = Nothing
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'FileExists
'======================================================================================
'FilePath       | File path to check for
'---------------¦----------------------------------------------------------------------
'Returns        | True if the given file exists, False otherwise
'======================================================================================
Public Function FileExists( _
    ByRef FilePath As String _
) As Boolean
    'The Unicode API will return a sharing violation for system files, _
     which tells us that the file does exist, but that we're not allowed to touch it
    'This method of testing is based upon this answer: _
     <vbforums.com/showthread.php?784047&viewfull=1#post4810609>
    If (api_GetFileAttributes(StrPtr(FilePath)) And vbDirectory) = 0 _
        Then Let FileExists = True _
        Else Let FileExists = (Err.LastDllError = ERROR_SHARING_VIOLATION)
End Function

'FileFolder : Get the portion of a path that corresponds to the folder a file is in
'======================================================================================
'FilePath       | A path containing a file name
'---------------¦----------------------------------------------------------------------
'Returns        | A folder path, sans end-slash
'======================================================================================
Public Function FileFolder( _
    ByRef FilePath As String _
) As String
    Let APIString.Text = FilePath
    'Now call the Windows API to trim the filename
    If api_PathRemoveFileSpec(APIString.Pointer) = API_TRUE Then
        'If something was removed, then the new length needs to be set
        Call APIString.TrimToNull
    End If
    Let FileFolder = APIString.Text
End Function

'FileReadBinary_AsArray : Read a binary file into a byte-array
'======================================================================================
'FilePath       | String containing the path to the file to read
'ReturnArray()  | Uinitialised byte array to accept the file contents
'---------------¦----------------------------------------------------------------------
'Returns        | An error number
'======================================================================================
Public Function FileReadBinary_AsArray( _
    ByRef FilePath As String, _
    ByRef ReturnArray() As Byte _
) As Long
    'Use the Windows API to access the file, _
     this avoids VB's slow and unwieldy error handling
    Dim FileHandle As Long
    'NOTE: For reasons not yet understood, this call is actually *slower* reading from _
           USB/SSD drives than using VB's `InputB`... go figure
    Let FileHandle = api_CreateFile( _
                    FileNamePointer:=StrPtr(FilePath), _
                      DesiredAccess:=GENERIC_READ, _
                          ShareMode:=FILE_SHARE_READ, _
          SecurityAttributesPointer:=0&, _
                CreationDisposition:=OPEN_EXISTING, _
                 FlagsAndAttributes:=FILE_FLAGS.FILE_ATTRIBUTE_NORMAL _
                                  Or FILE_FLAGS.FILE_FLAG_SEQUENTIAL_SCAN, _
                 TemplateFileHandle:=0& _
    )
    
    If FileHandle = -1 Then
        Stop
        Exit Function
    End If
    
    'Get the file size using the Windows API
    Dim FileSize As Currency
    If api_GetFileSizeEx(FileHandle, FileSize) = API_FALSE Then
        Call api_CloseHandle(FileHandle)
        Stop
        Exit Function
    End If
    
    'The Currency type has two decimal places, so push this up to whole bytes
    Let FileSize = FileSize * 10000
    
    'Is the file too big?
    'Note that we cannot open a file larger than 2 GB as we will be walking _
     the buffer using a signed Long which will go negative above 2 Billion
    If FileSize > FILE_MAX Then
        Call api_CloseHandle(FileHandle)
        Stop
        Exit Function
    End If
    
    'For speed, use a Long instead of a Currency
    Dim FileLength As Long
    Let FileLength = FileSize
    
    ReDim ReturnArray(0 To FileLength) As Byte
    
    Call api_ReadFile( _
           FileHandle:=FileHandle, _
        BufferPointer:=VarPtr(ReturnArray(0)), _
  NumberOfBytesToRead:=FileLength, _
    NumberOfBytesRead:=FileLength, _
    OverlappedPointer:=0& _
    )
    
    Call api_CloseHandle(FileHandle)
End Function

'FileReadText_AsArray : Returns an Integer array of UTF-16 Unicode points for a file
'======================================================================================
'FilePath       | Path to the file to read
'---------------¦----------------------------------------------------------------------
'Returns        | An error number
'======================================================================================
Public Function FileReadText_AsArray( _
    ByRef FilePath As String, _
    ByRef ReturnArray() As Integer _
) As Long
    'Get the text file as raw binary, as the file encoding could be unknown. _
     we'll work out the encoding (UTF-8/16, ANSI/ASCII) and convert the file _
     into a standard Windows UTF-16 (UCS-2) string
    Dim FileBuffer() As Byte
    Let FileReadText_AsArray = FileReadBinary_AsArray( _
        FilePath, FileBuffer _
    )
    
    If FileReadText_AsArray <> 0 Then
        Stop
        Exit Function
    End If
    
    Dim FileLength As Long
    Let FileLength = UBound(FileBuffer)

'    'Check for a Byte-Order-Mark:
'    '----------------------------------------------------------------------------------
'    'Not many files have a Byte-Order-Mark that specifies the file encoding, _
'     but if it is there it makes life a lot easier
'    Dim Encoding As blu_FileEncoding
'
'    'Read the first four bytes:
'    'TODO: Check the file is even four bytes long!
'    Dim BOM(1 To 4) As Byte
'    Get #FileNumber, , BOM()
'
'    ' FF FE         UTF-16, little endian
'    ' FE FF         UTF-16, big endian
'    ' EF BB BF      UTF-8
'    ' FF FE 00 00   UTF-32, little endian
'    ' 00 00 FE FF   UTF-32, big-endian
'
'    'Check for either UTF-16 Little Endian or UTF32 Little Endian
'    If BOM(1) = &HFF Then
'        'In both cases the second byte must be $FE
'        If BOM(2) = &HFE Then
'            'If UTF-32 there will be two nulls indicating a 4-byte character
'            If (BOM(3) = 0) And (BOM(4) = 0) _
'                Then Let Encoding = UTF32_LE _
'                Else Let Encoding = UTF16_LE
'        End If
'    'UTF-16 Big Endian begins with the $FE byte
'    ElseIf BOM(1) = &HFE Then
'        'Verify that the second byte is $FF
'        If BOM(2) = &HFF Then Let Encoding = UTF16_BE
'    'Look for the UTF-8 three-byte sequence
'    ElseIf BOM(1) = &HEF Then
'        'Note that UTF-8 is neither Little or Big Endian
'        If (BOM(2) = &HBB) And (BOM(3) = &HBF) Then Let Encoding = UTF8
'    'Nulls first could be a Big Endian UTF-32 file
'    ElseIf BOM(1) = 0 Then
'        'Check the next three bytes for the unique pattern
'        If (BOM(2) = 0) And (BOM(3) = &HFE) And (BOM(4) = &HFF) _
'            Then Let Encoding = UTF32_BE
'    End If

    'Walk, validate and convert:
    '-----------------------------------------------------------------------------------
    Dim i As Long
    'When parsing ASCII/UTF-8, bytes above 127 will be treated as ANSI _
     (Windows 1252 code-page) if invalid as a UTF-8 sequence
    Dim Windows1252ToUTF16(0 To &HFF) As Integer
    'Up to $80, ANSI is the same as ASCII
    For i = 0 To &H7F&: Let Windows1252ToUTF16(i) = i: Next i
    'These mappings were produced with the help of _
     <lingua-systems.com/unicode-converter/unicode-mappings/encode-windows-1252-to-utf-8-unicode.html>
    'NOTE: For undefined points, the Unicode Replacement Character is suplimented _
     so as not to leave a Null in our string which may break API calls using it
    Let Windows1252ToUTF16(&H80&) = &H20AC      ' €   EURO SIGN
    Let Windows1252ToUTF16(&H81&) = &HFFFD      ' ?   REPLACEMENT CHARACTER
    Let Windows1252ToUTF16(&H82&) = &H201A      ' ‚   SINGLE LOW-9 QUOTATION MARK
    Let Windows1252ToUTF16(&H83&) = &H192       ' ƒ   LATIN SMALL LETTER F WITH HOOK
    Let Windows1252ToUTF16(&H84&) = &H201E      ' „   DOUBLE LOW-9 QUOTATION MARK
    Let Windows1252ToUTF16(&H85&) = &H2026      ' …   HORIZONTAL ELLIPSIS
    Let Windows1252ToUTF16(&H86&) = &H2020      ' †   DAGGER
    Let Windows1252ToUTF16(&H87&) = &H2021      ' ‡   DOUBLE DAGGER
    Let Windows1252ToUTF16(&H88&) = &H2C6       ' ˆ   MODIFIER LETTER CIRCUMFLEX ACCENT
    Let Windows1252ToUTF16(&H89&) = &H2030      ' ‰   PER MILLE SIGN
    Let Windows1252ToUTF16(&H8A&) = &H160       ' Š   LATIN CAPITAL LETTER S WITH CARON
    Let Windows1252ToUTF16(&H8B&) = &H2039      ' ‹   SINGLE LEFT-POINTING ANGLE QUOTATION MARK
    Let Windows1252ToUTF16(&H8C&) = &H152       ' Œ   LATIN CAPITAL LIGATURE OE
    Let Windows1252ToUTF16(&H8D&) = &HFFFD      ' ?   REPLACEMENT CHARACTER
    Let Windows1252ToUTF16(&H8E&) = &H17D       ' Ž   LATIN CAPITAL LETTER Z WITH CARON
    Let Windows1252ToUTF16(&H8F&) = &HFFFD      ' ?   REPLACEMENT CHARACTER
    Let Windows1252ToUTF16(&H90&) = &HFFFD      ' ?   REPLACEMENT CHARACTER
    Let Windows1252ToUTF16(&H91&) = &H2018      ' ‘   LEFT SINGLE QUOTATION MARK
    Let Windows1252ToUTF16(&H92&) = &H2019      ' ’   RIGHT SINGLE QUOTATION MARK
    Let Windows1252ToUTF16(&H93&) = &H201C      ' “   LEFT DOUBLE QUOTATION MARK
    Let Windows1252ToUTF16(&H94&) = &H201D      ' ”   RIGHT DOUBLE QUOTATION MARK
    Let Windows1252ToUTF16(&H95&) = &H2022      ' •   BULLET
    Let Windows1252ToUTF16(&H96&) = &H2013      ' –   EN DASH
    Let Windows1252ToUTF16(&H97&) = &H2014      ' —   EM DASH
    Let Windows1252ToUTF16(&H98&) = &H20DC      ' ˜    SMALL TILDE
    Let Windows1252ToUTF16(&H99&) = &H2122      ' ™   TRADE MARK SIGN
    Let Windows1252ToUTF16(&H9A&) = &H161       ' š   LATIN SMALL LETTER S WITH CARON
    Let Windows1252ToUTF16(&H9B&) = &H203A      ' ›   SINGLE RIGHT-POINTING ANGLE QUOTATION MARK
    Let Windows1252ToUTF16(&H9C&) = &H153       ' œ   LATIN SMALL LIGATURE OE
    Let Windows1252ToUTF16(&H9D&) = &HFFFD      ' ?   REPLACEMENT CHARACTER
    Let Windows1252ToUTF16(&H9E&) = &H17E       ' ž   LATIN SMALL LETTER Z WITH CARON
    Let Windows1252ToUTF16(&H9F&) = &H178       ' Ÿ   LATIN CAPITAL LETTER Y WITH DIAERESIS
    'Only points $80 to $9F differ from standard Unicode points
    For i = &HA0& To &HFF&: Let Windows1252ToUTF16(i) = i: Next i

'    Dim CodePoint(1 To 4) As Byte
'    Dim CodePoint8 As Byte
'    Dim CodePoint16 As Integer
'    Dim CodePoint32 As Long
'
'    'For UTF-32 we only need to split the surrogate pairs
'    '..................................................................................
'    If Encoding = UTF32_LE Then
'        Do While VBA.EOF(FileNumber) = True
'            'Read the full four bytes
'            Get #FileNumber, , CodePoint32
'            'If the upper 2 bytes are 0, then the code point is 16-bits only
'            If (CodePoint32 \ &H10000) = 0 Then
'                'Add as-is to our string
'                Call Win32.MoveMemory( _
'                    OutputPointer:=Data(i), _
'                    SourcePointer:=VarPtr(CodePoint32), _
'                      LengthBytes:=2 _
'                )
'                Let i = i + 1
'
'            'If a full 32-bit value, then it needs to be converted to a pair of UTF-16 _
'             characters using the low/high surrogate characters (D8xx & DCxx)
'            Else
'
'                Stop
'            End If
'        Loop
'
'    End If

UTF8:
    '----------------------------------------------------------------------------------
    '<vovisoft.com/unicode/unifunctions.htm#ToUTF16>
    '<rsdn.ru/forum/vb/2316535.1>
    
    'It's not possible for a UTF-8 file to have more characters than it has bytes, _
     so we can set the number of characters in our string to match (at least) the _
     number of bytes in the file. After the file is parsed, the buffer will be cut _
     down to the final number of characters
    
    ReDim ReturnArray(0 To FileLength) As Integer
    
    Dim Byte1 As Byte, Byte2 As Byte, Byte3 As Byte
    Dim B As Long
    
    Let i = 0
    Do
        'Read one byte to begin with
        Let Byte1 = FileBuffer(B): Let B = B + 1
        
        If Byte1 = 0 Then
            '
            
        'If this is <128 then it's the same in ASCII/ANSI/UTF-8
        '..............................................................................
        ElseIf (Byte1 And &H80) = 0 Then
            'Add it to our string, and continue
            Let ReturnArray(i) = Byte1
            Let i = i + 1
            
        'UTF-8 byte sequences will begin with either "110?????" ($C0-$DF), _
         "1110????" ($E0-$EF) or "11110???" ($F0-$F7). Therefore bytes between $80 _
         and $9F will be treated as ANSI and converted from the Windows-1252 code-page
        '..............................................................................
        'This will test that the top three bits are "100?????" ($80-$9F)
        ElseIf (Byte1 And &HE0&) = &H80& Then
            Let ReturnArray(i) = Windows1252ToUTF16(Byte1)
            Let i = i + 1
            
        'Is this a 2-byte UTF-8 sequence? _
         (check that the top three bits are "110?????")
        '..............................................................................
        ElseIf (Byte1 And &HE0&) = &HC0& Then
            'Fetch another byte to see if this is a UTF-8 sequence
            Let Byte2 = FileBuffer(B): Let B = B + 1
            
            'The bytes in a UTF-8 sequence must be in the form "10??????"
            If (Byte2 And &HC0&) = &H80& Then
                'Decode the two UTF-8 bytes into a Unicode point
                Let ReturnArray(i) = (Byte1 And &H1F) * &H40& + (Byte2 And &H3F&)
                Let i = i + 1
            Else
                'The second byte is not part of a UTF-8 sequence, _
                 the first byte is above 127 so has to be treated as ANSI
                Let ReturnArray(i) = Windows1252ToUTF16(Byte1)
                Let i = i + 1
                'The second byte could be either ASCII or ANSI, but our conversion _
                 table keeps the ASCII bytes the same anyway
                Let ReturnArray(i) = Windows1252ToUTF16(Byte2)
                Let i = i + 1
            End If
            
        'Is this a 3-byte UTF-8 sequence? _
         (check that the top four bits are "1110????")
        '..............................................................................
        ElseIf (Byte1 And &HF0&) = &HE0& Then
            'Fetch another byte to see if this is a UTF-8 sequence
            Let Byte2 = FileBuffer(B): Let B = B + 1
            
            'The bytes in a UTF-8 sequence must be in the form "10??????"
            If (Byte2 And &HC0&) = &H80& Then
                'Get the third byte
                Let Byte3 = FileBuffer(B): Let B = B + 1
                
                'Check that this too follows the correct form "10??????"
                If (Byte2 And &HC0&) = &H80& Then
                    'Decode the three UTF-8 bytes into a Unicode point
                    Let ReturnArray(i) = (Byte1 And &HF&) * &H1000 _
                                       + (Byte2 And &H3F&) * &H40 _
                                       + (Byte3 And &H3F&)
                    Let i = i + 1
                Else
                    'Not a valid UTF-8 sequence, treat all three bytes as ASCII/ANSI
                    Let ReturnArray(i) = Windows1252ToUTF16(Byte1): Let i = i + 1
                    Let ReturnArray(i) = Windows1252ToUTF16(Byte2): Let i = i + 1
                    Let ReturnArray(i) = Windows1252ToUTF16(Byte3): Let i = i + 1
                End If
            Else
                'The second byte is not part of a UTF-8 sequence, _
                 the first byte is above 127 so has to be treated as ANSI
                Let ReturnArray(i) = Windows1252ToUTF16(Byte1)
                Let i = i + 1
                'The second byte could be either ASCII or ANSI, but our conversion _
                 table keeps the ASCII bytes the same anyway
                Let ReturnArray(i) = Windows1252ToUTF16(Byte2)
                Let i = i + 1
            End If
            
        'Not a plausible UTF-8 sequence byte, treat as ANSI
        Else
            Let ReturnArray(i) = Windows1252ToUTF16(Byte1)
            Let i = i + 1
        End If
    Loop While B <= FileLength
    
    ReDim Preserve ReturnArray(0 To i - 1) As Integer
    
    'Clean up the small conversion table
    Erase Windows1252ToUTF16
    
Finish:
End Function

'PathCombine
'======================================================================================
'LeftPath       | The left-hand part of the path, e.g. "C:"
'RightPath      | The right-hand part of the path, e.g. "path\file.ext"
'---------------¦----------------------------------------------------------------------
'Returns        | A new, combined and fixed, path; e.g. "C:\path\file.ext"
'======================================================================================
Public Function PathCombine( _
    ByRef LeftPath As String, _
    ByRef RightPath As String _
) As String
    'The Windows API requires an adequate buffer for the result
    Let PathCombine = Space$(MAX_PATH)
    
    'Call the Windows API. This will return 0 if it failed
    If api_PathCombine( _
        ResultStringPointer:=APIString.Pointer, _
         Path1StringPointer:=StrPtr(LeftPath), Path2StringPointer:=StrPtr(RightPath) _
    ) = 0 Then
        Let PathCombine = vbNullString
    Else
        'We need to trim the result to he null to know the actual length
        Let PathCombine = APIString.TrimToNull().Text
    End If
End Function

'PathFix : Fix a path, accounting for relative elements (e.g. "..", ".")
'======================================================================================
'Path           | Path to fix
'---------------¦----------------------------------------------------------------------
'Returns        | The path, now fixed. Any error returns null-string
'======================================================================================
Public Function PathFix( _
    ByRef Path As String _
) As String
    'The Windows API requires an adequate buffer for the result
    Let PathFix = Space$(MAX_PATH)
    
    'Call the Windows API. This will write a new null-terminated string
    If api_PathCanonicalize(StrPtr(PathFix), StrPtr(Path)) = API_TRUE Then
        'We need to trim the result to he null to know the actual length
        Let PathFix = Left$(PathFix, InStr(PathFix, vbNullChar) - 1)
    Else
        'An error with the API, return null-string
        Let PathFix = vbNullString
    End If
End Function

'PathIsAbsolute
'======================================================================================
'Path           |
'---------------¦----------------------------------------------------------------------
'Returns        |
'======================================================================================
Public Function PathIsAbsolute( _
    ByRef Path As String _
) As Boolean
    Let PathIsAbsolute = (api_PathIsRelative(StrPtr(Path)) = API_FALSE)
End Function

Public Function PathChangeExtension( _
    ByRef Path As String, _
    ByRef NewExtension As String _
) As String
    'Let's not modify the string passed to us!
    Let PathChangeExtension = Path
    
    If Path = vbNullString Then Exit Function
    
    'We can use the Windows API to strip the file-extension for us
    Call api_PathRemoveExtension(StrPtr(PathChangeExtension))
    Let PathChangeExtension = Left$(PathChangeExtension, InStr(PathChangeExtension, vbNullChar) - 1)
    
    Let PathChangeExtension = PathChangeExtension & NewExtension
End Function

'SpecialFolder : Get the path to a system folder, e.g. AppData
'======================================================================================
'Returns        |
'======================================================================================
Public Function SpecialFolder( _
    ByVal Folder As CSIDL _
) As String
    'Return null should this fail
    Let SpecialFolder = vbNullString
    
    'Fill a buffer to receive the path
    Dim Result As String
    Let Result = Space$(MAX_PATH)
    'Attempt to get the special folder path, creating it if it doesn't exist _
     (e.g. the user deleted the "My Pictures" folder)
    If api_SHGetFolderPath( _
        0&, Folder Or CSIDL_FLAG_CREATE, 0&, SHGFP_TYPE_CURRENT, StrPtr(Result) _
    ) = S_OK Then
        'The string will be null-terminated; find the end and trim
        Let SpecialFolder = Left$(Result, InStr(1, Result, vbNullChar) - 1)
    End If
End Function
