VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bluArrayStrings"
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
'CLASS :: bluArrayString

'Dependencies       bluString.bas
'Last Updated       08-DEC-14
'Last Update        Changed to a 1-based array to make certain aspects simpler and _
                    support an Index property to return the Index given a string, _
                    or 0 for 'not found'. This is more practical for the caller than _
                    checking for existance and then finding the Index number

'--------------------------------------------------------------------------------------

'This is a class to make handling an array of Strings much easier to manage without _
 polluting your own code with redimming and bounds checking

'/// CLASS ////////////////////////////////////////////////////////////////////////////

Private My_Strings() As String          'Our array of VB Strings
Private My_Count As Long                'Number of Strings in the array (1-based)

'String comparison will be case-insensitive by default
Private My_CaseSensitive As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    'This will unload each of the bluStrings in our array
    Erase My_Strings
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Add : Add a VB String to our array
'======================================================================================
'Text           | A VB String to add to our array
'---------------+----------------------------------------------------------------------
'Returns        | Index of the new string in our array
'======================================================================================
Public Function Add( _
    ByRef Text As String _
) As Long
    'Check if the String already exists
    Let Add = Index(Text)
    'If not, increase the array and get the next index to use
    If Add = 0 Then Let Add = Increase()
    'And add the text
    Let My_Strings(Add) = Text
End Function

'AddVariant : Convert Variants to Strings, including recursive array walking
'======================================================================================
'AddVar         | Some value of any type, including arrays
'======================================================================================
Public Sub AddVariant( _
    ByRef AddVar As Variant _
)
    'We don't want to calculate array bounds on every loop
    Dim First As Long, Last As Long
    
    'Check the type of the parameter (Long, String, Double &c.)
    Dim AddVarType As VBA.VbVarType
    Let AddVarType = VarType(AddVar)
       
    Dim i As Long
    Select Case AddVarType
        'These types can be converted to a String
        Case vbString, vbBoolean, vbByte, vbCurrency, vbDate, vbDecimal, _
             vbDouble, vbEmpty, vbError, vbInteger, vbLong, vbSingle
            '..........................................................................
            Call Add(CStr(AddVar))
            
        Case vbObject
            '..........................................................................
            'Use the contents of a bluString
            If TypeOf AddVar Is bluString Then
                Call Add(AddVar.Text)
                
            'Walk another bluArrayString
            ElseIf TypeOf AddVar Is bluArrayStrings Then
                Let First = 1: Let Last = AddVar.Count
                For i = First To Last: Call Add(AddVar.Item(i)): Next i
            Else
                Err.Raise (13)
            End If
            
        'Note that vbArray is combined with the data type inside the array, _
         including Variant too, i.e. `vbArray + vbVariant`
        Case Is >= vbArray
            '..........................................................................
            'Arrays need to be walked, pushing each value as its own parameter
            Let First = LBound(AddVar): Let Last = UBound(AddVar)
            For i = First To Last: Call AddVariant(AddVar(i)): Next i
            
        Case Else
            '..........................................................................
            'All other data types can't be co-erced
            Err.Raise (13)
    End Select
End Sub

'CaseSensitive : Turn case-sensitivity on or off
'======================================================================================
Public Property Get CaseSensitive() As Boolean
    Let CaseSensitive = My_CaseSensitive
End Property
Public Property Let CaseSensitive(ByRef Mode As Boolean)
    Let My_CaseSensitive = Mode
End Property

'Clear : Empty the array
'======================================================================================
Public Sub Clear()
    Erase My_Strings
End Sub

'Count : Number of Strings in our array
'======================================================================================
Public Property Get Count() As Long: Let Count = My_Count: End Property

'Index : Get the Index number for a string in the array
'======================================================================================
'SearchFor      | A String to search for in the array.
'               | Uses the `CaseSensitive` property to determine search behaviour
'---------------+----------------------------------------------------------------------
'Returns        | The Index number of the String in the array, otherwise 0
'======================================================================================
Public Property Get Index( _
    ByRef SearchFor As String _
) As Long
    'If the array is empty return 0 immediately
    If My_Count = 0 Then Exit Property
    
    'We'll use a bluString for easy, quick case-[in]sensitive comparisons
    Dim Comparison As bluString
    Set Comparison = New bluString
    Let Comparison.CaseSensitive = My_CaseSensitive
    Let Comparison.Text = SearchFor
    
    Dim i As Long
    For i = 1 To My_Count
        'Do the comparison with our array, and return if matched
        If Comparison.Equals(My_Strings(i)) Then Let Index = i: Exit Property
    Next i
End Property

'Item : Get access to a String in the array
'======================================================================================
'Index          | An Index number for the array contents
'---------------+----------------------------------------------------------------------
'Returns        | The VB String in the array at the given Index
'======================================================================================
Public Property Get Item( _
    ByVal Index As Long _
) As String
Attribute Item.VB_MemberFlags = "200"
    'If there are no items in the array, we're out of bounds whatever the Index is
    If My_Count = 0 Then Err.Raise (9): Exit Property
    'Is the index within bounds?
    If (Index < 1) Or (Index > My_Count) Then Err.Raise (9): Exit Property
    'Return the VB String from our array
    Let Item = My_Strings(Index)
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
    Let Length = Len(My_Strings(Index))
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
    Let Pointer = StrPtr(My_Strings(Index))
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
    ReDim Preserve My_Strings(1 To Increase) As String
End Function
