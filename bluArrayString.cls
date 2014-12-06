VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bluArrayString"
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
'Last Updated       06-DEC-14
'Last Update        Initial implementation

'This is a class to make handling an array of Strings much easier to manage without _
 polluting your own code with redimming and bounds checking

'/// CLASS ////////////////////////////////////////////////////////////////////////////

Private My_Strings() As String          'Our array of VB Strings
Private Initialised As Long             'If it's dimmed yet or not (Boolean)
Private My_Count As Long                'Number of Strings in the array

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
    'Increase the array and get the next index to use
    Let Add = Increase()
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
            ElseIf TypeOf AddVar Is bluArrayString Then
                For i = 0 To AddVar.Count - 1
                    Call Add(AddVar.Item(i))
                Next i
            Else
                Err.Raise (13)
            End If
            
        'Note that vbArray is combined with the data type inside the array, _
         including Variant too, i.e. `vbArray + vbVariant`
        Case Is >= vbArray
            '..........................................................................
            'Arrays need to be walked, pushing each value as its own parameter
            For i = LBound(AddVar) To UBound(AddVar)
                Call AddVariant(AddVar(i))
            Next i
            
        Case Else
            '..........................................................................
            'All other data types can't be co-erced
            Err.Raise (13)
    End Select
End Sub

'Clear : Empty the array
'======================================================================================
Public Sub Clear()
    Erase My_Strings
    Let Initialised = 0
End Sub

'Count : Number of Strings in our array
'======================================================================================
Public Property Get Count() As Long: Let Count = My_Count + 1: End Property

'Item : Get access to a String in the array
'======================================================================================
'Index          | An Index number for the array contents
'---------------+----------------------------------------------------------------------
'Returns        | The VB String in the array at the given Index
'======================================================================================
Public Property Get Item( _
    ByRef Index As Long _
) As String
Attribute Item.VB_MemberFlags = "200"
    'If there are no items in the array, we're out of bounds whatever the Index is
    If Initialised = 0 Then Err.Raise (9): Exit Property
    'Is the index within bounds?
    If (Index < 0) Or (Index > My_Count) Then Err.Raise (9): Exit Property
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
    ByRef Index As Long _
) As Long
    'If there are no items in the array, we're out of bounds whatever the Index is
    If Initialised = 0 Then Err.Raise (9): Exit Property
    'Is the index within bounds?
    If (Index < 0) Or (Index > My_Count) Then Err.Raise (9): Exit Property
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
    ByRef Index As Long _
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
    ByRef Index As Long _
) As Long
    'If there are no items in the array, we're out of bounds whatever the Index is
    If Initialised = 0 Then Err.Raise (9): Exit Property
    'Is the index within bounds?
    If (Index < 0) Or (Index > My_Count) Then Err.Raise (9): Exit Property
    'Return the String pointer -- be careful with this!
    Let Pointer = StrPtr(My_Strings(Index))
End Property

'/// PRIVATE INTERFACE ////////////////////////////////////////////////////////////////

'Increase : Increases the size of the array
'======================================================================================
'Returns        | The new index number
'======================================================================================
Private Function Increase() As Long
    'If there are no strings in the array, we will need to initialise it
    If Initialised = 0 Then
        'Initialise the array
        ReDim My_Strings(0) As String
        'This item will be filled
        Let Increase = 0
        Let Initialised = 1
    Else
        'The next index will be filled
        Let My_Count = My_Count + 1
        Let Increase = My_Count
        'Increase the array bounds
        ReDim Preserve My_Strings(Increase) As String
    End If
End Function
