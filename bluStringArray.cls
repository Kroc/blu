VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bluStringArray"
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
'CLASS :: bluStringArray

'Dependencies       bluString.bas
'Last Updated       25-NOV-14
'Last Update        Initial implementation

'A stack-like, indexed array of bluStrings. Whilst you can use a standard VB Array, _
 this class will make adding, growing and walking the array much easier without so _
 much glue on your end

'/// CLASS ////////////////////////////////////////////////////////////////////////////

Private Strings() As bluString          'Our array of bluStrings
Private StringsInit As Long             'If it's dimmed yet or not
Private My_Count As Long                'Number of strings in the array

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    '
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    'This will unload each of the bluStrings in our array
    Erase Strings
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

'Add : Add a VB String to our array
'======================================================================================
'Text           | A VB String to add to the array
'Returns        | Index of the new string in our array
'======================================================================================
Public Function Add( _
    ByRef Text As String _
) As Long
    'Increase the array and get the next index to use
    Let Add = Increase()
    
    'Instantiate the new bluString
    Set Strings(Add) = New bluString
    'Now load the given text
    Let Strings(Add).Text = Text
End Function

Public Function AddBluString( _
    ByRef Text As bluString _
) As Long
    'Increase the array and get the next index to use
    Let AddBluString = Increase()
    
    'Set the object pointer to that of the passed bluString. _
     Note that this will keep your bluString alive outside of scope!
    Set Strings(AddBluString) = Text
End Function

'Count
'======================================================================================
Public Property Get Count() As Long: Let Count = My_Count + 1: End Property

'Item : Get access to a bluString instance in the array
'======================================================================================
'Index          | An Index number for the array contents
'Returns        | The bluString in the array at the given Index
'======================================================================================
Public Property Get Item( _
    ByRef Index As Long _
) As bluString
Attribute Item.VB_MemberFlags = "200"
    'If there are no items in the array, we're out of bounds whatever the Index is
    If StringsInit = 0 Then Err.Raise (9): Exit Property
    'Is the index within bounds?
    If (Index < 0) Or (Index > My_Count) Then Err.Raise (9): Exit Property
    'Return the bluString from our array
    Set Item = Strings(Index)
End Property

'/// PRIVATE INTERFACE ////////////////////////////////////////////////////////////////

'Increase : Increases the size of the array and returns the next index to be used
'======================================================================================
Private Function Increase() As Long
    'If there are no strings in the array, we will need to initialise it
    If StringsInit = 0 Then
        'Initialise the array
        ReDim Strings(0) As bluString
        'This item will be filled
        Let Increase = 0
        Let StringsInit = 1
    Else
        'The next index will be filled
        Let My_Count = My_Count + 1
        Let Increase = My_Count
        'Increase the array bounds
        ReDim Preserve Strings(Increase) As bluString
    End If
End Function
