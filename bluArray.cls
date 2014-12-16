VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bluArray"
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
'INTERFACE :: bluArray

'Dependencies       None
'Last Updated       16-DEC-14
'Last Update        Added `AllowDuplicates`, `Clear`

'--------------------------------------------------------------------------------------

'This is a generic interface to allow you to handle different types of bluArrays _
 without knowing their specific type. If your class uses this interface, then you _
 won't need to include every type of bluArray class if you don't need them

'AddVariant : Add a Variant to the array
'======================================================================================
Public Sub AddVariant( _
    ByRef AddVar As Variant _
): End Sub

'AllowDuplicates : Allow adding of duplicate values
'======================================================================================
Public Property Get AllowDuplicates() As Boolean: End Property
Public Property Let AllowDuplicates(ByVal Mode As Boolean): End Property

'AsString : Return an item in the array as a String equivilent
'======================================================================================
Public Property Get AsString( _
    ByVal Index As Long _
) As String: End Sub

'Clear : Empty the array
'======================================================================================
Public Sub Clear(): End Sub

'Count : The number of items in the array (1-based), 0 = none
'======================================================================================
Public Property Get Count( _
) As Long: End Property
