VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "blu_Sys"
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
'CLASS :: blu_Sys

'Dependencies       None (self-contained)
'Last Updated       12-JUL-15
'Last Update        Initial version - INCOMPLETE

'/// API //////////////////////////////////////////////////////////////////////////////

'Windows Version:
'--------------------------------------------------------------------------------------

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

'/// CLASS ////////////////////////////////////////////////////////////////////////////

Public InIDE As Boolean

'Once the Windows version has been detected, we want to be able to refer to it _
 as fast as possible. It's far too slow to fetch the Windows Version every time we _
 need it, but it's also a little slow to test the version number thousands of times. _
 We therefore cache the result of the version comparison in this fashion:
Public Enum bluSys_WindowsAPILevel
    XP_OR_LOWER = -1
    VISTA_OR_HIGHER = 0
End Enum
Public WindowsAPILevel As bluSys_WindowsAPILevel

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    
    Let InIDE = Not CBool(App.LogMode = 1)
    
    '----------------------------------------------------------------------------------
    
    'Get the Windows version as Windows Vista and above use alternative APIs for _
     locale-dependent methods (such as `Equals`)
    Dim VersionInfo As OSVERSIONINFO
    Let VersionInfo.SizeOfMe = Len(VersionInfo)
    'NOTE: If the app is in compatibility mode, this will return the compatible _
     Windows version, not the actual version. Windows 8.1 and above also cause _
     this API to always return as Windows 8.0 unless the app opts-in to correct _
     version info via manifest, but this isn't an issue to bluString as we are _
     only concerned with testing for Windows Vista+
    If api_GetVersionEx(VersionInfo) = BOOL_TRUE Then
        If VersionInfo.MajorVersion > 5 _
            Then Let WindowsAPILevel = VISTA_OR_HIGHER _
            Else Let WindowsAPILevel = XP_OR_LOWER
    End If
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////



