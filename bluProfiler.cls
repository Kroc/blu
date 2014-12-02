VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bluProfiler"
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
'CLASS :: bluProfiler

'Dependencies       None (self-contained)
'Last Updated       30-NOV-14
'Last Update        Initial version

'--------------------------------------------------------------------------------------

'A tool for measuring the speed of various parts of your program
'How to use this class:

'1. Define an Enum with each event you desire included, giving you fast constant _
    IDs for each event

'2. Register your event names early as possible, ideally in `Sub Main`. E.g. _
    `Call Profiler.RegisterEvent(PAINT_SCREEN, "Paint Screen")`

'3. When you want to record how long a piece of code takes, begin the event: _
    `Call Profiler.EventBegin(PAINT_SCREEN)`
    
'4. After the code being measured, end the event: _
    `Call Profiler.EventEnd(PAINT_SCREEN)`

'--------------------------------------------------------------------------------------

'Inspired by "PhotoDemon Performance Profiler" ["pdProfiler.cls"], _
 part of PhotoDemon <photodemon.org>, copyright (C)2013-2014 by Tanner Helland _
 (though this class is an entirely different approach)

'/// API //////////////////////////////////////////////////////////////////////////////

Private Enum MMRESULT
    MMSYSERR_NOERROR = 0
    MMSYSERR_ERROR = 1
    TIMERR_NOERROR = 0
    TIMERR_NOCANDO = 97
    TIMERR_STRUCT = 129
End Enum

'Structure for retrieving timer resolution _
 <msdn.microsoft.com/library/windows/desktop/dd757625(v=vs.85).aspx>
Private Type TIMECAPS
    PeriodMin As Long
    PeriodMax As Long
End Type

'Gets the minimum and maximum resolution of the system timer _
 <msdn.microsoft.com/library/windows/desktop/dd757627(v=vs.85).aspx>
Private Declare Function api_timeGetDevCaps Lib "winmm" Alias "timeGetDevCaps" ( _
    ByRef Result As TIMECAPS, _
    ByVal SizeOf As Long _
) As MMRESULT

'Set the timer resolution _
 <msdn.microsoft.com/library/dd757624(v=vs.85).aspx>
Private Declare Function api_timeBeginPeriod Lib "winmm" Alias "timeBeginPeriod" ( _
    ByVal Period As Long _
) As MMRESULT

'Undo the change in timer resolution _
 <msdn.microsoft.com/library/dd757626(v=vs.85).aspx>
Private Declare Function api_timeEndPeriod Lib "winmm" Alias "timeEndPeriod" ( _
    ByVal Period As Long _
) As MMRESULT

'Get the current time in milliseconds _
 <msdn.microsoft.com/library/dd757629(v=vs.85).aspx>
Private Declare Function api_timeGetTime Lib "winmm" Alias "timeGetTime" ( _
) As Long

'/// CLASS DATA ///////////////////////////////////////////////////////////////////////

'The lowest timer resolution available
Private My_Resolution As Long

'This is the data recorded for each event
Private Type bluProfilerEvent
    Title As String                     'Friendly name, i.e. "Paint Screen"
    Begin As Long                       'Starting time of current event timing
    Elapsed As Long                     'Total running time
    Count As Long                       'Number of iterations
    Min As Long                         'Shortest event duration
    Max As Long                         'Longest event duration
End Type

'Where we store that data
Private Events() As bluProfilerEvent
'Number of registered events (0-based)
Private EventCount As Long

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Initialize
'======================================================================================
Private Sub Class_Initialize()
    'The number of events will begin on -1 so that we can easily initialise the array
    Let EventCount = -1
    
    'How fast can you go?
    Dim Result As TIMECAPS
    'TODO: handle error
    Call api_timeGetDevCaps(Result, Len(Result))
    Let My_Resolution = Result.PeriodMin
    
    'Speed up the system timer
    Call api_timeBeginPeriod(My_Resolution)
    
    'For the benefit of the IDE
    Debug.Print "[bluProfiler] Instantiated; Resolution: " & CStr(My_Resolution) & "ms"
End Sub

'CLASS Terminate
'======================================================================================
Private Sub Class_Terminate()
    'Restore the system timer to the previous speed
    Call api_timeEndPeriod(My_Resolution)
    
    'Dump a report
    Debug.Print "[bluProfiler] Terminated"
    Call Me.DebugPrint
    
    'Clear the data out
    Erase Events
End Sub

'/// PUBLIC INTERFACE /////////////////////////////////////////////////////////////////

Public Sub DebugPrint()
    Debug.Print "[bluProfiler] Event Report:                   Iterations Elapsed    Min    Average    Max"
    Debug.Print "========================================================================================="
    
    Dim i As Long
    For i = 0 To EventCount: With Events(i)
        Debug.Print "* " _
            & Left$(.Title & Space$(46), 46) _
            & " " & Right$(Space$(7) & Format$(CStr(.Count), "#,0"), 7) _
            & " " & Right$(Space$(7) & Format$(CStr(.Elapsed), "#,0"), 7) _
            & " " & Right$(Space$(6) & Format$(CStr(.Min), "#,0"), 6) _
            & " " & Right$(Space$(10) & Format$(CStr(.Elapsed / .Count), "#,0.00"), 10) _
            & " " & Right$(Space$(6) & Format$(CStr(.Max), "#,0"), 6)
    End With: Next i
End Sub

'EventBegin
'======================================================================================
Public Sub EventBegin( _
    ByRef EventID As Long _
)
    Let Events(EventID).Begin = api_timeGetTime()
End Sub

'EventEnd
'======================================================================================
Public Sub EventEnd( _
    ByRef EventID As Long _
)
    Dim Time As Long
    
    With Events(EventID)
        Let Time = api_timeGetTime() - .Begin
        Let .Elapsed = .Elapsed + Time
        If Time < .Min Then Let .Min = Time
        If Time > .Max Then Let .Max = Time
        Let .Count = .Count + 1
    End With
End Sub

'RegisterEvent
'======================================================================================
'EventID        | An index number to identify the event
'EventTitle     | A friendly name for the event
'======================================================================================
Public Sub RegisterEvent( _
    ByRef EventID As Long, _
    ByRef EventTitle As String _
)
    'Is there enough events in the array?
    If EventID > EventCount Then
        Let EventCount = EventID
        ReDim Preserve Events(-1 To EventCount) As bluProfilerEvent
    End If
    
    'Assign the friendly name and reset the data
    With Events(EventID)
        Let .Title = EventTitle
        Let .Begin = 0
        Let .Elapsed = 0
        Let .Min = 2 ^ 30
        Let .Max = 0
    End With
    
    Debug.Print "[bluProfiler] Registered: '" & EventTitle & "'"
End Sub

'Resolution : The minimum time (in milliseconds) between time requests
'======================================================================================
Public Property Get Resolution() As Long: Let Resolution = My_Resolution: End Property
