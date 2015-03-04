VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bluMagic"
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
'CLASS :: bluMagic

'Everything this class does is pure magic

'Status             Planning to refactor
'Dependencies       None (self-contained)
'Last Updated       24-JUL-13

'NOTE: I intend to tidy this code, format it in my style and generally make it a _
 little more understandable

'The following documentation is slightly re-formatted from the original and not _
 written by me (Kroc Camen)

'--------------------------------------------------------------------------------------

'LaVolpe: _
"Paul Caton has done some amazing things combining ASM and VB. Unfortunately, he has _
 moved on from VB and left this code to me. His original code can be found at the _
 links provided below. I am taking the initiative to learn from his thunks, update, _
 rewrite, or otherwise improve those thunks. Therefore, some of this code will differ _
 significantly from Paul's original work. As an ASM newbie, I welcome any _
 suggestions/corrections for the ASM code itself.

'If you are new to these thunking routines, here is the basic idea:

'1. Make subclassing, hooking, callbacks IDE-safe. No crashes due to END or STOP
'2. Be able to allow subclassing, hooking, callbacks to use a class, form, _
    property page or usercontrol to call back to (just like an AddressOf pointer)
'3. Allow that 'AddressOf' pointer to exist in any VB code object

'Original work done by Paul Caton. _
 Contributions made by Tai Chi Minh Ralph Eastwood and myself

'Caton: <planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=64867&lngWId=1> _
 Caton: <planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=51403&lngWId=1> _
 Eastwood: <planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=66665&lngWId=1> _
 This Post: <planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=68737>

'I have now completely revamped the subclassing ASM. It is larger than Paul's original _
 work but I believe it is now 100% END safe, IDE-safe. I have even placed END _
 statements inside of the callback procedures and IDE did not crash. The additional _
 ASM code added to Paul's original work, included use of a self-killing delay timer _
 that keeps the ASM alive in virtual memory and doubles as a trigger to delete the _
 virtual memory.

'The issue with Paul's code regarding usercontrols that subclass ContainerHwnd _
 multiple times has been resolved with the use of the timers.

' Other subclassing thunk modifications/enhancements are listed below:
'   - thunks can also be used inside of property pages if so desired
'   - subclassing thunks will automatically release after window processes WM_NCDestroy message (last message a window gets)
'   - subclassing thunks can destroy a parent-less API window that is subclassed, when IDE ends (optional parameter)
'   - as mentioned above, use of a delay timer appears to make subclasser 100% safe, crash-proof
'   - callback thunks can be individually set & released
'   - callback thunks have optional parameter to auto-kill timers when IDE ends
'   - hook thunks can filter on before and/or after messages just like the subclassing thunks

'   Above all being said, the ASM thunks are but a small part of what Paul Caton put together.
'   Without his zAddressOf and zProbe functions, easy subclassing and containment would not
'   be where it is right now. We all owe him a bow of respect for that.

' Unicode is supported in subclassing and hooking. Not applicable with callbacks because pointers are only used.
'   - Since VB windows are not unicode windows, unicode support is not provided by default
'       However, VB can create unicode windows via CreateWindowExW, therefore...
'   - To request unicode support, supply the appropriate parameter to the Subclass and/or shk_SetHook routines
'       -- check the bUnicode parameter's return value to validate success
'       -- if successful, ensure you use unicode API calls, unicode strings as needed
'           --- Exceptions. If either the following tests (depending on situation) returns false, unicode will not be used
'           1) Subclassing: The window being subclassed must return non-zero from IsWindowUnicode(hWnd)
'           2) Hooking: Non-zero must be returned from IsWindowUnicode(GetDesktopWindow)

'/// API //////////////////////////////////////////////////////////////////////////////

'Shared: _
 --------------------------------------------------------------------------------------
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Enum eThunkType
    SubclassThunk = 0
    HookThunk = 1
    CallbackThunk = 2
End Enum

'Subclassing: _
 --------------------------------------------------------------------------------------
Public Enum eMsgWhen                            'When to callback
    MSG_BEFORE = 1                              'Callback before the original WndProc
    MSG_AFTER = 2                               'Callback after the original WndProc
    MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER
End Enum

' see Subclass for complete listing of indexes and what they relate to
Private Const IDX_PARM_USER As Long = 13    'Thunk data index of the User-defined callback parameter data index
Private Const IDX_UNICODE   As Long = 107   'Must be UBound(subclass thunkdata)+1; index for unicode support
Private Const MSG_ENTRIES   As Long = 32    'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all subclassed windows
Private Const ALL_MESSAGES  As Long = -1    'All messages will callback

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Hooking: _
 --------------------------------------------------------------------------------------
Private Declare Function SetWindowsHookExA Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function SetWindowsHookExW Lib "user32" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long

'<msdn2.microsoft.com/en-us/library/ms644990.aspx>
Public Enum eHookType
    WH_MSGFILTER = -1
    WH_JOURNALRECORD = 0
    WH_JOURNALPLAYBACK = 1
    WH_KEYBOARD = 2
    WH_GETMESSAGE = 3
    WH_CALLWNDPROC = 4
    WH_CBT = 5
    WH_SYSMSGFILTER = 6
    WH_MOUSE = 7
    WH_DEBUG = 9
    WH_SHELL = 10
    WH_FOREGROUNDIDLE = 11
    WH_CALLWNDPROCRET = 12
    WH_KEYBOARD_LL = 13       ' NT/2000/XP+ only, Global hook only
    WH_MOUSE_LL = 14          ' NT/2000/XP+ only, Global hook only
End Enum

'========================================================================================================
' COMMON-USE ITEMS USED IN ALL THUNKS
'========================================
'   Routines: must include these regardless if using subclassing, hooking, callbacks
'         zError - returns error information as needed; entire routine can be rem'd out if you so desire
'         zFnAddr - Returns the AddressOf pointer for API DLL functions
'         zMap_VFunction - returns memory handle to the ASM for a specific thunk
'         zAddressOf - returns the AddressOf function pointer for a form/uc/class/ppg routine.
'         zProbe - a worker function for zAddressOf
'         zData (Get/Let) - used to set/get subclass related data
'         zInIDE - used to determine if running IDE or not
'         zUnThunk - used to release a specific subclass, hook or callback
'         zTerminateThunks - used to release all subclasses, hooks or callbacks
'     all above are local use only routines. Do not make public
'     jump to zMap_Function. All common routines start there, copy & paste as needed
  
' Local variables/constants: must declare these regardless if using subclassing, hooking, callbacks
Private asmFunk             As Collection   'hWnd/thunk-address collection; initialized as needed
Private z_hkFunk            As Collection   'hook/thunk-address collection; initialized as needed
Private z_cbFunk            As Collection   'callback/thunk-address collection; initialized as needed
Private Const IDX_INDEX     As Long = 2     'index of the subclassed hWnd OR hook type
Private Const IDX_PREVPROC  As Long = 9     'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11    'Thunk data index of the Before table for messages
Private Const IDX_ATABLE    As Long = 12    'Thunk data index of the After table for messages
Private Const IDX_CALLBACKORDINAL As Long = 36 ' Ubound(callback thunkdata)+1, index of the callback

'========================================================================================================
' TO USE IDE-SAFE SUBCLASSING...
'===============================
'   Special Notes:
'     a. Subclass uses function ordinals. Last routine in the class/form/uc/ppg is Ordinal #1, second to last is Ordinal #2, etc
'           IMPORTANT: Routines used for subclass procedure MUST NOT be Public, can be Private or Friend
'     b. VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
'     c. Do not terminate application while stopped in the subclass procedure
'     d. At a minimum, review Subclass, UnSubclass & Terminate
'
'   1. Include all common-use items above
'   2. Include the following routines:  ssc_ = self-subclassing
'       Subclass - used to initiate a subclassing session
'    -  Terminate - used to terminate all subclassing sessions
'    -  UnSubclass - used to terminate a single subclassing session
'       AddMsg - adds a message to subclass
'       DelMsg - removes a message from subclassing
'    -  ssc_CallOrigWndProc - enables the CallWindowProc API function on subclassed windows
'    -  zGet_lParamUser, zSet_lParamUser - returns/replaces the optional lParamUser set in Subclass & shk_SetHook
'    *  zAddMsg - this does the work of AddMsg. See above
'    *  zDelMsg - this does the work of DelMsg. See above
'  (*) are local use only, do not make public. All others can be made public if desired
'  (-) can be removed if not needed. See their routine's comments
'   3. In your form/uc/class/ppg unload/terminate event, include the statement: Terminate or zTerminateThunks SubclassThunk
'   4. Include the generic window procedure routine: myWndProc at bottom of this module
'   5. Also include the following declarations

'========================================================================================================
' TO USE IDE-SAFE HOOKING...
'==============================
'   Special Notes:
'     a. shk_SetHook uses function ordinals. Last routine in the class/form/uc is Ordinal #1, second to last is Ordinal #2, etc
'           IMPORTANT: Routines used for hooks MUST NOT be Public, can be Private or Friend
'     b. VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
'     c. When END executed, app does not crash, but hook may not be released immediately. Should you notice this,
'        simply do something that would trigger the hook (i.e., move  mouse, keystroke, depending on hook type).
'        The hook then will release.
'     d. Do not terminate application while stopped in the hook procedure
'     e. At a minimum, review shk_SetHook, shk_Unhook & shk_TerminateHooks
'
'   1. Include all common-use items from above "common-use" section
'   2. Include the following routines:  shk_ = self-Hooking
'       shk_SetHook - used to initiate a hooking session
'    -  shk_TerminateHooks - used to terminate all hooks
'    -  shk_Unhook - used to terminate a single hooking session
'    -  shk_CallNextHook - enables the CallNextHookEx API function on hooked threads
'    -  zGet_lParamUser, zSet_lParamUser - returns/replaces the optional lParamUser set in Subclass & shk_SetHook
'   (-) can be removed if not needed. See their comments
'   all of above can be made public if so desired.
'   3. In your form/uc/class/ppg unload/terminate event, include the statement: shk_TerminateHooks or zTerminateThunks HookThunk
'   4. Include the generic hook procedure routine: myHookProc near bottom of this module
'   5. Include the following if they do not exist in your code already (also used in subclass section above):
'        Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Long
'        Private Const IDX_PARM_USER As Long = 13       'Thunk data index of the User-defined callback parameter data index
'        Private Enum eMsgWhen                          'When to callback
'          MSG_BEFORE = 1                               'Callback before the original WndProc/HookProc
'          MSG_AFTER = 2                                'Callback after the original WndProc/HookProc
'          MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER   'Callback before and after the original WndProc/HookProc
'        End Enum
'   6. Also include the following declarations which do not exist in the common-use or subclassing sections

'========================================================================================================
' TO USE IDE-SAFE CALLBACKS...
'==============================
'   Special notes:
'    a. scb_SetCallbackAddr uses function ordinals. Last routine in the class/form/uc/ppg is Ordinal #1, second to last is Ordinal #2, etc
'           IMPORTANT: Routines used for callbacks MUST NOT be Public, can be Private or Friend
'    b. The callback function in your class/form/uc/ppg must be a Function that returns Long.
'       For example. The callback routine for a timer has no return value per MSDN, but when that timer
'       procedure is coded, the callback function must return a long. The return value is not important in that case
'    c. VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
'    d. Do not terminate application while stopped in the callback procedure
'    e. Never call scb_ReleaseCallback from within the callback procedure
'    f. At a minimum, review scb_SetCallbackAddr, scb_ReleaseCallback & scb_TerminateCallbacks
'
'   1. Include all common-use items in above "common-use" section
'   2. Include the following routines:  scb_ = self-callback
'       scb_SetCallbackAddr - used the same as AddressOf for class/form/uc/ppg routines
'    -  scb_ReleaseCallback - used to release memory for a specific callback
'    -  scb_TerminateCallbacks - used to terminate all callback addresses (release allocated memory)
'  (-) can be removed if not needed. See their comments
'   all of above can be made public if so desired
'   3. In your form/uc/class unload/terminate event, include the statement: scb_TerminateCallbacks or zTerminateThunks CallbackThunk
'   4. No additional delcarations are required, all that are needed exist in the common-use section above

'/// EVENTS ///////////////////////////////////////////////////////////////////////////

'CLASS Terminate _
 ======================================================================================
Private Sub Class_Terminate()
    'Stop any subclassing
    Call zTerminateThunks(SubclassThunk)
    'Remove any hooks
    Call zTerminateThunks(HookThunk)
    'Detatch any callbacks
    Call zTerminateThunks(CallbackThunk)
End Sub

'/// SUBCLASSING //////////////////////////////////////////////////////////////////////
'Routines specific to subclassing

'Subclass : Subclass the specified window handle _
 ======================================================================================
Public Function Subclass( _
             ByVal lng_hWnd As Long, _
    Optional ByVal lParamUser As Long = 0, _
    Optional ByVal nOrdinal As Long = 1, _
    Optional ByVal oCallback As Object = Nothing, _
    Optional ByVal UseIDESafety As Boolean = True, _
    Optional ByRef bUnicode As Boolean = False, _
    Optional ByVal bIsAPIwindow As Boolean = False _
) As Boolean
    '*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* UseIDESafety - Optional, enable/disable IDE safety measures. There is not reason to set this to False
    '* bUnicode - Optional, if True, Unicode API calls should be made to the window vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know if ANSI to be used or not
    '* bIsAPIwindow - Optional, if True DestroyWindow will be called if IDE ENDs
    '*****************************************************************************************
    '** Subclass.asm - subclassing thunk
    '**
    '** Paul_Caton@hotmail.com
    '** Copyright free, use and abuse as you see fit.
    '**
    '** v2.0 Re-write by LaVolpe, based mostly on Paul Caton's original thunks....... 20070720
    '** .... Reorganized & provided following additional logic
    '** ....... Unsubclassing only occurs after thunk is no longer recursed
    '** ....... Flag used to bypass callbacks until unsubclassing can occur
    '** ....... Timer used as delay mechanism to free thunk memory afer unsubclassing occurs
    '** .............. Prevents crash when one window subclassed multiple times
    '** .............. More END safe, even if END occurs within the subclass procedure
    '** ....... Added ability to destroy API windows when IDE terminates
    '** ....... Added auto-unsubclass when WM_NCDESTROY received
    '*****************************************************************************************
    ' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

    'Thunk machine-code initialised here
    Dim asm(0 To IDX_UNICODE) As Long
    
    Const SUB_NAME      As String = "Subclass"          'This routine's name
    Const CODE_LEN      As Long = 4 * IDX_UNICODE + 4   'Thunk length in bytes
    Const PAGE_RWX      As Long = &H40&                 'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&               'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&               'Release allocated memory flag
    Const GWL_WNDPROC   As Long = -4                    'SetWindowsLong WndProc index
    Const WNDPROC_OFF   As Long = &H60                  'Thunk offset to the WndProc execution address
    Const MEM_LEN       As Long = _
        CODE_LEN + (8 * (MSG_ENTRIES + 1))              'Bytes to allocate per thunk, _
                                                         (data + code + msg tables)
    
    'This is the complete listing of thunk offset values and what they point/relate to.
    'Those rem'd out are used elsewhere or are initialized in Declarations section
  
  'Const IDX_RECURSION  As Long = 0     'Thunk data index of callback recursion count
  'Const IDX_SHUTDOWN   As Long = 1     'Thunk data index of the termination flag
  'Const IDX_INDEX      As Long = 2     'Thunk data index of the subclassed hWnd
   Const IDX_EBMODE     As Long = 3     'Thunk data index of the EbMode function address
   Const IDX_CWP        As Long = 4     'Thunk data index of the CallWindowProc function address
   Const IDX_SWL        As Long = 5     'Thunk data index of the SetWindowsLong function address
   Const IDX_FREE       As Long = 6     'Thunk data index of the VirtualFree function address
   Const IDX_BADPTR     As Long = 7     'Thunk data index of the IsBadCodePtr function address
   Const IDX_OWNER      As Long = 8     'Thunk data index of the Owner object's vTable address
  'Const IDX_PREVPROC   As Long = 9     'Thunk data index of the original WndProc
   Const IDX_CALLBACK   As Long = 10    'Thunk data index of the callback method address
  'Const IDX_BTABLE     As Long = 11    'Thunk data index of the Before table
  'Const IDX_ATABLE     As Long = 12    'Thunk data index of the After table
  'Const IDX_PARM_USER  As Long = 13    'Thunk data index of the User-defined callback parameter data index
   Const IDX_DW         As Long = 14    'Thunk data index of the DestroyWinodw function address
   Const IDX_ST         As Long = 15    'Thunk data index of the SetTimer function address
   Const IDX_KT         As Long = 16    'Thunk data index of the KillTimer function address
   Const IDX_EBX_TMR    As Long = 20    'Thunk code patch index of the thunk data for the delay timer
   Const IDX_EBX        As Long = 26    'Thunk code patch index of the thunk data
  'Const IDX_UNICODE    As Long = xx    'Must be UBound(subclass thunkdata)+1; index for unicode support
    
    Dim asmMem          As Long         'Thunk base address
    Dim nAddr           As Long
    Dim nID             As Long
    Dim nMyID           As Long
    Dim bIDE            As Boolean

    'Ensure the window handle is valid
    If IsWindow(lng_hWnd) = 0 Then
        Call zError(SUB_NAME, "Invalid window handle")
        Exit Function
    End If
    
    Let nMyID = GetCurrentProcessId                     'Get this process's ID
    Call GetWindowThreadProcessId(lng_hWnd, nID)        'Get the process ID associated with the window handle
    'Ensure that the window handle doesn't belong to another process
    If nID <> nMyID Then
        Call zError(SUB_NAME, "Window handle belongs to another process")
        Exit Function
    End If
    
    'If the user hasn't specified the callback owner
    If oCallback Is Nothing Then Set oCallback = Me
    
    nAddr = zAddressOf(oCallback, nOrdinal)             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                   'Ensure that we've found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If
        
    asmMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    
    If asmMem <> 0 Then                                  'Ensure the allocation succeeded
    
      If asmFunk Is Nothing Then Set asmFunk = New Collection 'If this is the first time through, do the one-time initialization
      On Error GoTo CatchDoubleSub                              'Catch double subclassing
        asmFunk.Add asmMem, "h" & lng_hWnd                    'Add the hWnd/thunk-address to the collection
      On Error GoTo 0
      
   'asm (0) thru asm(17) are used as storage for the thunks & IDX_ constants above relate to these thunk positions which are filled in below
    asm(18) = &HD231C031: asm(19) = &HBBE58960: asm(21) = &H21E8F631: asm(22) = &HE9000001: asm(23) = &H12C&: asm(24) = &HD231C031: asm(25) = &HBBE58960: asm(27) = &H3FFF631: asm(28) = &H75047339: asm(29) = &H2873FF23: asm(30) = &H751C53FF: asm(31) = &HC433913: asm(32) = &H53FF2274: asm(33) = &H13D0C: asm(34) = &H18740000: asm(35) = &H875C085: asm(36) = &H820443C7: asm(37) = &H90000000: asm(38) = &H87E8&: asm(39) = &H22E900: asm(40) = &H90900000: asm(41) = &H2C7B8B4A: asm(42) = &HE81C7589: asm(43) = &H90&: asm(44) = &H75147539: asm(45) = &H6AE80F: asm(46) = &HD2310000: asm(47) = &HE8307B8B: asm(48) = &H7C&: asm(49) = &H7D810BFF: asm(50) = &H8228&: asm(51) = &HC7097500: asm(52) = &H80000443: asm(53) = &H90900000: asm(54) = &H44753339: asm(55) = &H74047339: asm(56) = &H2473FF3F: asm(57) = &HFFFFFC68
    asm(58) = &H2475FFFF: asm(59) = &H811453FF: asm(60) = &H82047B: asm(61) = &HC750000: asm(62) = &H74387339: asm(63) = &H2475FF07: asm(64) = &H903853FF: asm(65) = &H81445B89: asm(66) = &H484443: asm(67) = &H73FF0000: asm(68) = &H646844: asm(69) = &H56560000: asm(70) = &H893C53FF: asm(71) = &H90904443: asm(72) = &H10C261: asm(73) = &H53E8&: asm(74) = &H3075FF00: asm(75) = &HFF2C75FF: asm(76) = &H75FF2875: asm(77) = &H2473FF24: asm(78) = &H891053FF: asm(79) = &H90C31C45: asm(80) = &H34E30F8B: asm(81) = &H1078C985: asm(82) = &H4C781: asm(83) = &H458B0000: asm(84) = &H75AFF228: asm(85) = &H90909023: asm(86) = &H8D144D8D: asm(87) = &H8D503443: asm(88) = &H75FF1C45: asm(89) = &H2C75FF30: asm(90) = &HFF2875FF: asm(91) = &H51502475: asm(92) = &H2073FF52: asm(93) = &H902853FF: asm(94) = &H909090C3: asm(95) = &H74447339: asm(96) = &H4473FFF7
    asm(97) = &H4053FF56: asm(98) = &HC3447389: asm(99) = &H89285D89: asm(100) = &H45C72C75: asm(101) = &H800030: asm(102) = &H20458B00: asm(103) = &H89145D89: asm(104) = &H81612445: asm(105) = &H4C4&: asm(106) = &H1862FF00

    ' cache callback related pointers & offsets
      asm(IDX_EBX) = asmMem                                                 'Patch the thunk data address
      asm(IDX_EBX_TMR) = asmMem                                             'Patch the thunk data address
      asm(IDX_INDEX) = lng_hWnd                                              'Store the window handle in the thunk data
      asm(IDX_BTABLE) = asmMem + CODE_LEN                                   'Store the address of the before table in the thunk data
      asm(IDX_ATABLE) = asmMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
      asm(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
      asm(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
      asm(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
      
      ' validate unicode request & cache unicode usage
      If bUnicode Then bUnicode = (IsWindowUnicode(lng_hWnd) <> 0&)
      asm(IDX_UNICODE) = bUnicode                                            'Store whether the window is using unicode calls or not
      
      ' get function pointers for the thunk
      If UseIDESafety = True Then                                               'If the user wants IDE protection
          Debug.Assert zInIDE(bIDE)
          If bIDE = True Then asm(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode) 'Store the EbMode function address in the thunk data
                                                        '^^ vb5 users, change vba6 to vba5
      End If
      If bIsAPIwindow Then                                                    'If user wants DestroyWindow sent should IDE end
          asm(IDX_DW) = zFnAddr("user32", "DestroyWindow", bUnicode)
      End If
      asm(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)           'Store the VirtualFree function address in the thunk data
      asm(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)        'Store the IsBadCodePtr function address in the thunk data
      asm(IDX_ST) = zFnAddr("user32", "SetTimer", bUnicode)                  'Store the SetTimer function address in the thunk data
      asm(IDX_KT) = zFnAddr("user32", "KillTimer", bUnicode)                 'Store the KillTimer function address in the thunk data
      
      If bUnicode Then
          asm(IDX_CWP) = zFnAddr("user32", "CallWindowProcW", bUnicode)      'Store CallWindowProc function address in the thunk data
          asm(IDX_SWL) = zFnAddr("user32", "SetWindowLongW", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory asmMem, VarPtr(asm(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          asm(IDX_PREVPROC) = SetWindowLongW(lng_hWnd, GWL_WNDPROC, asmMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      Else
          asm(IDX_CWP) = zFnAddr("user32", "CallWindowProcA", bUnicode)      'Store CallWindowProc function address in the thunk data
          asm(IDX_SWL) = zFnAddr("user32", "SetWindowLongA", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory asmMem, VarPtr(asm(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          asm(IDX_PREVPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, asmMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      End If
      If asm(IDX_PREVPROC) = 0 Then                                          'Ensure the new WndProc was set correctly
          zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
          GoTo ReleaseMemory
      End If
      'Store the original WndProc address in the thunk data
      RtlMoveMemory asmMem + IDX_PREVPROC * 4, VarPtr(asm(IDX_PREVPROC)), 4&
      Subclass = True                                                     'Indicate success
      
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
        
    End If

 Exit Function                                                                'Exit Subclass
    
CatchDoubleSub:
    zError SUB_NAME, "Window handle is already subclassed"
      
ReleaseMemory:
    'Subclass has failed after memory allocation, so release the memory
    Call VirtualFree(asmMem, 0, MEM_RELEASE)
End Function

'UnSubclass : UnSubclass the specified window handle _
 ======================================================================================
Public Sub UnSubclass(ByVal lng_hWnd As Long)
    ' can be made public, can be removed & zUnthunk can be called instead
    zUnThunk lng_hWnd, SubclassThunk
End Sub

'AddMsg : Add the message value to the window handle's specified callback table _
 ======================================================================================
Friend Sub AddMsg( _
    ByVal lng_hWnd As Long, ByVal When As eMsgWhen, _
    ParamArray Messages() As Variant _
)
    'Thunk base address
    Dim asmMem As Long
    
    asmMem = zMap_VFunction(lng_hWnd, SubclassThunk)           'Ensure that the thunk hasn't already released its memory
    If asmMem Then
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)
        Select Case VarType(Messages(M))                        ' ensure no strings, arrays, doubles, objects, etc are passed
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                         'If the message is to be added to the before original WndProc table...
              If zAddMsg(Messages(M), IDX_BTABLE, asmMem) = False Then 'Add the message to the before table
                When = (When And Not MSG_BEFORE)
              End If
            End If
            If When And MSG_AFTER Then                          'If message is to be added to the after original WndProc table...
              If zAddMsg(Messages(M), IDX_ATABLE, asmMem) = False Then 'Add the message to the after table
                When = (When And Not MSG_AFTER)
              End If
            End If
        End Select
      Next
    End If
End Sub

'DelMsg : Delete the message value from the window handle's specified callback table _
 ======================================================================================
Public Sub DelMsg( _
    ByVal lng_hWnd As Long, ByVal When As eMsgWhen, _
    ParamArray Messages() As Variant _
)
    Dim asmMem       As Long                                                   'Thunk base address
    
    asmMem = zMap_VFunction(lng_hWnd, SubclassThunk)                           'Ensure that the thunk hasn't already released its memory
    If asmMem Then
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)                              ' ensure no strings, arrays, doubles, objects, etc are passed
        Select Case VarType(Messages(M))
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                                         'If the message is to be removed from the before original WndProc table...
              zDelMsg Messages(M), IDX_BTABLE, asmMem                          'Remove the message to the before table
            End If
            If When And MSG_AFTER Then                                          'If message is to be removed from the after original WndProc table...
              zDelMsg Messages(M), IDX_ATABLE, asmMem                          'Remove the message to the after table
            End If
        End Select
      Next
    End If
End Sub

'Call the original WndProc
Public Function ssc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' can be made public, can be removed if you will not use this in your window procedure
    Dim asmMem       As Long                           'Thunk base address
    asmMem = zMap_VFunction(lng_hWnd, SubclassThunk)
    If asmMem Then                                     'Ensure that the thunk hasn't already released its memory
        If zData(IDX_UNICODE, asmMem) Then
            ssc_CallOrigWndProc = CallWindowProcW(zData(IDX_PREVPROC, asmMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        Else
            ssc_CallOrigWndProc = CallWindowProcA(zData(IDX_PREVPROC, asmMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        End If
    End If
End Function

'Get the subclasser lParamUser callback parameter
Private Function zGet_lParamUser(ByVal hWnd_Hook_ID As Long, ByVal vType As eThunkType) As Long
    ' can be removed if you never will retrieve or replace the user-defined parameter
    If vType <> CallbackThunk Then
        Dim asmMem       As Long                                       'Thunk base address
        asmMem = zMap_VFunction(hWnd_Hook_ID, vType)
        If asmMem Then                                                 'Ensure that the thunk hasn't already released its memory
          zGet_lParamUser = zData(IDX_PARM_USER, asmMem)               'Get the lParamUser callback parameter
        End If
    End If
End Function

'Let the subclasser lParamUser callback parameter
Private Sub zSet_lParamUser(ByVal hWnd_Hook_ID As Long, ByVal vType As eThunkType, ByVal newValue As Long)
    ' can be removed if you never will retrieve or replace the user-defined parameter
    If vType <> CallbackThunk Then
        Dim asmMem       As Long                                       'Thunk base address
        asmMem = zMap_VFunction(hWnd_Hook_ID, vType)
        If asmMem Then                                                 'Ensure that the thunk hasn't already released its memory
          zData(IDX_PARM_USER, asmMem) = newValue                      'Set the lParamUser callback parameter
        End If
    End If
End Sub

'Add the message to the specified table of the window handle
Private Function zAddMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal asmMem As Long) As Boolean
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long
      Dim i      As Long                                                        'Loop index
    
      zAddMsg = True
      nBase = zData(nTable, asmMem)                                            'Map zData() to the specified table
      
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
      Else
        
        nCount = zData(0, nBase)                                                'Get the current table entry count
        For i = 1 To nCount                                                     'Loop through the table entries
          If zData(i, nBase) = 0 Then                                           'If the element is free...
            zData(i, nBase) = uMsg                                              'Use this element
            GoTo Bail                                                           'Bail
          ElseIf zData(i, nBase) = uMsg Then                                    'If the message is already in the table...
            GoTo Bail                                                           'Bail
          End If
        Next                                                                  'Next message table entry
    
        nCount = i                                                             'On drop through: i = nCount + 1, the new table entry count
        If nCount > MSG_ENTRIES Then                                           'Check for message table overflow
          zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
          zAddMsg = False
          GoTo Bail
        End If
        
        zData(nCount, nBase) = uMsg                                            'Store the message in the appended table entry
      End If
    
      zData(0, nBase) = nCount                                                 'Store the new table entry count
Bail:
End Function

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal asmMem As Long)
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long
      Dim i      As Long                                                        'Loop index
    
      nBase = zData(nTable, asmMem)                                            'Map zData() to the specified table
    
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
        zData(0, nBase) = 0                                                     'Zero the table entry count
      Else
        nCount = zData(0, nBase)                                                'Get the table entry count
        
        For i = 1 To nCount                                                     'Loop through the table entries
          If zData(i, nBase) = uMsg Then                                        'If the message is found...
            zData(i, nBase) = 0                                                 'Null the msg value -- also frees the element for re-use
            GoTo Bail                                                           'Bail
          End If
        Next                                                                  'Next message table entry
        
       ' zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
      End If
Bail:
End Sub

'-SelfHook code------------------------------------------------------------------------------------
'-The following routines are exclusively for the shk_SetHook routines----------------------------
Friend Function shk_SetHook(ByVal HookType As eHookType, _
                    Optional ByVal bGlobal As Boolean, _
                    Optional ByVal When As eMsgWhen = MSG_BEFORE, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal UseIDESafety As Boolean = True, _
                    Optional ByRef bUnicode As Boolean = False) As Boolean 'Setting specified hook

    '*************************************************************************************************
    '* HookType - One of the eHookType enumerators
    '* bGlobal - If False, then hook applies to app's thread else it applies Globally (only supported by WH_KEYBOARD_LL & WH_MOUSE_LL)
    '* When - either MSG_AFTER, MSG_BEFORE or MSG_BEFORE_AFTER
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* UseIDESafety - Optional, enable/disable IDE safety measures. There is no reason to set this to false
    '* bUnicode - Optional, if True, Unicode API calls should be made to the window vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know if ANSI to be used or not
    '*************************************************************************************************
    ' Hook procedure must be declared identical to the one near the end of this class (Sample at Ordinal #2)

    ' \\LaVolpe - The ASM for this procedure rewritten to mirror Paul Caton's SelfSub ASM
    '       Therefore, it appears to be crash proof and allows a choice of whether you want
    '       hook messages before and/or after the VB gets the message
    
    Dim asm(0 To 66) As Long                   'Thunk machine-code initialised here
    Const MEM_LEN      As Long = 4 * 67         'Thunk length in bytes (last # must be = UBound zSc() + 1)
    
    Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&       'Release allocated memory flag
    Const IDX_EBMODE    As Long = 3             'Thunk data index of the EbMode function address
    Const IDX_CNH       As Long = 4             'Thunk data index of the CallNextHook function address
    Const IDX_UNW       As Long = 5             'Thunk data index of the UnhookWindowsEx function address
    Const IDX_OBJCHK    As Long = 6             'Thunk data index of the callback validation token
    Const IDX_BADPTR    As Long = 7             'Thunk data index of the IsBadCodePtr function address
    Const IDX_OWNER     As Long = 8             'Thunk data index of the Owner object's vTable address
    Const IDX_CALLBACK  As Long = 10            'Thunk data index of the callback method address
    Const IDX_BTABLE    As Long = 11            'Thunk data index of the Before flag
    Const IDX_ATABLE    As Long = 12            'Thunk data index of the After flag
    Const IDX_EBX       As Long = 16            'Thunk code patch index of the thunk data
    Const PROC_OFF      As Long = &H38          'Thunk offset to the HookProc execution address
    Const SUB_NAME      As String = "shk_SetHook" 'This routine's name
    
    Dim nAddr           As Long
    Dim nID             As Long
    Dim nMyID           As Long
    Dim bIDE            As Boolean
    Dim asmMem         As Long                 'Thunk base address

    If oCallback Is Nothing Then Set oCallback = Me 'If the user hasn't specified the callback owner
      
    nAddr = zAddressOf(oCallback, nOrdinal)         'Get the address of the specified ordinal method
    If nAddr = 0 Then                               'Ensure that we've found the ordinal method
      zError SUB_NAME, "Callback method not found"
      Exit Function
    End If
       
    If Not bGlobal Then nID = App.ThreadID                      ' thread ID to be used if not global hook
      
    asmMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)    'Allocate executable memory
    
    If asmMem <> 0 Then                                        'Ensure the allocation succeeded
      
        If z_hkFunk Is Nothing Then Set z_hkFunk = New Collection   'If this is the first time through, do the one-time initialization
        On Error GoTo CatchDoubleSub                                'Catch double subclassing
          z_hkFunk.Add asmMem, "h" & HookType                      'Add the hook/thunk-address to the collection
        On Error GoTo 0
        
        ' create the thunk; zSc(16) filled in below along with asm(2-13)
        asm(14) = &HD231C031: asm(15) = &HBBE58960: asm(17) = &H3FFF631: asm(18) = &H75047339: asm(19) = &H2873FF2F: asm(20) = &H751C53FF: asm(21) = &HC43391F: asm(22) = &H7B8B2E74: asm(23) = &H8478B20: asm(24) = &H75184339: asm(25) = &HC53FF0F: asm(26) = &H13D&: asm(27) = &H85197400: asm(28) = &H900975C0: asm(29) = &H443C7: asm(30) = &H90000080: asm(31) = &H47E8&: asm(32) = &H2AE900: asm(33) = &H90900000: asm(34) = &H742C7339: asm(35) = &H75894A0F: asm(36) = &H46E81C: asm(37) = &H75390000: asm(38) = &H90157514: asm(39) = &H27E8&: asm(40) = &H30733900: asm(41) = &HD2310A74: asm(42) = &H2FE8&: asm(43) = &H90909000: asm(44) = &H33390BFF: asm(45) = &H73390E75: asm(46) = &HFF097404: asm(47) = &H53FF2473: asm(48) = &H90909014: asm(49) = &HCC261: asm(50) = &HFF2C75FF: asm(51) = &H75FF2875: asm(52) = &H2473FF24
        asm(53) = &H891053FF: asm(54) = &H90C31C45: asm(55) = &H2873FF52: asm(56) = &H5A1C53FF: asm(57) = &H438D2275: asm(58) = &H144D8D34: asm(59) = &H1C458D50: asm(60) = &HFF0873FF: asm(61) = &H75FF2C75: asm(62) = &H2475FF28: asm(63) = &HFF525150: asm(64) = &H53FF2073: asm(65) = &H90909028: asm(66) = &HC3&

        asm(IDX_EBX) = asmMem                         'Patch the thunk data address
        asm(IDX_INDEX) = HookType                      'Store the hook type in the thunk data
        asm(IDX_OWNER) = ObjPtr(oCallback)             'Store the callback owner's object address in the thunk data
        asm(IDX_CALLBACK) = nAddr                      'Store the callback address in the thunk data
        asm(IDX_PARM_USER) = lParamUser                'Store the lParamUser callback parameter in the thunk data
        
        ' get a piece of the oCallback to use as a validation token
        RtlMoveMemory VarPtr(asm(IDX_OBJCHK)), asm(IDX_OWNER) + 8&, 4&

        ' validate unicode request & cache unicode usage
        If bUnicode Then bUnicode = (IsWindowUnicode(GetDesktopWindow) <> 0&)
        
        asm(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)    'Store the IsBadCodePtr function address in the thunk data
        asm(IDX_CNH) = zFnAddr("user32", "CallNextHookEx", bUnicode)       'Store CallWindowProc function address in the thunk data
        asm(IDX_UNW) = zFnAddr("user32", "UnhookWindowsHookEx", bUnicode)  'Store the SetWindowLong function address in the thunk data
        
        If UseIDESafety = True Then                                           'If the user wants IDE protection
            Debug.Assert zInIDE(bIDE)
            If bIDE = True Then asm(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode) 'Store the EbMode function address in the thunk data
        End If
        
        If (When And MSG_BEFORE) = MSG_BEFORE Then asm(IDX_BTABLE) = 1     ' non-zero flag if Before messages desired
        If (When And MSG_AFTER) = MSG_AFTER Then asm(IDX_ATABLE) = 1       ' non-zero flag if After messages desired
        
        RtlMoveMemory asmMem, VarPtr(asm(0)), MEM_LEN                     'Copy the thunk code/data to the allocated memory
        'Set the new HookProc, return the address of the original HookProc
        If bUnicode Then
            asm(IDX_PREVPROC) = SetWindowsHookExW(HookType, asmMem + PROC_OFF, App.hInstance, nID)
        Else
            asm(IDX_PREVPROC) = SetWindowsHookExA(HookType, asmMem + PROC_OFF, App.hInstance, nID)
        End If
        
        If asm(IDX_PREVPROC) = 0 Then                                              'Ensure the new HookProc was set correctly
          zError SUB_NAME, "SetWindowsHookEx failed, error #" & Err.LastDllError
          GoTo ReleaseMemory
        End If
        RtlMoveMemory asmMem + IDX_PREVPROC * 4, VarPtr(asm(IDX_PREVPROC)), 4&    'Store the callback address
            
        shk_SetHook = True                                                          'Indicate success
      Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
      End If
      
      Exit Function
    
CatchDoubleSub:
      zError SUB_NAME, "Hook is already established"
      
ReleaseMemory:
      VirtualFree asmMem, 0, MEM_RELEASE                                     'shk_SetHook has failed after memory allocation, so release the memory
    
End Function

'Call the next hook proc
Private Function shk_CallNextHook(ByVal HookType As eHookType, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' can be made public, can be removed if you will not use this in your hook procedure
    Dim asmMem       As Long                           'Thunk base address
    asmMem = zMap_VFunction(HookType, HookThunk)
    If asmMem Then                                     'Ensure that the thunk hasn't already released its memory
        shk_CallNextHook = CallNextHookEx(zData(IDX_PREVPROC, asmMem), nCode, wParam, lParam) 'Call the next hook proc
    End If
End Function

Friend Function shk_UnHook(ByVal HookType As eHookType) As Boolean
    ' can be made public, can be removed & zUnThunk can be called instead
    zUnThunk HookType, HookThunk
End Function


'-SelfCallback code------------------------------------------------------------------------------------
'-The following routines are exclusively for the scb_SetCallbackAddr routines----------------------------
Private Function scb_SetCallbackAddr(ByVal nParamCount As Long, _
                     Optional ByVal nOrdinal As Long = 1, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal UseIDESafety As Boolean = True, _
                     Optional ByVal bIsTimerCallback As Boolean) As Long   'Return the address of the specified callback thunk
    '*************************************************************************************************
    '* nParamCount  - The number of parameters that will callback
    '* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
    '* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* UseIDESafety   - Optional, set to false to disable IDE protection.
    '* bIsTimerCallback - optional, set to true for extra protection when used as a SetTimer callback
    '       If True, timer will be destroyed when IDE/app terminates. See scb_ReleaseCallback.
    '*************************************************************************************************
    ' Callback procedure must return a Long even if, per MSDN, the callback procedure is a Sub vs Function
    ' The number of parameters and their types are dependent on the individual callback procedures
    
    Const MEM_LEN     As Long = IDX_CALLBACKORDINAL * 4 + 4     'Memory bytes required for the callback thunk
    Const PAGE_RWX    As Long = &H40&                           'Allocate executable memory
    Const MEM_COMMIT  As Long = &H1000&                         'Commit allocated memory
    Const SUB_NAME      As String = "scb_SetCallbackAddr"       'This routine's name
    Const INDX_OWNER    As Long = 0                             'Thunk data index of the Owner object's vTable address
    Const INDX_CALLBACK As Long = 1                             'Thunk data index of the EbMode function address
    Const INDX_EBMODE   As Long = 2                             'Thunk data index of the IsBadCodePtr function address
    Const INDX_BADPTR   As Long = 3                             'Thunk data index of the IsBadCodePtr function address
    Const INDX_KT       As Long = 4                             'Thunk data index of the KillTimer function address
    Const INDX_EBX      As Long = 6                             'Thunk code patch index of the thunk data
    Const INDX_PARAMS   As Long = 18                            'Thunk code patch index of the number of parameters expected in callback
    Const INDX_PARAMLEN As Long = 24                            'Thunk code patch index of the bytes to be released after callback
    Const PROC_OFF      As Long = &H14                          'Thunk offset to the callback execution address

    Dim asmMem       As Long                                   'Thunk base address
    Dim z_Cb()    As Long                                       'Callback thunk array
    Dim nValue    As Long
    Dim nCallback As Long
    Dim bIDE      As Boolean
      
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    If z_cbFunk Is Nothing Then
        Set z_cbFunk = New Collection           'If this is the first time through, do the one-time initialization
    Else
        On Error Resume Next                    'Catch already initialized?
        asmMem = z_cbFunk.Item("h" & ObjPtr(oCallback) & "." & nOrdinal) 'Test it
        If Err = 0 Then
            scb_SetCallbackAddr = asmMem + PROC_OFF  'we had this one, just reference it
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    If nParamCount < 0 Then                     ' validate parameters
        zError SUB_NAME, "Invalid Parameter count"
        Exit Function
    End If
    
    nCallback = zAddressOf(oCallback, nOrdinal)         'Get the callback address of the specified ordinal
    If nCallback = 0 Then
        zError SUB_NAME, "Callback address not found."
        Exit Function
    End If
    asmMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
        
    If asmMem = 0& Then
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError  ' oops
        Exit Function
    End If
    z_cbFunk.Add asmMem, "h" & ObjPtr(oCallback) & "." & nOrdinal 'Add the callback/thunk-address to the collection
        
    ReDim z_Cb(0 To IDX_CALLBACKORDINAL) As Long          'Allocate for the machine-code array
    
    ' Create machine-code array
    z_Cb(5) = &HBB60E089: z_Cb(7) = &H73FFC589: z_Cb(8) = &HC53FF04: z_Cb(9) = &H59E80A74: z_Cb(10) = &HE9000000
    z_Cb(11) = &H30&: z_Cb(12) = &H87B81: z_Cb(13) = &H75000000: z_Cb(14) = &H9090902B: z_Cb(15) = &H42DE889: z_Cb(16) = &H50000000: z_Cb(17) = &HB9909090: z_Cb(19) = &H90900AE3
    z_Cb(20) = &H8D74FF: z_Cb(21) = &H9090FAE2: z_Cb(22) = &H53FF33FF: z_Cb(23) = &H90909004: z_Cb(24) = &H2BADC261: z_Cb(25) = &H3D0853FF: z_Cb(26) = &H1&: z_Cb(27) = &H23DCE74: z_Cb(28) = &H74000000: z_Cb(29) = &HAE807
    z_Cb(30) = &H90900000: z_Cb(31) = &H4589C031: z_Cb(32) = &H90DDEBFC: z_Cb(33) = &HFF0C75FF: z_Cb(34) = &H53FF0475: z_Cb(35) = &HC310&

    z_Cb(INDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", False)
    z_Cb(INDX_OWNER) = ObjPtr(oCallback)                  'Set the Owner
    z_Cb(INDX_CALLBACK) = nCallback                       'Set the callback address
    z_Cb(IDX_CALLBACKORDINAL) = nOrdinal                  'Cache ordinal used for zTerminateThunks
      
    If UseIDESafety = True Then                                           'If the user wants IDE protection
        Debug.Assert zInIDE(bIDE)
        If bIDE = True Then z_Cb(INDX_EBMODE) = zFnAddr("vba6", "EbMode", False) 'Store the EbMode function address in the thunk data
    End If
    If bIsTimerCallback Then
        z_Cb(INDX_KT) = zFnAddr("user32", "KillTimer", False)
    End If
        
    z_Cb(INDX_PARAMS) = nParamCount                         'Set the parameter count
    RtlMoveMemory VarPtr(z_Cb(INDX_PARAMLEN)) + 2, VarPtr(nParamCount * 4), 2&
      
    z_Cb(INDX_EBX) = asmMem                                'Set the data address relative to virtual memory pointer
      
    RtlMoveMemory asmMem, VarPtr(z_Cb(INDX_OWNER)), MEM_LEN 'Copy thunk code to executable memory
    scb_SetCallbackAddr = asmMem + PROC_OFF                       'Thunk code start address

End Function

Private Sub scb_ReleaseCallback(ByVal nOrdinal As Long, Optional ByVal oCallback As Object)
    ' can be made public, can be removed & zUnThunk can be called instead
    ' NEVER call this from within the callback routine itself
    
    ' oCallBack is the object containing nOrdinal to be released
    ' if oCallback was already closed (say it was a class or form), then you won't be
    '   able to release it here, but it will be released when zTerminateThunks is
    '   eventually called
    
    ' Special Warning. If the callback thunk is used for a recurring callback (i.e., Timer),
    ' then ensure you terminate what is using the callback before releasing the thunk,
    ' otherwise you are subject to a crash when that item tries to callback to zeroed memory
    zUnThunk nOrdinal, CallbackThunk, oCallback
End Sub


'-The following routines are used for each of the three types of thunks ----------------------------

'Maps zData() to the memory address for the specified thunk type
Private Function zMap_VFunction(vFuncTarget As Long, _
                                vType As eThunkType, _
                                Optional oCallback As Object, _
                                Optional bIgnoreErrors As Boolean) As Long
    
    Dim thunkCol As Collection
    Dim colID As String
    Dim asmMem       As Long         'Thunk base address
    
    If vType = CallbackThunk Then
        Set thunkCol = z_cbFunk
        If oCallback Is Nothing Then Set oCallback = Me
        colID = "h" & ObjPtr(oCallback) & "." & vFuncTarget
    ElseIf vType = HookThunk Then
        Set thunkCol = z_hkFunk
        colID = "h" & vFuncTarget
    ElseIf vType = SubclassThunk Then
        Set thunkCol = asmFunk
        colID = "h" & vFuncTarget
    Else
        zError "zMap_Vfunction", "Invalid thunk type passed"
        Exit Function
    End If
    
    If thunkCol Is Nothing Then
        zError "zMap_VFunction", "Thunk hasn't been initialized"
    Else
        If thunkCol.Count Then
            On Error GoTo Catch
            asmMem = thunkCol(colID)               'Get the thunk address
            If IsBadCodePtr(asmMem) Then asmMem = 0&
            zMap_VFunction = asmMem
        End If
    End If
    Exit Function                                               'Exit returning the thunk address
    
Catch:
    ' error ignored when zUnThunk is called, error handled there
    If Not bIgnoreErrors Then zError "zMap_VFunction", "Thunk type for " & vType & " does not exist"
End Function

' sets/retrieves data at the specified offset for the specified memory address
Private Property Get zData(ByVal nIndex As Long, ByVal asmMem As Long) As Long
  RtlMoveMemory VarPtr(zData), asmMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal asmMem As Long, ByVal nValue As Long)
  RtlMoveMemory asmMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'Error handler
Private Sub zError(ByRef sRoutine As String, ByVal sMsg As String)
  ' Note. These two lines can be rem'd out if you so desire. But don't remove the routine
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long
  If asUnicode Then
    zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
  Else
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
  End If
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
  ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    ' Note: used both in subclassing and hooking routines
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim J     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H710, i, bSub) Then                            'Probe for a PropertyPage method
        If Not zProbe(nAddr + &H7A4, i, bSub) Then                          'Probe for a UserControl method
            Exit Function                                                   'Bail...
        End If
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  J = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < J
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                               'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Do                                                             'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Function zInIDE(ByRef bIDE As Boolean) As Boolean
    ' only called in IDE, never called when compiled
    bIDE = True
    zInIDE = bIDE
End Function

Private Sub zUnThunk(ByVal thunkID As Long, ByVal vType As eThunkType, Optional ByVal oCallback As Object)

    ' thunkID, depends on vType:
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Hooking:      the hook type created
    '   - Callbacks:    the ordinal of the callback
    '       ensure KillTimer is already called, if any callback used for SetTimer
    ' oCallback only used when vType is CallbackThunk

    Const IDX_SHUTDOWN  As Long = 1
    Const MEM_RELEASE As Long = &H8000&             'Release allocated memory flag
    
    Dim asmMem       As Long                       'Thunk base address
    
    asmMem = zMap_VFunction(thunkID, vType, oCallback, True)
    Select Case vType
    Case SubclassThunk
        If asmMem Then                                 'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN, asmMem) = 1            'Set the shutdown indicator
            zDelMsg ALL_MESSAGES, IDX_BTABLE, asmMem   'Delete all before messages
            zDelMsg ALL_MESSAGES, IDX_ATABLE, asmMem   'Delete all after messages
        End If
        asmFunk.Remove "h" & thunkID                   'Remove the specified thunk from the collection
        
    Case HookThunk
        If asmMem Then                                 'Ensure that the thunk hasn't already released its memory
            ' if not unhooked, then unhook now
            If zData(IDX_SHUTDOWN, asmMem) = 0 Then UnhookWindowsHookEx zData(IDX_PREVPROC, asmMem)
            If zData(0, asmMem) = 0 Then               ' not recursing then
                VirtualFree asmMem, 0, MEM_RELEASE     'Release allocated memory
                z_hkFunk.Remove "h" & thunkID           'Remove the specified thunk from the collection
            Else
                zData(IDX_SHUTDOWN, asmMem) = 1        ' Set the shutdown indicator
                zData(IDX_ATABLE, asmMem) = 0          ' want no more After messages
                zData(IDX_BTABLE, asmMem) = 0          ' want no more Before messages
                ' when zTerminate is called this thunk's memory will be released
            End If
        Else
            z_hkFunk.Remove "h" & thunkID       'Remove the specified thunk from the collection
        End If
    Case CallbackThunk
        If asmMem Then                         'Ensure that the thunk hasn't already released its memory
            VirtualFree asmMem, 0, MEM_RELEASE 'Release allocated memory
        End If
        z_cbFunk.Remove "h" & ObjPtr(oCallback) & "." & thunkID           'Remove the specified thunk from the collection
    End Select

End Sub

Private Sub zTerminateThunks(ByVal vType As eThunkType)

    ' Terminates all thunks of a specific type
    ' Any subclassing, hooking, recurring callbacks should have already been canceled

    Dim i As Long
    Dim oCallback As Object
    Dim thunkCol As Collection
    Dim asmMem       As Long                           'Thunk base address
    Const INDX_OWNER As Long = 0
    
    Select Case vType
    Case SubclassThunk
        Set thunkCol = asmFunk
    Case HookThunk
        Set thunkCol = z_hkFunk
    Case CallbackThunk
        Set thunkCol = z_cbFunk
    Case Else
        Exit Sub
    End Select
    
    If Not (thunkCol Is Nothing) Then                 'Ensure that hooking has been started
      With thunkCol
        For i = .Count To 1 Step -1                   'Loop through the collection of hook types in reverse order
          asmMem = .Item(i)                          'Get the thunk address
          If IsBadCodePtr(asmMem) = 0 Then           'Ensure that the thunk hasn't already released its memory
            Select Case vType
                Case SubclassThunk
                    zUnThunk zData(IDX_INDEX, asmMem), SubclassThunk    'Unsubclass
                Case HookThunk
                    zUnThunk zData(IDX_INDEX, asmMem), HookThunk        'Unhook
                Case CallbackThunk
                    ' zUnThunk expects object not pointer, convert pointer to object
                    RtlMoveMemory VarPtr(oCallback), VarPtr(zData(INDX_OWNER, asmMem)), 4&
                    zUnThunk zData(IDX_CALLBACKORDINAL, asmMem), CallbackThunk, oCallback ' release callback
                    ' remove the object pointer reference
                    RtlMoveMemory VarPtr(oCallback), VarPtr(INDX_OWNER), 4&
            End Select
          End If
        Next                                         'Next member of the collection
      End With
      Set thunkCol = Nothing                         'Destroy the hook/thunk-address collection
    End If


End Sub

'//////////////////////////////////////////////////////////////////////////////////////

' EXAMPLE FUNCTIONS/CALLBACKS. Examples pulled from various postings

' ordinal #5 - Example of a window enumeration callback used with scb_SetCallbackAddr
' http://msdn2.microsoft.com/en-us/library/aa922950.aspx
Private Function myEnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    
'    Debug.Print hWnd
'    myEnumWindowsProc = 1 'Continue

End Function


' ordinal #4 - Example of font callback procedure used with scb_SetCallbackAddr
' http://msdn2.microsoft.com/en-us/library/Aa911409.aspx

' the LOGFONT and NEWTEXTMETRIC UDTs not declared in this template. Add this routine to a new class and unrem it for testing
'Private Function myEnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, ByVal lParam As Long) As Long
'
'    ' Very cool; note that the call backs can receive UDTs too, just ensure they are ByRef and not ByVal
'    Dim FaceName As String
'
'    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
'    FaceName = Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
'    Debug.Print FaceName
'
'    myEnumFontFamProc = 1 'Continue
'
'End Function


' ordinal #3 - Example of a timer procedure callback used with scb_SetCallbackAddr
' http://msdn2.microsoft.com/en-us/library/ms644907.aspx
Private Function myTimerProc(ByVal hWnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long
    
    ' note: although a timer procedure, per MSDN, does not return a value,
    ' the function that is used for callbacks must return a value therefore
    ' all callback routines must be functions, even if the return value is not used
                        ' YOUR CODE HERE
                        
    ' Note that the scb_SetCallbackAddr has an optional parameter for timers created by
    ' SetTimer. Set that parameter to true to provide extra protection while in IDE.
    ' See scb_SetCallbackAddr & scb_ReleaseCallback

End Function

' ordinal #2 ' Example of a hook procedure used with shk_SetHook
Private Sub myHookProc(ByVal bBefore As Boolean, _
                        ByRef bHandled As Boolean, _
                        ByRef lReturn As Long, _
                        ByVal nCode As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long, _
                        ByVal lHookType As eHookType, _
                        ByRef lParamUser As Long)
'*************************************************************************************************
' http://msdn2.microsoft.com/en-us/library/ms644990.aspx
'* bBefore    - Indicates whether the callback is before or after the next hook in chain.
'* bHandled   - In a before next hook in chain callback, setting bHandled to True will prevent the
'*              message being passed to the next hook in chain and (if set to do so).
'* lReturn    - Return value. For Before messages, set per the MSDN documentation for the hook type
'* nCode      - A code the hook procedure uses to determine how to process the message
'* wParam     - Message related data, hook type specific
'* lParam     - Message related data, hook type specific
'* lHookType  - Type of hook calling this callback
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************
    
    ' note. All hook type procedures are identically declared like this one;
    '   however, the meaning of nCode, wParam, lParam are specific to the type of hook

                        ' YOUR CODE HERE

    ' nice to know stuff if hooking the keyboard
    ' http://msdn2.microsoft.com/en-us/library/ms644984.aspx
    If lHookType = WH_KEYBOARD Then
        If nCode = 0 Then 'HC_ACTION
            If lParam < 0 Then          'wParam are vbKey... variables
                Debug.Print "Keycode: "; wParam; Chr$(wParam), "key is up"
            Else
                If (lParam \ &H40000000) Then
                    Debug.Print "Keycode: "; wParam; Chr$(wParam), "key is held down"
                Else
                    Debug.Print "keycode: "; wParam; Chr$(wParam), "key is down first time"
                End If
                ' to test if ALT is also pressed: (((lParam \ &H20000000) And 1) = 1)
                ' to test if extended key is pressed: (((lParam \ &H1000000) And 1) = 1)
            End If
        End If
    End If


End Sub

'- ordinal #1, example of a subclassing procedure used with Subclass
Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc.
'*              Not applicable with After messages
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************
    
                        ' YOUR CODE HERE
    
' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
'   add this warning banner to the last routine in your class
' *************************************************************
End Sub

