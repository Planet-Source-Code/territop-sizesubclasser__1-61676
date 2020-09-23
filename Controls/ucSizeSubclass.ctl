VERSION 5.00
Begin VB.UserControl ucSizeSubclass 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ucSizeSubclass.ctx":0000
   ScaleHeight     =   16
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   16
   ToolboxBitmap   =   "ucSizeSubclass.ctx":0342
End
Attribute VB_Name = "ucSizeSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+  File Description:
'       ucSizeSubclass - Size Subclasser to provide Flicker-Free Size Restrictions
'
'   Product Name:
'       ucSizeSubclass.ctl
'
'   Compatability:
'       Windows: 98, ME, NT4, 2000, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'       Adapted from the following online article(s):
'       Based in large part from Paul Caton's Self-Subclassing Example (see URL below)...
'       http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'       http://www.vb-helper.com/howtoint.htm (Article: "Find the system's color depth (bits per pixel)")
'
'   Legal Copyright & Trademarks (Current Implementation):
'       Copyright © 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul R. Territo, Ph.D shall not be liable
'       for any incidental or consequential damages suffered by any use of
'       this software.
'
'-  Modification(s) History:
'       23Jun05 - Initial build of the SizeSubclass Control
'       23Jun05 - Fixed bug with Screen size routine which reported the
'                 incorrect values of pixels
'       05Jul05 - Added additional public events and properties
'       07Jul05 - Added Error handling for multiple control instances loaded on
'                 the same form at on time.
'       12Jul05 - Added additional error checking for previous existance based on
'                 suggestions by LaVolpe and Fred.cpp. Current version now has a
'                 Public Enabled property which checks for other instances when
'                 it is set to true.
'               - Added MDI Form support and parent form subclassing for QueryClose
'                 events to make sure the Subclasser is shutdown correctly.
'               - Added User feedback to the user by "X"ing out the controls
'                 GUI when disabled...
'
'   Force Declarations
Option Explicit

'==================================================================================================
'Application declarations
'==================================================================================================
Private Const WM_ACTIVATEAPP        As Long = &H1C
Private Const WM_DISPLAYCHANGE      As Long = &H7E
Private Const WM_EXITSIZEMOVE       As Long = &H232
Private Const WM_GETMINMAXINFO      As Long = &H24
Private Const WM_MOUSELEAVE         As Long = &H2A3
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_MOVING             As Long = &H216
Private Const WM_SIZING             As Long = &H214
Private Const WM_SYSCOLORCHANGE     As Long = &H15
Private Const WM_THEMECHANGED       As Long = &H31A
Private Const BITSPIXEL             As Long = 12

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                          As Long
    dwFlags                         As TRACKMOUSEEVENT_FLAGS
    hwndTrack                       As Long
    dwHoverTime                     As Long
End Type

Private Type POINTAPI
    X                               As Long
    Y                               As Long
End Type

Private Type MINMAXINFO
    ptReserved                      As POINTAPI
    ptMaxSize                       As POINTAPI
    ptMaxPosition                   As POINTAPI
    ptMinTrackSize                  As POINTAPI
    ptMaxTrackSize                  As POINTAPI
End Type

Private Enum pDirctionEnum
    pXDirection = 0
    pYDirection = 1
End Enum

Private bTrack                      As Boolean
Private bTrackUser32                As Boolean
Private SizeInfo                    As MINMAXINFO
Private m_DisplayColorDepth         As Long
Private m_DisplayChanged            As Boolean
Private m_DisplayHeight             As Long
Private m_DisplayWidth              As Long
Private m_Enabled                   As Boolean
Private m_MaxHeight                 As Long
Private m_MinHeight                 As Long
Private m_MaxWidth                  As Long
Private m_MinWidth                  As Long
Private m_MaximizedHeight           As Long
Private m_MaximizedWidth            As Long
Private m_MaximizedXOffset          As Long
Private m_MaximizedYOffset          As Long
Private m_SysColorChanged           As Boolean
Private m_SysThemeChanged           As Boolean

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Event DisplayChanged()
Public Event SysColorChanged()
Public Event SysThemeChanged()
Public Event ParentMoving()
Public Event ParentSizing()
Public Event FinishedSizeMove()

Private WithEvents SDIParentForm As Form
Attribute SDIParentForm.VB_VarHelpID = -1
Private WithEvents MDIParentForm As MDIForm
Attribute MDIParentForm.VB_VarHelpID = -1
'==================================================================================================
'==================================================================================================
' ucSubclass - A template UserControl for control authors that require self-subclassing without ANY
'              external dependencies. IDE safe.
'
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.0.0000 20040525 First cut.....................................................................
' v1.1.0000 20040602 Multi-subclassing version.....................................................
' v1.1.0001 20040604 Optimized the subclass code...................................................
' v1.1.0002 20040607 Substituted byte arrays for strings for the code buffers......................
' v1.1.0003 20040618 Re-patch when adding extra hWnds..............................................
' v1.1.0004 20040619 Optimized to death version....................................................
' v1.1.0005 20040620 Use allocated memory for code buffers, no need to re-patch....................
' v1.1.0006 20040628 Better protection in zIdx, improved comments..................................
' v1.1.0007 20040629 Fixed InIDE patching oops.....................................................
' v1.1.0008 20040910 Fixed bug in UserControl_Terminate, zSubclass_Proc procedure hidden...........
'==================================================================================================
'Subclasser Declarations
'==================================================================================================
Private Enum eMsgWhen
    MSG_AFTER = 1                                           'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                          'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE          'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES          As Long = -1            'All messages added or deleted
Private Const GMEM_FIXED            As Long = 0             'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC           As Long = -4            'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04              As Long = 88            'Table B (before) address patch offset
Private Const PATCH_05              As Long = 93            'Table B (before) entry count patch offset
Private Const PATCH_08              As Long = 132           'Table A (after) address patch offset
Private Const PATCH_09              As Long = 137           'Table A (after) entry count patch offset

Private Type tSubData                                       'Subclass data type
    hWnd                            As Long                 'Handle of the window being subclassed
    nAddrSub                        As Long                 'The address of our new WndProc (allocated memory).
    nAddrOrig                       As Long                 'The address of the pre-existing WndProc
    nMsgCntA                        As Long                 'Msg after table entry count
    nMsgCntB                        As Long                 'Msg before table entry count
    aMsgTblA()                      As Long                 'Msg after table array
    aMsgTblB()                      As Long                 'Msg Before table array
End Type

Private sc_aSubData()               As tSubData             'Subclass data array

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also
'======================================================================================================
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    'Parameters:
    'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
    'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
    'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
    'hWnd     - The window handle
    'uMsg     - The message number
    'wParam   - Message related data
    'lParam   - Message related data
    Static bTheme As Boolean
    Static bInForm As Boolean
    
    Select Case uMsg
        'Application is being activated/deactivated
    Case WM_ACTIVATEAPP
        'Debug.Print "Application " & IIf(wParam = 0, "deactivated", "activated")
        
        'Screen resolution and or color depth has changed
    Case WM_DISPLAYCHANGE
        'Debug.Print "Display changed: " & WordLo(lParam) & "x" & WordHi(lParam) & ", bpp = " & wParam
        m_DisplayChanged = True
        RaiseEvent DisplayChanged
        m_DisplayHeight = WordHi(lParam)
        m_DisplayWidth = WordLo(lParam)
        m_DisplayColorDepth = wParam
        PropertyChanged "DisplayHeight"
        PropertyChanged "DisplayWidth"
        PropertyChanged "DisplayColorDepth"
        'The user has ceased sizing or moving
    Case WM_EXITSIZEMOVE
        'Debug.Print vbNullString
        RaiseEvent FinishedSizeMove
        
        'OS is asking us for min/max info
    Case WM_GETMINMAXINFO
        Call RtlMoveMemory(ByVal lParam, SizeInfo, LenB(SizeInfo))
        
        'Mouse has left the tracked window
    Case WM_MOUSELEAVE
        'Debug.Print "ParentForm Mouse Leave:"
        bInForm = False

        'Mouse has moved
    Case WM_MOUSEMOVE
        If lng_hWnd = UserControl.Parent.hWnd Then
            'Debug.Print "ParentForm Mouse Move: " & WordLo(lParam) & "," & WordHi(lParam)
            If Not bInForm Then
                bInForm = True
                Call TrackMouseLeave(UserControl.Parent.hWnd)
            End If
        Else
            'Debug.Print "hWnd " & Hex$(lng_hWnd) & " - mouse move: " & WordLo(lParam) & "," & WordHi(lParam)
        End If
        
        'Window is being moved
    Case WM_MOVING
        'Debug.Print "Moving..."
        RaiseEvent ParentMoving
        
        'Window is being sized
    Case WM_SIZING
        'Debug.Print "Sizing..."
        RaiseEvent ParentSizing
        
        'The system colors have been changed
    Case WM_SYSCOLORCHANGE
        If bTheme Then
            bTheme = False
            'Debug.Print "XP theme and system colors changed"
            m_SysColorChanged = True
            m_SysThemeChanged = True
            RaiseEvent SysColorChanged
            RaiseEvent SysThemeChanged
        Else
            'Debug.Print "System colors changed"
            m_SysColorChanged = True
            RaiseEvent SysColorChanged
        End If
        
        'The Windows XP theme has been changed
    Case WM_THEMECHANGED
        'Theme changes are almost bound to change the system colors, the theme change message comes first,
        'therefore I'm setting a flag so that when the WM_SYSCOLORCHANGED message comes microseconds after
        'that we don't miss the theme change message in the status bar.
        bTheme = True
        'Debug.Print "XP theme changed"
        m_SysThemeChanged = True
        RaiseEvent SysColorChanged
    End Select
    
    'Notes:
    'If you really know what you're doing, it's possible to change the values of the
    'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
    'values get passed to the default handler.. and optionaly, the 'after' callback
End Sub

Private Function ConvertScale(inValue As Long, pDirection As pDirctionEnum) As Long
    Select Case pDirection
        Case pXDirection
            ConvertScale = inValue \ Screen.TwipsPerPixelX
        Case pYDirection
            ConvertScale = inValue \ Screen.TwipsPerPixelY
    End Select
End Function

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    m_Enabled = Value
    '   Check for other instances of this control...
    If Value = True Then
        Call VerifyExistance
    End If
    PropertyChanged "Enabled"
End Property

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    Dim hMod        As Long
    Dim bLibLoaded  As Boolean
    
    hMod = GetModuleHandleA(sModule)
    
    If hMod = 0 Then
        hMod = LoadLibraryA(sModule)
        If hMod Then
            bLibLoaded = True
        End If
    End If
    
    If hMod Then
        If GetProcAddress(hMod, sFunction) Then
            IsFunctionExported = True
        End If
    End If
    
    If bLibLoaded Then
        Call FreeLibrary(hMod)
    End If
End Function

Public Property Get DisplayColorDepth() As Long
    DisplayColorDepth = m_DisplayColorDepth
End Property

Public Property Get DisplayChanged() As Boolean
    DisplayChanged = m_DisplayChanged
End Property

Public Property Get DisplayHeight() As Long
    DisplayHeight = m_DisplayHeight
End Property

Public Property Get DisplayWidth() As Long
    DisplayWidth = m_DisplayWidth
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get MaxHeight() As Long
    MaxHeight = m_MaxHeight
End Property

Public Property Let MaxHeight(New_Size As Long)
    m_MaxHeight = New_Size
    SizeInfo.ptMaxTrackSize.Y = New_Size
    PropertyChanged "MaxHeight"
End Property

Public Property Get MinHeight() As Long
    MinHeight = m_MinHeight
End Property

Public Property Let MinHeight(New_Size As Long)
    m_MinHeight = New_Size
    SizeInfo.ptMinTrackSize.Y = New_Size
    PropertyChanged "MinHeight"
End Property

Public Property Get MaxWidth() As Long
    MaxWidth = m_MaxWidth
End Property

Public Property Let MaxWidth(New_Size As Long)
    m_MaxWidth = New_Size
    SizeInfo.ptMaxTrackSize.X = New_Size
    PropertyChanged "MaxWidth"
End Property

Public Property Get MinWidth() As Long
    MinWidth = m_MinWidth
End Property

Public Property Let MinWidth(New_Size As Long)
    m_MinWidth = New_Size
    SizeInfo.ptMinTrackSize.X = New_Size
    PropertyChanged "MinWidth"
End Property

Public Property Get MaximizedHeight() As Long
    MaximizedHeight = m_MaximizedHeight
End Property

Public Property Let MaximizedHeight(New_Size As Long)
    m_MaximizedHeight = New_Size
    SizeInfo.ptMaxSize.Y = New_Size
    PropertyChanged "MaximizedHeight"
End Property

Public Property Get MaximizedWidth() As Long
    MaximizedWidth = m_MaximizedWidth
End Property

Public Property Let MaximizedWidth(New_Size As Long)
    m_MaximizedWidth = New_Size
    SizeInfo.ptMaxSize.X = New_Size
    PropertyChanged "MaximizedWidth"
End Property

Public Property Get MaximizedXOffset() As Long
    MaximizedXOffset = m_MaximizedXOffset
End Property

Public Property Let MaximizedXOffset(New_Size As Long)
    m_MaximizedXOffset = New_Size
    SizeInfo.ptMaxPosition.X = New_Size
    PropertyChanged "MaximizedXOffset"
End Property

Public Property Get MaximizedYOffset() As Long
    MaximizedYOffset = m_MaximizedYOffset
End Property

Public Property Let MaximizedYOffset(New_Size As Long)
    m_MaximizedYOffset = New_Size
    SizeInfo.ptMaxPosition.Y = New_Size
    PropertyChanged "MaximizedYOffset"
End Property

Private Sub PaintDisabledImage()
    Dim OldWidth        As Long
    Dim OldColor        As Long
    
    '   Draw a Red "X" on the control surface when Enabled = False
    With UserControl
        '   Save the old settings for later
        OldWidth = .DrawWidth
        OldColor = .ForeColor
        .DrawWidth = 3
        .ForeColor = &HFF
        '   Draw UL to LR Line
        UserControl.Line (0, 0)-(.ScaleWidth, .ScaleHeight)
        '   Drae LL to UR Line
        UserControl.Line (0, .ScaleHeight)-(.ScaleWidth, 0)
        '   Set them back...
        .DrawWidth = OldWidth
        .ForeColor = OldColor
    End With
End Sub

Public Property Get SysColorChanged() As Boolean
    SysColorChanged = m_SysColorChanged
End Property

Public Property Get SysThemeChanged() As Boolean
    SysThemeChanged = m_SysThemeChanged
End Property
'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
    'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
    'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
    'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
    'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
    'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
    'When      - Whether the msg is to be removed from the before, after or both callback tables
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
    'Parameters:
    'lng_hWnd  - The handle of the window to be subclassed
    'Returns;
    'The sc_aSubData() index
    Const CODE_LEN              As Long = 200                   'Length of the machine code in bytes
    Const FUNC_CWP              As String = "CallWindowProcA"   'We use CallWindowProc to call the original WndProc
    Const FUNC_EBM              As String = "EbMode"            'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL              As String = "SetWindowLongA"    'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_USER              As String = "user32"            'Location of the SetWindowLongA & CallWindowProc functions
    Const MOD_VBA5              As String = "vba5"              'Location of the EbMode function if running VB5
    Const MOD_VBA6              As String = "vba6"              'Location of the EbMode function if running VB6
    Const PATCH_01              As Long = 18                    'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02              As Long = 68                    'Address of the previous WndProc
    Const PATCH_03              As Long = 78                    'Relative address of SetWindowsLong
    Const PATCH_06              As Long = 116                   'Address of the previous WndProc
    Const PATCH_07              As Long = 121                   'Relative address of CallWindowProc
    Const PATCH_0A              As Long = 186                   'Address of the owner object
    Static aBuf(1 To CODE_LEN)  As Byte                         'Static code buffer byte array
    Static pCWP                 As Long                         'Address of the CallWindowsProc
    Static pEbMode              As Long                         'Address of the EbMode IDE break/stop/running function
    Static pSWL                 As Long                         'Address of the SetWindowsLong function
    Dim i                       As Long                         'Loop index
    Dim j                       As Long                         'Loop index
    Dim nSubIdx                 As Long                         'Subclass data index
    Dim sHex                    As String                       'Hex code string
    
    
    'If it's the first time through here..
    If aBuf(1) = 0 Then
        
        'The hex pair machine code representation.
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
        "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
        "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
        "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        
        'Convert the string from hex pairs to bytes and store in the static machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))              'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                    'Next pair of hex characters
        
        'Get API function addresses
        If Subclass_InIDE Then                                  'If we're running in the VB IDE
            aBuf(16) = &H90                                     'Patch the code buffer to enable the IDE state code
            aBuf(17) = &H90                                     'Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)             'Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then                                 'Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)         'VB5 perhaps
            End If
        End If

        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                    'Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                    'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                   'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then                                    'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                 'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData 'Create a new sc_aSubData element
        End If

        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hWnd = lng_hWnd                                        'Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)           'Allocate memory for the machine code WndProc
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub) 'Set our WndProc in place
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)  'Copy the machine code from the static byte array to the code array in sc_aSubData
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)            'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)         'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)               'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)         'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)               'Patch the relative address of the CallWindowProc api function
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))         'Patch the address of this object instance into the static machine code buffer
    End With
    
End Function

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
    'Parameters:
    'lng_hWnd  - The handle of the window to stop being subclassed
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)     'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                  'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                  'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                              'Release the machine code memory
        .hWnd = 0                                               'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                           'Clear the before table
        .nMsgCntA = 0                                           'Clear the after table
        Erase .aMsgTblB                                         'Erase the before table
        Erase .aMsgTblA                                         'Erase the after table
    End With
End Sub

'Stop all subclassing
Private Sub Subclass_StopAll()
    Dim i As Long
    
    i = UBound(sc_aSubData())                                   'Get the upper bound of the subclass data array
    Do While i >= 0                                             'Iterate through each element
        With sc_aSubData(i)
            If .hWnd <> 0 Then                                  'If not previously Subclass_Stop'd
            Call Subclass_Stop(.hWnd)                           'Subclass_Stop
        End If
    End With
    
    i = i - 1                                                   'Next element
Loop
End Sub

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
    Dim tme As TRACKMOUSEEVENT_STRUCT
    
    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With
        
        If bTrackUser32 Then
            Call TrackMouseEvent(tme)
        Else
            Call TrackMouseEventComCtl(tme)
        End If
    End If
End Sub

Private Sub MDIParentForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '   This event is required to prevent GPF in the IDE if the user
    '   has loaded more than 1 control at a time on the form...
    If Enabled Then
        Call Subclass_StopAll
    End If
End Sub

Private Sub SDIParentForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '   This event is required to prevent GPF in the IDE if the user
    '   has loaded more than 1 control at a time on the form...
    If Enabled Then
        Call Subclass_StopAll
    End If
End Sub

Private Sub UserControl_InitProperties()
    Call VerifyExistance
    '   All Property Dimensions are in Pixels, which is
    '   what the system uses...
    m_DisplayColorDepth = GetDeviceCaps(hdc, BITSPIXEL)
    m_DisplayHeight = Screen.Height
    m_DisplayWidth = Screen.Width
    m_MinHeight = 0
    m_MaxHeight = Screen.Height
    m_MinWidth = 0
    m_MaxWidth = Screen.Width
    m_MaximizedHeight = Screen.Height
    m_MaximizedWidth = Screen.Width
    m_MaximizedXOffset = 0
    m_MaximizedYOffset = 0
End Sub

'Read the properties from the property bag - also, a good place to start the subclassing (if we're running)
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim Ctl As Control
    
    With PropBag
        'Read your properties here
        m_Enabled = .ReadProperty("Enabled", True)
        m_DisplayColorDepth = .ReadProperty("DisplayColorDepth", GetDeviceCaps(hdc, BITSPIXEL))
        m_DisplayHeight = .ReadProperty("DisplayHeight", Screen.Height)
        m_DisplayWidth = .ReadProperty("DisplayWidth", Screen.Width)
        m_MinHeight = .ReadProperty("MinHeight", 0)
        m_MaxHeight = .ReadProperty("MaxHeight", Screen.Height)
        m_MinWidth = .ReadProperty("MinWidth", 0)
        m_MaxWidth = .ReadProperty("MaxWidth", Screen.Width)
        m_MaximizedHeight = .ReadProperty("MaximizedHeight", Screen.Height)
        m_MaximizedWidth = .ReadProperty("MaximizedWidth", Screen.Width)
        m_MaximizedXOffset = .ReadProperty("MaximizedXOffset", 0)
        m_MaximizedYOffset = .ReadProperty("MaximizedYOffset", 0)
        If UserControl.Ambient.UserMode Then
            '   Reference the parent form and start recieving events
            If Not TypeOf UserControl.Parent Is MDIForm Then
                '   Single Document Interfaces only....
                Set SDIParentForm = UserControl.Parent
            Else
                '   Multiple Document Interfaces....
                Set MDIParentForm = UserControl.Parent
            End If
            '   Check to see that we only have one instance running!
            If (Not m_Enabled) Then
                Call VerifyExistance
            End If
        End If
    End With
    
    '   Are we in design mode and is the control enbled?
    If (Ambient.UserMode) And (Enabled = True) Then
        '   Print the parent controls name
        Debug.Print UserControl.Extender.Name
        '   Initialize the form's size information
        With SizeInfo
            '   Maximised position
            With .ptMaxPosition
                .X = ConvertScale(m_MaximizedXOffset, pXDirection)
                .Y = ConvertScale(m_MaximizedYOffset, pYDirection)
            End With
            '   Maximized size
            With .ptMaxSize
                .X = ConvertScale(m_MaximizedWidth, pXDirection)
                .Y = ConvertScale(m_MaximizedHeight, pYDirection)
            End With
            '   Maximum size while re-sizing (Dragging)
            With .ptMaxTrackSize
                .X = ConvertScale(m_MaxWidth, pXDirection)
                .Y = ConvertScale(m_MaxHeight, pYDirection)
            End With
            '   Minimum size while re-sizing (Dragging)
            With .ptMinTrackSize
                .X = ConvertScale(m_MinWidth, pXDirection)
                .Y = ConvertScale(m_MinHeight, pYDirection)
            End With
        End With
        
        With UserControl.Parent
            '   If we're not in design mode
            '   Start subclassing the Form
            Call Subclass_Start(.hWnd)
            
            '   Add the messages that we're interested in
            Call Subclass_AddMsg(.hWnd, WM_ACTIVATEAPP)
            Call Subclass_AddMsg(.hWnd, WM_DISPLAYCHANGE)
            Call Subclass_AddMsg(.hWnd, WM_EXITSIZEMOVE)
            Call Subclass_AddMsg(.hWnd, WM_GETMINMAXINFO)
            Call Subclass_AddMsg(.hWnd, WM_MOVING)
            Call Subclass_AddMsg(.hWnd, WM_SIZING)
            Call Subclass_AddMsg(.hWnd, WM_SYSCOLORCHANGE)
            Call Subclass_AddMsg(.hWnd, WM_THEMECHANGED)
        End With 'Usercontrol.Parent
    End If
End Sub

Private Sub UserControl_Resize()
    With UserControl
        .Width = 240
        .Height = 240
    End With
End Sub

Private Sub UserControl_Show()
    If Not m_Enabled Then
        '   Provide feedback about the disabled state...
        Call PaintDisabledImage
    End If
End Sub

'The control is terminating - a good place to stop the subclasser
Private Sub UserControl_Terminate()
    On Error GoTo Catch
    'Stop all subclassing
    Call Subclass_StopAll
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Enabled", m_Enabled, True)
        Call .WriteProperty("MinHeight", m_MinHeight, 0)
        Call .WriteProperty("MaxHeight", m_MaxHeight, Screen.Height)
        Call .WriteProperty("MinWidth", m_MinWidth, 0)
        Call .WriteProperty("MaxWidth", m_MaxWidth, Screen.Width)
        Call .WriteProperty("MaximizedHeight", m_MaximizedHeight, Screen.Height)
        Call .WriteProperty("MaximizedWidth", m_MaximizedWidth, Screen.Width)
        Call .WriteProperty("MaximizedXOffset", m_MaximizedXOffset, 0)
        Call .WriteProperty("MaximizedYOffset", m_MaximizedYOffset, 0)
    End With
End Sub

'Return the upper 16 bits of the passed 32 bit value
Private Function WordHi(lngValue As Long) As Long
    If (lngValue And &H80000000) = &H80000000 Then
        WordHi = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
    Else
        WordHi = (lngValue And &HFFFF0000) \ &H10000
    End If
End Function

'Return the lower 16 bits of the passed 32 bit value
Private Function WordLo(lngValue As Long) As Long
    WordLo = (lngValue And &HFFFF&)
End Function

Private Sub VerifyExistance()
    Dim Ctl     As Control
    Dim i       As Long
    
    '   We are in desgin mode so check for other instances of the
    '   SizeSubclasser!
    If Not UserControl.Ambient.UserMode Then
        For Each Ctl In UserControl.Parent.Controls
            If TypeOf Ctl Is ucSizeSubclass Then
                i = i + 1
                If (Ctl.hWnd = UserControl.hWnd) And (i > 1) Then
                    '   We have more than one control loaded!
                    MsgBox "     Only one Control Permitted Per Form!" & vbCrLf & vbCrLf & "SubClassing Wiil Be Disabled for " & Ctl.Name, vbExclamation, "ucSizeSubclass"
                    Call PaintDisabledImage
                    Me.Enabled = False
                    Exit Sub
                End If
                m_Enabled = True
            End If
        Next Ctl
    Else
    '   We are running the control now, so set all but the first
    '   control to disabled, so we don't cause a GPF of the Application
        For Each Ctl In UserControl.Parent.Controls
            If TypeOf Ctl Is ucSizeSubclass Then
                i = i + 1
                If (Ctl.hWnd <> UserControl.hWnd) And (i > 1) Then
                    '   We have more than one control loaded!
                    '   So disable the other controls from here...
                    Ctl.Enabled = False
                End If
            End If
        Next Ctl
    End If
    
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry  As Long                                         'Message table entry index
    Dim nOff1   As Long                                         'Machine code buffer offset 1
    Dim nOff2   As Long                                         'Machine code buffer offset 2
    
    If uMsg = ALL_MESSAGES Then                                 'If all messages
        nMsgCnt = ALL_MESSAGES                                  'Indicates that all messages will callback
    Else                                                        'Else a specific message number
        Do While nEntry < nMsgCnt                               'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1
            
            If aMsgTbl(nEntry) = 0 Then                         'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                          'Re-use this entry
                Exit Sub                                        'Bail
            ElseIf aMsgTbl(nEntry) = uMsg Then                  'The msg is already in the table!
                Exit Sub                                        'Bail
            End If
        Loop                                                    'Next entry
        nMsgCnt = nMsgCnt + 1                                   'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long            'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                 'Store the message number in the table
    End If

    If When = eMsgWhen.MSG_BEFORE Then                          'If before
        nOff1 = PATCH_04                                        'Offset to the Before table
        nOff2 = PATCH_05                                        'Offset to the Before table entry count
    Else                                                        'Else after
        nOff1 = PATCH_08                                        'Offset to the After table
        nOff2 = PATCH_09                                        'Offset to the After table entry count
    End If

    If uMsg <> ALL_MESSAGES Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))        'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                       'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                              'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long
    
    If uMsg = ALL_MESSAGES Then                         'If deleting all messages
        nMsgCnt = 0                                     'Message count is now zero
        If When = eMsgWhen.MSG_BEFORE Then              'If before
            nEntry = PATCH_05                           'Patch the before table message count location
        Else                                            'Else after
            nEntry = PATCH_09                           'Patch the after table message count location
        End If
        Call zPatchVal(nAddr, nEntry, 0)                'Patch the table message count to zero
    Else                                                'Else deleteting a specific message
        Do While nEntry < nMsgCnt                       'For each table entry
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then              'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                     'Mark the table slot as available
                Exit Do                                 'Bail
            End If
        Loop                                            'Next entry
    End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                  'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If .hWnd = lng_hWnd Then                    'If the hWnd of this element is the one we're looking for
                If Not bAdd Then                        'If we're searching not adding
                    Exit Function                       'Found
                End If
            ElseIf .hWnd = 0 Then                       'If this an element marked for reuse.
                If bAdd Then                            'If we're adding
                    Exit Function                       'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                                 'Decrement the index
    Loop

    If Not bAdd Then
        'hWnd not found, programmer error
        Debug.Assert False
    End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function

