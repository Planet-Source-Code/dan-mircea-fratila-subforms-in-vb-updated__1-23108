VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScrollBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===========================================================================
' Name:     cScrollBars
' Author:   Steve McMahon (steve@dogma.demon.co.uk)
' Date:     24 December 1998
' Requires: SSUBTMR.DLL
'
' ---------------------------------------------------------------------------
' Copyright ® 1998 Steve McMahon (steve@dogma.demon.co.uk)
' Visit vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' ---------------------------------------------------------------------------
'
' Description:
' A class which can add scroll bars to VB Forms, Picture Boxes and
' UserControls.
' Features:
'  * True API scroll bars, which don't flash or draw badly like
'    the VB ones
'  * Scroll bar values are long integers, i.e. >2 billion values
'  * Set Flat or Encarta scroll bar modes if your COMCTL32.DLL version
'    supports it (>4.72)
'
' FREE SOURCE CODE! - ENJOY.
' - Please report bugs to the author for incorporation into future releases
' - Don't sell this code.
' ===========================================================================

'private declare function InitializeFlatSB(HWND) as long
Private Declare Function InitialiseFlatSB Lib "COMCTL32.DLL" Alias "InitializeFlatSB" (ByVal lHWnd As Long) As Long

' Memory functions:
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' Window style bit functions:
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
    ) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long _
    ) As Long
' Window Long indexes:
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_HINSTANCE = (-6)
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_ID = (-12)
Private Const GWL_STYLE = (-16)
Private Const GWL_USERDATA = (-21)
Private Const GWL_WNDPROC = (-4)

' Creating new windows:
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
' General window styles:
Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Private Const WS_CHILD = &H40000000
Private Const WS_CHILDWINDOW = (WS_CHILD)
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_DISABLED = &H8000000
Private Const WS_DLGFRAME = &H400000
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_TRANSPARENT = &H20&
Private Const WS_GROUP = &H20000
Private Const WS_HSCROLL = &H100000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_SYSMENU = &H80000
Private Const WS_TABSTOP = &H10000
Private Const WS_THICKFRAME = &H40000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_VISIBLE = &H10000000
Private Const WS_VSCROLL = &H200000
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)

' Window appearance control:
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
' Show window styles
Private Const SW_SHOWNORMAL = 1
Private Const SW_ERASE = &H4
Private Const SW_HIDE = 0
Private Const SW_INVALIDATE = &H2
Private Const SW_MAX = 10
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_NORMAL = 1
Private Const SW_OTHERUNZOOM = 4
Private Const SW_OTHERZOOM = 2
Private Const SW_PARENTCLOSING = 1
Private Const SW_RESTORE = 9
Private Const SW_PARENTOPENING = 3
Private Const SW_SHOW = 5
Private Const SW_SCROLLCHILDREN = &H1
Private Const SW_SHOWDEFAULT = 10
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_SHOWNOACTIVATE = 4
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const HWND_NOTOPMOST = -2

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

' Window relationship functions:
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_MAX = 5
Private Const GW_OWNER = 4
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

' Message functions:
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Key functions:
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' Font
Private Const LF_FACESIZE = 32
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1
Private Declare Function CreateFontIndirect& Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT)
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Const BITSPIXEL = 12
    Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
    Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y


Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXHSCROLL = 21
Private Const SM_CYHSCROLL = 3
Private Const SM_CYVSCROLL = 20
Private Const SM_CYVTHUMB = 9
Private Const SM_CXVSCROLL = 2
Private Const SM_CXHTHUMB = 10

' Scroll bar:
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Private Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
    Private Const SB_BOTH = 3
    Private Const SB_BOTTOM = 7
    Private Const SB_CTL = 2
    Private Const SB_ENDSCROLL = 8
    Private Const SB_HORZ = 0
    Private Const SB_LEFT = 6
    Private Const SB_LINEDOWN = 1
    Private Const SB_LINELEFT = 0
    Private Const SB_LINERIGHT = 1
    Private Const SB_LINEUP = 0
    Private Const SB_PAGEDOWN = 3
    Private Const SB_PAGELEFT = 2
    Private Const SB_PAGERIGHT = 3
    Private Const SB_PAGEUP = 2
    Private Const SB_RIGHT = 7
    Private Const SB_THUMBPOSITION = 4
    Private Const SB_THUMBTRACK = 5
    Private Const SB_TOP = 6
    Private Const SB_VERT = 1

    Private Const SIF_RANGE = &H1
    Private Const SIF_PAGE = &H2
    Private Const SIF_POS = &H4
    Private Const SIF_DISABLENOSCROLL = &H8
    Private Const SIF_TRACKPOS = &H10
    Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

   Private Const ESB_DISABLE_BOTH = &H3
   Private Const ESB_ENABLE_BOTH = &H0
   
Private Declare Function EnableScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long


Private Const WM_VSCROLL = &H115
Private Const WM_HSCROLL = &H114

' Flat scroll bars:
Private Const WSB_PROP_CYVSCROLL = &H1&
Private Const WSB_PROP_CXHSCROLL = &H2&
Private Const WSB_PROP_CYHSCROLL = &H4&
Private Const WSB_PROP_CXVSCROLL = &H8&
Private Const WSB_PROP_CXHTHUMB = &H10&
Private Const WSB_PROP_CYVTHUMB = &H20&
Private Const WSB_PROP_VBKGCOLOR = &H40&
Private Const WSB_PROP_HBKGCOLOR = &H80&
Private Const WSB_PROP_VSTYLE = &H100&
Private Const WSB_PROP_HSTYLE = &H200&
Private Const WSB_PROP_WINSTYLE = &H400&
Private Const WSB_PROP_PALETTE = &H800&
Private Const WSB_PROP_MASK = &HFFF&

Private Const FSB_FLAT_MODE = 2&
Private Const FSB_ENCARTA_MODE = 1&
Private Const FSB_REGULAR_MODE = 0&

Private Declare Function FlatSB_EnableScrollBar Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal int2 As Long, ByVal UINT3 As Long) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long

Private Declare Function FlatSB_GetScrollRange Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, ByVal LPINT1 As Long, ByVal LPINT2 As Long) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function FlatSB_GetScrollPos Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long) As Long
Private Declare Function FlatSB_GetScrollProp Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal propIndex As Long, ByVal LPINT As Long) As Long

Private Declare Function FlatSB_SetScrollPos Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, ByVal pos As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollInfo Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollRange Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal code As Long, ByVal Min As Long, ByVal Max As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "COMCTL32.DLL" (ByVal hwnd As Long, ByVal index As Long, ByVal newValue As Long, ByVal fRedraw As Boolean) As Long

Private Declare Function InitializeFlatSB Lib "COMCTL32.DLL" (ByVal hwnd As Long) As Long
Private Declare Function UninitializeFlatSB Lib "COMCTL32.DLL" (ByVal hwnd As Long) As Long


' Message response:
Implements ISubclass
Private m_emr As EMsgResponse

' Initialisation state:
Private m_bInitialised As Boolean

' Orientation
Public Enum EFSOrientationConstants
    efsoHorizontal
    efsoVertical
    efsoBoth
End Enum
Private m_eOrientation As EFSOrientationConstants

' Style
Public Enum EFSStyleConstants
    efsRegular = FSB_REGULAR_MODE
    efsEncarta = FSB_ENCARTA_MODE
    efsFlat = FSB_FLAT_MODE
End Enum
Private m_eStyle As EFSStyleConstants
' Bars:
Public Enum EFSScrollBarConstants
   efsHorizontal = SB_HORZ
   efsVertical = SB_VERT
End Enum

' Can we have flat scroll bars?
Private m_bNoFlatScrollBars As Boolean

' hWnd we're adding scroll bars too:
Private m_hWnd As Long

' Small change amount
Private m_lSmallChangeHorz As Long
Private m_lSmallChangeVert As Long
' Enabled:
Private m_bEnabledHorz As Boolean
Private m_bEnabledVert As Boolean
' Visible
Private m_bVisibleHorz As Boolean
Private m_bVisibleVert As Boolean

Public Event Scroll(eBar As EFSScrollBarConstants)
Public Event Change(eBar As EFSScrollBarConstants)

Public Property Get Visible(ByVal eBar As EFSScrollBarConstants) As Boolean
   If (eBar = efsHorizontal) Then
      Visible = m_bVisibleHorz
   Else
      Visible = m_bVisibleVert
   End If
End Property
Public Property Let Visible(ByVal eBar As EFSScrollBarConstants, ByVal bState As Boolean)
   If (eBar = efsHorizontal) Then
      m_bVisibleHorz = bState
   Else
      m_bVisibleVert = bState
   End If
   If (m_bNoFlatScrollBars) Then
      ShowScrollBar m_hWnd, eBar, Abs(bState)
   Else
      FlatSB_ShowScrollBar m_hWnd, eBar, Abs(bState)
   End If
End Property

Public Property Get Orientation() As EFSOrientationConstants
   Orientation = m_eOrientation
End Property

Public Property Let Orientation(ByVal eOrientation As EFSOrientationConstants)
   m_eOrientation = eOrientation
   pSetOrientation
End Property

Private Sub pSetOrientation()
   ShowScrollBar m_hWnd, SB_HORZ, Abs((m_eOrientation = efsoBoth) Or (m_eOrientation = efsoHorizontal))
   ShowScrollBar m_hWnd, SB_VERT, Abs((m_eOrientation = efsoBoth) Or (m_eOrientation = efsoVertical))
End Sub

Private Sub pGetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
Dim lO As Long
    
    lO = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)
    
    If (m_bNoFlatScrollBars) Then
        GetScrollInfo m_hWnd, lO, tSI
    Else
        FlatSB_GetScrollInfo m_hWnd, lO, tSI
    End If

End Sub
Private Sub pLetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
Dim lO As Long
        
    lO = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)
    
    If (m_bNoFlatScrollBars) Then
        SetScrollInfo m_hWnd, lO, tSI, True
    Else
        FlatSB_SetScrollInfo m_hWnd, lO, tSI, True
    End If
    
End Sub

Property Get Style() As EFSStyleConstants
   Style = m_eStyle
End Property
Property Let Style(ByVal eStyle As EFSStyleConstants)
Dim lR As Long
   If (eStyle <> efsRegular) Then
      If (m_bNoFlatScrollBars) Then
         ' can't do it..
         'Debug.Print "Can't set non-regular style mode on this system - COMCTL32.DLL version < 4.71."
         Exit Property
      End If
   End If
   If (m_eOrientation = efsoHorizontal) Or (m_eOrientation = efsoBoth) Then
      lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_HSTYLE, eStyle, True)
   End If
   If (m_eOrientation = efsoVertical) Or (m_eOrientation = efsoBoth) Then
      lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_VSTYLE, eStyle, True)
   End If
   
   'Debug.Print lR
   m_eStyle = eStyle
End Property

Public Property Get SmallChange(ByVal eBar As EFSScrollBarConstants) As Long
   If (eBar = efsHorizontal) Then
      SmallChange = m_lSmallChangeHorz
   Else
      SmallChange = m_lSmallChangeVert
   End If
End Property
Property Let SmallChange(ByVal eBar As EFSScrollBarConstants, ByVal lSmallChange As Long)
   If (eBar = efsHorizontal) Then
      m_lSmallChangeHorz = lSmallChange
   Else
      m_lSmallChangeVert = lSmallChange
   End If
End Property
Property Get Enabled(ByVal eBar As EFSScrollBarConstants) As Boolean
   If (eBar = efsHorizontal) Then
      Enabled = m_bEnabledHorz
   Else
      Enabled = m_bEnabledVert
   End If
End Property
Property Let Enabled(ByVal eBar As EFSScrollBarConstants, ByVal bEnabled As Boolean)
Dim lO As Long
Dim lF As Long
        
   lO = eBar
   If (bEnabled) Then
      lF = ESB_ENABLE_BOTH
   Else
      lF = ESB_DISABLE_BOTH
   End If
   If (m_bNoFlatScrollBars) Then
      EnableScrollBar m_hWnd, lO, lF
   Else
      FlatSB_EnableScrollBar m_hWnd, lO, lF
   End If
    
End Property
Property Get Min(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI As SCROLLINFO
    pGetSI eBar, tSI, SIF_RANGE
    Min = tSI.nMin
End Property
Property Get Max(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI As SCROLLINFO
    pGetSI eBar, tSI, SIF_RANGE Or SIF_PAGE
    Max = tSI.nMax - tSI.nPage
End Property
Property Get Value(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI As SCROLLINFO
    pGetSI eBar, tSI, SIF_POS
    Value = tSI.nPos
End Property
Property Get LargeChange(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI As SCROLLINFO
    pGetSI eBar, tSI, SIF_PAGE
    LargeChange = tSI.nPage
End Property
Property Let Min(ByVal eBar As EFSScrollBarConstants, ByVal iMin As Long)
Dim tSI As SCROLLINFO
    tSI.nMin = iMin
    tSI.nMax = Max(eBar) + LargeChange(eBar)
    pLetSI eBar, tSI, SIF_RANGE
End Property
Property Let Max(ByVal eBar As EFSScrollBarConstants, ByVal iMax As Long)
Dim tSI As SCROLLINFO
    tSI.nMax = iMax + LargeChange(eBar)
    tSI.nMin = Min(eBar)
    pLetSI eBar, tSI, SIF_RANGE
    pRaiseEvent eBar, False
End Property
Property Let Value(ByVal eBar As EFSScrollBarConstants, ByVal iValue As Long)
Dim tSI As SCROLLINFO
    If (iValue <> Value(eBar)) Then
        tSI.nPos = iValue
        pLetSI eBar, tSI, SIF_POS
        pRaiseEvent eBar, False
    End If
End Property
Property Let LargeChange(ByVal eBar As EFSScrollBarConstants, ByVal iLargeChange As Long)
Dim tSI As SCROLLINFO
Dim lCurMax As Long
Dim lCurLargeChange As Long
    
   pGetSI eBar, tSI, SIF_ALL
   tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
   tSI.nPage = iLargeChange
   pLetSI eBar, tSI, SIF_PAGE Or SIF_RANGE
End Property
Property Get CanBeFlat() As Boolean
   CanBeFlat = Not (m_bNoFlatScrollBars)
End Property
Private Sub pCreateScrollBar()
Dim lR As Long
Dim lStyle As Long
Dim hParent As Long

   ' Redraw the window so it will take (pinched from MFC):
   hParent = GetParent(m_hWnd)
   SetParent m_hWnd, hParent
   ' Redraw for added insurance ('Slimy Windows Hacks')
   SetWindowPos m_hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
           (SWP_NOSIZE Or SWP_NOZORDER Or _
           SWP_NOMOVE Or SWP_DRAWFRAME)
   
   On Error Resume Next
   lR = InitialiseFlatSB(m_hWnd)
   If (Err.Number <> 0) Then
       'Can't find DLL entry point InitializeFlatSB in COMCTL32.DLL
       ' Means we have version prior to 4.71
       ' We get standard scroll bars.
       m_bNoFlatScrollBars = True
   Else
      Style = m_eStyle
   End If
   
End Sub

Public Sub Create(ByVal hWndA As Long)
   pClearUp
   m_hWnd = hWndA
   pCreateScrollBar
   pAttachMessages
End Sub

Private Sub pClearUp()
   If m_hWnd <> 0 Then
      On Error Resume Next
      ' Stop flat scroll bar if we have it:
      If Not (m_bNoFlatScrollBars) Then
         UninitializeFlatSB m_hWnd
      End If
    
      On Error GoTo 0
      ' Remove subclass:
      DetachMessage Me, m_hWnd, WM_HSCROLL
      DetachMessage Me, m_hWnd, WM_VSCROLL
   End If
   m_hWnd = 0
   m_bInitialised = False
End Sub
Private Sub pAttachMessages()
   If (m_hWnd <> 0) Then
      AttachMessage Me, m_hWnd, WM_HSCROLL
      AttachMessage Me, m_hWnd, WM_VSCROLL
      m_bInitialised = True
   End If
End Sub

Private Sub Class_Initialize()
   m_lSmallChangeHorz = 1
   m_lSmallChangeVert = 1
   m_eStyle = efsRegular
   m_eOrientation = efsoBoth
End Sub

Private Sub Class_Terminate()
   pClearUp
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse) ' SSubTimer6.EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse 'SSubTimer6.EMsgResponse
   ISubclass_MsgResponse = emrPostProcess
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lScrollCode As Long
Dim tSI As SCROLLINFO
Dim lV As Long, lSC As Long
Dim eBar As EFSScrollBarConstants

   Select Case iMsg
   Case WM_VSCROLL, WM_HSCROLL
      If (iMsg = WM_HSCROLL) Then
         eBar = efsHorizontal
      Else
         eBar = efsVertical
      End If
      lScrollCode = (wParam And &HFFFF&)
      Select Case lScrollCode
      Case SB_THUMBTRACK
         ' Is vertical/horizontal?
         pGetSI eBar, tSI, SIF_TRACKPOS
         Value(eBar) = tSI.nTrackPos
         pRaiseEvent eBar, True
         
      Case SB_LEFT, SB_BOTTOM
         Value(eBar) = Min(eBar)
         pRaiseEvent eBar, False
         
      Case SB_RIGHT, SB_TOP
         Value(eBar) = Max(eBar)
         pRaiseEvent eBar, False
          
      Case SB_LINELEFT, SB_LINEUP
         'Debug.Print "Line"
         lV = Value(eBar)
         If (eBar = efsHorizontal) Then
            lSC = m_lSmallChangeHorz
         Else
            lSC = m_lSmallChangeVert
         End If
         If (lV - lSC < Min(eBar)) Then
            Value(eBar) = Min(eBar)
         Else
            Value(eBar) = lV - lSC
         End If
         pRaiseEvent eBar, False
         
      Case SB_LINERIGHT, SB_LINEDOWN
          'Debug.Print "Line"
         lV = Value(eBar)
         If (eBar = efsHorizontal) Then
            lSC = m_lSmallChangeHorz
         Else
            lSC = m_lSmallChangeVert
         End If
         If (lV + lSC > Max(eBar)) Then
            Value(eBar) = Max(eBar)
         Else
            Value(eBar) = lV + lSC
         End If
         pRaiseEvent eBar, False
          
      Case SB_PAGELEFT, SB_PAGEUP
         Value(eBar) = Value(eBar) - LargeChange(eBar)
         pRaiseEvent eBar, False
         
      Case SB_PAGERIGHT, SB_PAGEDOWN
         Value(eBar) = Value(eBar) + LargeChange(eBar)
         pRaiseEvent eBar, False
         
      Case SB_ENDSCROLL
         pRaiseEvent eBar, False
         
      End Select
         
   End Select

End Function

Private Function pRaiseEvent(ByVal eBar As EFSScrollBarConstants, ByVal bScroll As Boolean)
Static s_lLastValue(0 To 1) As Long
   If (Value(eBar) <> s_lLastValue(eBar)) Then
      If (bScroll) Then
         RaiseEvent Scroll(eBar)
      Else
         RaiseEvent Change(eBar)
      End If
      s_lLastValue(eBar) = Value(eBar)
   End If
   
End Function
