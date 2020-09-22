VERSION 5.00
Begin VB.UserControl SubForm 
   Alignable       =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
   ControlContainer=   -1  'True
   ScaleHeight     =   2820
   ScaleWidth      =   3285
   ToolboxBitmap   =   "SubForm.ctx":0000
End
Attribute VB_Name = "SubForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ===========================================================================
'Copyright © 2001, Dan Mircea Fratila ( DanFratila@yahoo.com). All Rights Reserved.
'Author:Dan Mircea Fratila
' ===========================================================================
' FREE SOURCE CODE! - ENJOY.
' - Please report bugs to the author for incorporation into future releases
' - Don't sell this code.
' ===========================================================================

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000 '  WS_BORDER Or WS_DLGFRAME
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WM_SETTEXT = &HC

Private Const WM_MOVE = &H3
Private Const WM_SIZE = &H5
Private Const WM_SIZING = &H214

Private Type RECT
  left As Long
  top As Long
  right As Long
  bottom As Long
End Type

Private Type SIZE
  cX As Long
  cY As Long
End Type

Private Enum ESetWindowPosStyles 'Slimy Windows Hacks'
  SWP_SHOWWINDOW = &H40
  SWP_HIDEWINDOW = &H80
  SWP_FRAMECHANGED = &H20 ' The frame changed: send WM_NCCALCSIZE
  SWP_NOACTIVATE = &H10
  SWP_NOCOPYBITS = &H100
  SWP_NOMOVE = &H2
  SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
  SWP_NOREDRAW = &H8
  SWP_NOREPOSITION = SWP_NOOWNERZORDER
  SWP_NOSIZE = &H1
  SWP_NOZORDER = &H4
  SWP_DRAWFRAME = SWP_FRAMECHANGED
  HWND_NOTOPMOST = -2
End Enum

Public Enum ScrollStyleConstants
  efsRegular = 0& 'FSB_REGULAR_MODE
  efsEncarta = 1& 'FSB_ENCARTA_MODE
  efsFlat = 2& 'FSB_FLAT_MODE
End Enum

Public Enum sfAttachedWindowStyle
  sfOriginalStyle = 0
  sfSubformStyle = 1
  sfCustomStyle = 2
End Enum

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



Private WithEvents m_subForm As Form
Attribute m_subForm.VB_VarHelpID = -1
Private m_isOpen As Boolean


Private mSubFormSize As SIZE
Private mOffset As SIZE
Private m_IsScrolling As Boolean
Private m_OriginalAttachedWindowStyle As Long

Implements ISubclass
' ===========================================================================
'ISubclass
'Copyright © 1998-1999, Steve McMahon ( steve@vbaccelerator.com). All Rights Reserved.
'Author: Steve McMahon (steve@vbaccelerator.com)
'www.vbaccelerator.com
' ===========================================================================

Const m_def_Caption = ""
Const m_def_Moveable = False
Const m_def_Sizeable = False

Const m_def_ScrollBarStyle = 0
'Property Variables:
Dim m_Caption As String
Dim m_Moveable As Boolean
Dim m_Sizeable As Boolean

Private WithEvents m_cScroll As cScrollBars
Attribute m_cScroll.VB_VarHelpID = -1
Dim m_ScrollBarStyle As ScrollStyleConstants
' ===========================================================================
'cScrollBars
' Copyright ® 1998 Steve McMahon (steve@dogma.demon.co.uk)
' Visit vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' ===========================================================================
'Event Declarations:

Event AfterAttach(ByVal FormName As String, ByVal FormWidth As Long, ByVal FormHeight As Long)
Event Scroll(ByVal isHorizontalBar As Boolean)
Event ScrollBarChange(ByVal isHorizontalBar As Boolean)
Event GetAttachedWindowStyle(ByRef NewStyle As Long)
Event NewStyle(ByVal PropertyName As String, ByVal newValue As Boolean)



Public Sub AttachForm(frm As Object, Optional ByRef AttachedWindowNewStyle As sfAttachedWindowStyle = sfSubformStyle)
  DetachForm
  Set m_subForm = frm
  SetParent m_subForm.hwnd, UserControl.hwnd
Dim sfStyle As Long
  m_OriginalAttachedWindowStyle = GetWindowLong(m_subForm.hwnd, GWL_STYLE)
  
  Select Case AttachedWindowNewStyle
    Case sfOriginalStyle
      sfStyle = m_OriginalAttachedWindowStyle
    Case sfSubformStyle
      sfStyle = m_OriginalAttachedWindowStyle And (Not WS_CAPTION) And (Not WS_THICKFRAME)
    Case sfCustomStyle
      sfStyle = m_OriginalAttachedWindowStyle And (Not WS_CAPTION) And (Not WS_THICKFRAME)
      RaiseEvent GetAttachedWindowStyle(sfStyle)
    Case Else
      sfStyle = m_OriginalAttachedWindowStyle
  End Select
  
  SetWindowLong m_subForm.hwnd, GWL_STYLE, sfStyle
  'Slimy Windows Hacks'
  SetWindowPos m_subForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
                                (SWP_NOSIZE Or SWP_NOZORDER Or _
                                SWP_NOMOVE Or SWP_DRAWFRAME)
  
  AttachMessage Me, m_subForm.hwnd, WM_MOVE
  AttachMessage Me, m_subForm.hwnd, WM_SIZE

  AttachMessage Me, UserControl.hwnd, WM_SIZING

  m_subForm.Show
  
  Set m_cScroll = New cScrollBars
  m_cScroll.Create UserControl.hwnd
  m_cScroll.Style = m_ScrollBarStyle
  
  m_isOpen = True

Dim rcUC As RECT
  CalcSubFormSize
  GetClientRect UserControl.hwnd, rcUC
  CalcScrollSize rcUC

  SetSubFormPoz
  
Dim frmName As String, frmWith As Long, frmHeight As Long
  frmName = m_subForm.Name
  frmWith = m_subForm.Width
  frmHeight = m_subForm.Height
  
  RaiseEvent AfterAttach(frmName, frmWith, frmHeight)
  End Sub
Public Sub DetachForm()
If Not m_isOpen Then Exit Sub
  DetachMessage Me, UserControl.hwnd, WM_SIZING
  
  DetachMessage Me, m_subForm.hwnd, WM_MOVE
  DetachMessage Me, m_subForm.hwnd, WM_SIZE
  
  mOffset.cX = 0
  mOffset.cY = 0
  
  mSubFormSize.cX = 0
  mSubFormSize.cY = 0
  
  m_cScroll.Visible(efsHorizontal) = False
  m_cScroll.Visible(efsVertical) = False
  
  Set m_cScroll = Nothing
  
  m_isOpen = False
  Unload m_subForm
  
  Set m_subForm = Nothing

End Sub

Private Sub m_cScroll_Change(eBar As EFSScrollBarConstants)
Dim bIsHoriz As Boolean
  If eBar = efsHorizontal Then
    bIsHoriz = True
  Else
    bIsHoriz = False
  End If
  
   If (m_cScroll.Visible(eBar)) Then
      If (eBar = efsHorizontal) Then
         mOffset.cX = -m_cScroll.Value(eBar)
      Else
         mOffset.cY = -m_cScroll.Value(eBar)
      End If
  Else
    mOffset.cX = 0
    mOffset.cY = 0
   End If
  m_IsScrolling = True
    SetSubFormPoz
  m_IsScrolling = False
  RaiseEvent ScrollBarChange(bIsHoriz)
  
End Sub

Private Sub m_cScroll_Scroll(eBar As EFSScrollBarConstants)
  RaiseEvent Scroll(eBar)
  m_cScroll_Change eBar
End Sub

Private Sub m_subForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Cancel = m_isOpen
End Sub

Private Sub UserControl_EnterFocus()
  If m_isOpen Then m_subForm.SetFocus
End Sub


Private Sub UserControl_Resize()
If Not m_isOpen Then Exit Sub
Dim rcUC As RECT
  CalcSubFormSize
  GetClientRect UserControl.hwnd, rcUC
  CalcScrollSize rcUC

  SetSubFormPoz

End Sub

Private Sub UserControl_Terminate()
  Set m_cScroll = Nothing
  Set m_subForm = Nothing
End Sub
Private Sub CalcSubFormSize()
Dim rc As RECT
  GetWindowRect m_subForm.hwnd, rc
  With rc
    mSubFormSize.cX = .right - .left
    mSubFormSize.cY = .bottom - .top
  End With
End Sub

Private Sub SetSubFormPoz()
If Not m_isOpen Then Exit Sub
  MoveWindow m_subForm.hwnd, mOffset.cX, mOffset.cY, mSubFormSize.cX, mSubFormSize.cY, 1
End Sub
Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse) 'SSubTimer6.EMsgResponse)
'
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse 'SSubTimer6.EMsgResponse
'
End Property
Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim rcUC As RECT
  Select Case hwnd
  Case m_subForm.hwnd
    Select Case iMsg
      Case WM_MOVE
        If Not m_IsScrolling Then SetSubFormPoz
      Case WM_SIZE
      
        CalcSubFormSize
        GetClientRect UserControl.hwnd, rcUC
        CalcScrollSize rcUC
        
        SetSubFormPoz
    End Select
    
  Case UserControl.hwnd
    Select Case iMsg
      Case WM_SIZING 'WM_SIZE
      
        mOffset.cX = 0
        mOffset.cY = 0
      
        CalcSubFormSize
        
        CopyMemory rcUC, ByVal lParam, Len(rcUC)
        CalcScrollSize rcUC

        SetSubFormPoz
    End Select
  End Select
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ScrollBarStyle() As ScrollStyleConstants
Attribute ScrollBarStyle.VB_Description = "see Steve McMahon (steve@dogma.demon.co.uk) cScrollBars"
Attribute ScrollBarStyle.VB_ProcData.VB_Invoke_Property = ";Misc"
  ScrollBarStyle = m_ScrollBarStyle
End Property

Public Property Let ScrollBarStyle(ByVal New_ScrollBarStyle As ScrollStyleConstants)
  m_ScrollBarStyle = New_ScrollBarStyle
  If m_isOpen Then m_cScroll.Style = m_ScrollBarStyle

  PropertyChanged "ScrollBarStyle"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  m_ScrollBarStyle = m_def_ScrollBarStyle
  m_Moveable = m_def_Moveable
  m_Sizeable = m_def_Sizeable
  m_Caption = m_def_Caption
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_ScrollBarStyle = PropBag.ReadProperty("ScrollBarStyle", m_def_ScrollBarStyle)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000018)
  UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
  Set Picture = PropBag.ReadProperty("Picture", Nothing)
  m_Moveable = PropBag.ReadProperty("Moveable", m_def_Moveable)
  m_Sizeable = PropBag.ReadProperty("Sizeable", m_def_Sizeable)
  m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
  
  Me.Moveable = m_Moveable
  Me.Sizeable = m_Sizeable


End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("ScrollBarStyle", m_ScrollBarStyle, m_def_ScrollBarStyle)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000018)
  Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)

  Call PropBag.WriteProperty("Picture", Picture, Nothing)
  Call PropBag.WriteProperty("Moveable", m_Moveable, m_def_Moveable)
  Call PropBag.WriteProperty("Sizeable", m_Sizeable, m_def_Sizeable)
  Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property
Public Function isEmpty() As Boolean
  isEmpty = Not m_isOpen
End Function


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
  Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
  Set UserControl.Picture = New_Picture
  PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get Moveable() As Boolean
  Moveable = m_Moveable
End Property

Public Property Let Moveable(ByVal New_Moveable As Boolean)
Dim iOldWindowStyle As Long
Dim iNewWindowStyle As Long
    iOldWindowStyle = WindowStyle
    If New_Moveable Then
      iNewWindowStyle = iOldWindowStyle Or WS_CAPTION
      Me.Caption = Me.Caption
    Else
      iNewWindowStyle = iOldWindowStyle And (Not WS_CAPTION)
    End If
  WindowStyle = iNewWindowStyle
  If m_Moveable <> New_Moveable Then RaiseEvent NewStyle("Moveable", New_Moveable)
  m_Moveable = New_Moveable
  PropertyChanged "Moveable"
Dim rcUC As RECT
  If m_isOpen Then
    CalcSubFormSize
    GetClientRect UserControl.hwnd, rcUC
    CalcScrollSize rcUC

    mOffset.cX = 0
    mOffset.cY = 0
    SetSubFormPoz
  End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get Sizeable() As Boolean
  Sizeable = m_Sizeable
End Property

Public Property Let Sizeable(ByVal New_Sizeable As Boolean)
Dim iOldWindowStyle As Long
Dim iNewWindowStyle As Long
    iOldWindowStyle = WindowStyle
    If New_Sizeable Then
      iNewWindowStyle = iOldWindowStyle Or WS_THICKFRAME
    Else
      iNewWindowStyle = iOldWindowStyle And (Not WS_THICKFRAME)
    End If
  
  WindowStyle = iNewWindowStyle
  If m_Sizeable <> New_Sizeable Then RaiseEvent NewStyle("Sizeable", New_Sizeable)
  m_Sizeable = New_Sizeable
  PropertyChanged "Sizeable"
  
Dim rcUC As RECT
  If m_isOpen Then
    CalcSubFormSize
    GetClientRect UserControl.hwnd, rcUC
    CalcScrollSize rcUC
    mOffset.cX = 0
    mOffset.cY = 0
    SetSubFormPoz
  End If

End Property

Private Property Get WindowStyle() As Long
  WindowStyle = GetWindowLong(UserControl.hwnd, GWL_STYLE)
End Property

Private Property Let WindowStyle(ByVal New_WindowStyle As Long)
Dim iNewWindowStyle As Long, iOldWindowStyle As Long
  iNewWindowStyle = New_WindowStyle
  iOldWindowStyle = WindowStyle
  If (iNewWindowStyle <> iOldWindowStyle) Then
    SetWindowLong UserControl.hwnd, GWL_STYLE, iNewWindowStyle
    'Slimy Windows Hacks'
    SetWindowPos UserControl.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
                                  (SWP_NOSIZE Or SWP_NOZORDER Or _
                                  SWP_NOMOVE Or SWP_DRAWFRAME)
  End If

  
End Property
Private Sub CalcScrollSize(rcUC As RECT)
If Not m_isOpen Then Exit Sub
Dim lHeight As Long, lWidth As Long, lProportion As Long
  With rcUC
    .bottom = .bottom - .top
    .right = .right - .left
    .top = 0
    .left = 0
  End With
  
  lHeight = mSubFormSize.cY - rcUC.bottom + 15
   
   If (lHeight > 0) Then
      lProportion = lHeight / mSubFormSize.cY + 1
      m_cScroll.LargeChange(efsVertical) = lHeight \ lProportion
      m_cScroll.Max(efsVertical) = lHeight
      m_cScroll.Visible(efsVertical) = True
   Else
      m_cScroll.Visible(efsVertical) = False
   End If
   
  lWidth = mSubFormSize.cX - rcUC.right + 15
 
   If (lWidth > 0) Then
      lProportion = lWidth \ mSubFormSize.cX + 1
      m_cScroll.LargeChange(efsHorizontal) = lWidth \ lProportion
      m_cScroll.Max(efsHorizontal) = lWidth
      m_cScroll.Visible(efsHorizontal) = True
   Else
      m_cScroll.Visible(efsHorizontal) = False
   End If
End Sub



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,""
Public Property Get Caption() As String
  Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
Dim oldCaption As String
  oldCaption = m_Caption
  m_Caption = New_Caption
  SendMessage UserControl.hwnd, WM_SETTEXT, 0, ByVal m_Caption
  If oldCaption <> m_Caption Then
    PropertyChanged "Caption"
  End If
End Property

