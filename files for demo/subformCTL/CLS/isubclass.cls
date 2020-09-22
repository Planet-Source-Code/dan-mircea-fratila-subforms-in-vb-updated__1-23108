VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ===========================================================================
' Copyright ® 1998 Steve McMahon (steve@dogma.demon.co.uk)
' Visit vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' ---------------------------------------------------------------------------
' FREE SOURCE CODE! - ENJOY.
' - Please report bugs to the author for incorporation into future releases
' - Don't sell this code.
' ===========================================================================
'
'Note:
'You are free to use and modify any of the code on the site,
'but not to distribute modified versions of my downloads,
'ActiveX DLLs or OCXs with the same filename and/or ProgIds.
'If you have made changes which you think are beneficial,
'or have bug reports, then you can email me
'(steve@vbaccelerator.com) and I will do my utmost to get the a new version
'on the site.
'You can freely distribute the zips available from this site to other ones,
'but you must distribute them in their original state and particularly keep
'the information and disclaimer text file with the zip (if it has one!).
'Notification would be greatly appreciated!
'You can freely distribute any compiled code on the site,
'or any products you build using the code.
'If you wish to distribute the source code files by any other means
'(i.e. if you want to include it on a CD or any other software media)
'then the EXPRESS PERMISSION of the author is REQUIRED
' ===========================================================================

Public Enum EMsgResponse
    emrConsume      ' Process instead of original WindowProc
    emrPostProcess  ' Process after original WindowProc
    emrPreprocess   ' Process before original WindowProc
End Enum

Public MsgResponse As EMsgResponse
Attribute MsgResponse.VB_VarDescription = "Get/set how you will respond to a message.  Raised whenever a message is encountered.  Call the CurrentMessage property to determine what message you are currently responding to."

Function WindowProc(ByVal hwnd As Long, _
                    ByVal iMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long) As Long
Attribute WindowProc.VB_Description = "Raised whenever a message you have subclassed is sent."
End Function
'

