VERSION 5.00
Object = "{1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5}#2.6#0"; "ExTVwU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ExplorerTreeView 2.6 - CustomDraw sample"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2108
      TabIndex        =   1
      Top             =   6120
      Width           =   1335
   End
   Begin ExTVwLibUCtl.ExplorerTreeView ExTvw 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _cx             =   9340
      _cy             =   10398
      AllowDragDrop   =   0   'False
      AllowLabelEditing=   0   'False
      AlwaysShowSelection=   -1  'True
      Appearance      =   1
      AutoHScroll     =   -1  'True
      AutoHScrollPixelsPerSecond=   150
      AutoHScrollRedrawInterval=   15
      BackColor       =   -2147483643
      BlendSelectedItemsIcons=   0   'False
      BorderStyle     =   0
      BuiltInStateImages=   0
      CaretChangedDelayTime=   500
      DisabledEvents  =   263295
      DontRedraw      =   0   'False
      DragExpandTime  =   -1
      DragScrollTimeBase=   -1
      DrawImagesAsynchronously=   0   'False
      EditBackColor   =   -2147483643
      EditForeColor   =   -2147483640
      EditHoverTime   =   -1
      EditIMEMode     =   -1
      Enabled         =   -1  'True
      FadeExpandos    =   0   'False
      FavoritesStyle  =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FullRowSelect   =   0   'False
      GroupBoxColor   =   -2147483632
      HotTracking     =   0
      HoverTime       =   -1
      IMEMode         =   -1
      Indent          =   19
      IndentStateImages=   -1  'True
      InsertMarkColor =   0
      ItemBoundingBoxDefinition=   70
      ItemHeight      =   17
      ItemXBorder     =   3
      ItemYBorder     =   0
      LineColor       =   -2147483632
      LineStyle       =   1
      MaxScrollTime   =   100
      MousePointer    =   0
      OLEDragImageStyle=   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      ScrollBars      =   2
      ShowStateImages =   0   'False
      ShowToolTips    =   -1  'True
      SingleExpand    =   0
      SortOrder       =   0
      SupportOLEDragImages=   -1  'True
      TreeViewStyle   =   3
      UseSystemFont   =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Const ANSI_CHARSET = 0
  Private Const BALTIC_CHARSET = 186
  Private Const CHINESEBIG5_CHARSET = 136
  Private Const DEFAULT_CHARSET = 1
  Private Const EASTEUROPE_CHARSET = 238
  Private Const GB2312_CHARSET = 134
  Private Const GREEK_CHARSET = 161
  Private Const HANGUL_CHARSET = 129
  Private Const MAC_CHARSET = 77
  Private Const OEM_CHARSET = 255
  Private Const RUSSIAN_CHARSET = 204
  Private Const SHIFTJIS_CHARSET = 128
  Private Const SYMBOL_CHARSET = 2
  Private Const TURKISH_CHARSET = 162


  Private Type ItmData
    ItmType As Byte
    Tag As Long
  End Type

  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type


  Private hFontBold As Long
  Private themeableOS As Boolean
  Private ToolTip As clsToolTip


  Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT1) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectW" (ByVal hgdiobj As Long, ByVal cbBuffer As Long, lpvObject As Any) As Long
  Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal i As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMain
      lRet = HandleMessage_Form(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "frmMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Form(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_NOTIFYFORMAT = &H55
  Const WM_USER = &H400
  Const OCM__BASE = WM_USER + &H1C00
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_NOTIFYFORMAT
      ' give the control a chance to request Unicode notifications
      lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)

      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub cmdAbout_Click()
  ExTvw.About
End Sub

Private Sub ExTvw_CustomDraw(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, itemLevel As Long, TextColor As stdole.OLE_COLOR, textBackColor As stdole.OLE_COLOR, ByVal drawStage As ExTVwLibUCtl.CustomDrawStageConstants, ByVal itemState As ExTVwLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExTVwLibUCtl.RECTANGLE, furtherProcessing As ExTVwLibUCtl.CustomDrawReturnValuesConstants)
  Const COLOR_HIGHLIGHT = 13
  Const COLOR_WINDOW = 5
  Const DT_LEFT = &H0
  Const DT_VCENTER = &H4
  Const OBJ_FONT = 6
  Dim hFont As Long
  Dim hOldFont As Long
  Dim nOldClr As Long
  Dim rc As RECT

  Select Case drawStage
    Case CustomDrawStageConstants.cdsPrePaint
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvNotifyPostPaint Or CustomDrawReturnValuesConstants.cdrvNotifyItemDraw

    Case CustomDrawStageConstants.cdsItemPrePaint
      If (treeItem.ItemData >= -20) And (treeItem.ItemData <= -1) Then
        If itemState And CustomDrawItemStateConstants.cdisFocus Then
        ' use the following line instead of the previous line, if you've set AllowDragDrop to True
        'If textBackColor = GetSysColor(COLOR_HIGHLIGHT) Then
          SelectObject hDC, hFontBold
        Else
          TextColor = CLng("&H" & Mid(treeItem.Text, 3))
        End If
      Else
        hFont = treeItem.ItemData
        If GetObjectType(hFont) = OBJ_FONT Then SelectObject hDC, hFont
      End If
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvNewFont Or CustomDrawReturnValuesConstants.cdrvNotifyPostPaint

    Case CustomDrawStageConstants.cdsItemPostPaint
      ' GetRectangle will fail if the item is not yet visible to the user
      On Error GoTo ItemNotYetVisible
      treeItem.GetRectangle irtItem, rc.Left, rc.Top, rc.Right, rc.Bottom
      rc.Left = rc.Right + 5
      rc.Right = ExTvw.Width
      If (treeItem.ItemData >= -20) And (treeItem.ItemData <= -1) Then
        If itemState And CustomDrawItemStateConstants.cdisFocus Then
          nOldClr = SetTextColor(hDC, RGB(0, 0, 255))
          Select Case treeItem.ItemData
            Case -1
              DrawText hDC, StrPtr("White"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -2
              DrawText hDC, StrPtr("Red"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -3
              DrawText hDC, StrPtr("Green"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -4
              DrawText hDC, StrPtr("Blue"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -5
              DrawText hDC, StrPtr("Magenta"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -6
              DrawText hDC, StrPtr("Cyan"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -7
              DrawText hDC, StrPtr("Yellow"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -8
              DrawText hDC, StrPtr("Black"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -9
              DrawText hDC, StrPtr("Aquamarine"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -10
              DrawText hDC, StrPtr("Baker's Chocolate"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -11
              DrawText hDC, StrPtr("Blue Violet"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -12
              DrawText hDC, StrPtr("Brass"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -13
              DrawText hDC, StrPtr("Bright Gold"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -14
              DrawText hDC, StrPtr("Brown"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -15
              DrawText hDC, StrPtr("Bronze"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -16
              DrawText hDC, StrPtr("Bronze II"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -17
              DrawText hDC, StrPtr("Cadet Blue"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -18
              DrawText hDC, StrPtr("Cool Copper"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -19
              DrawText hDC, StrPtr("Copper"), -1, rc, DT_LEFT Or DT_VCENTER
            Case -20
              DrawText hDC, StrPtr("Coral"), -1, rc, DT_LEFT Or DT_VCENTER
          End Select
          SetTextColor hDC, nOldClr
        Else
          FillRect hDC, rc, COLOR_WINDOW + 1
        End If
      End If

ItemNotYetVisible:
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvDoDefault
  End Select
End Sub

Private Sub ExTvw_DestroyedControlWindow(ByVal hWndTreeView As Long)
  ToolTip.Detach
End Sub

Private Sub ExTvw_FreeItemData(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem)
  Const OBJ_FONT = 6
  Dim h As Long

  h = treeItem.ItemData
  If GetObjectType(h) = OBJ_FONT Then DeleteObject h
End Sub

Private Sub ExTvw_ItemGetInfoTipText(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  Const OBJ_FONT = 6

  If GetObjectType(treeItem.ItemData) = OBJ_FONT Then
    infoTipText = treeItem.Text
  End If
End Sub

Private Sub ExTvw_RecreatedControlWindow(ByVal hWndTreeView As Long)
  InsertItems
End Sub

Private Sub Form_Initialize()
  Dim hMod As Long

  InitCommonControls

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If

  Set ToolTip = New clsToolTip
End Sub

Private Sub Form_Load()
  Const FW_BOLD = 700
  Const WM_GETFONT = &H31
  Dim hFont As Long
  Dim lf As LOGFONT1

  ' this is required to make the control work as expected
  Subclass

  InsertItems

  hFont = SendMessage(ExTvw.hWnd, WM_GETFONT, 0, 0)
  GetObjectAPI hFont, LenB(lf), lf
  lf.lfWeight = lf.lfWeight Or FW_BOLD
  hFontBold = CreateFontIndirect(lf)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Const OBJ_FONT = 6
  Dim h As Long
  Dim treeItem As TreeViewItem

  ' The FreeItemData won't get fired on program termination (actually it gets fired, but we won't
  ' receive it anymore). So ensure the fonts get freed.
  For Each treeItem In ExTvw.TreeItems
    h = treeItem.ItemData
    If GetObjectType(h) = OBJ_FONT Then
      DeleteObject h
    End If
  Next treeItem
  ToolTip.Detach
  Set ToolTip = Nothing

  DeleteObject hFontBold

  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub


Private Sub InsertItems()
  Const WM_GETFONT = &H31
  Dim hDC As Long
  Dim hFont As Long
  Dim itm1 As ExTVwLibUCtl.TreeViewItem
  Dim itm2 As ExTVwLibUCtl.TreeViewItem

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExTvw.hWnd, StrPtr("explorer"), 0
  End If

  ToolTip.Attach ExTvw.hWndToolTip
  ToolTip.BalloonStyle = True
  ToolTip.TitleText = "Font name:"
  ToolTip.TitleIcon = ToolTipTitleIconConstants.tttiInfo

  ExTvw.MousePointer = ExTVwLibUCtl.MousePointerConstants.mpHourglass
  With ExTvw.TreeItems
    Set itm1 = .Add("Some Colors and their (HTML) names", Nothing, , heYes, , , , 3)

    .Add "0xFFFFFF   ", itm1, , , , , , -1
    .Add "0x0000FF   ", itm1, , , , , , -2
    .Add "0x00FF00   ", itm1, , , , , , -3
    .Add "0xFF0000   ", itm1, , , , , , -4
    .Add "0xFF00FF   ", itm1, , , , , , -5
    .Add "0xFFFF00   ", itm1, , , , , , -6
    .Add "0x00FFFF   ", itm1, , , , , , -7
    .Add "0x000000   ", itm1, , , , , , -8
    .Add "0x93DB70   ", itm1, , , , , , -9
    .Add "0x17335C   ", itm1, , , , , , -10
    .Add "0x9F5F9F   ", itm1, , , , , , -11
    .Add "0x42A6B5   ", itm1, , , , , , -12
    .Add "0x19D9D9   ", itm1, , , , , , -13
    .Add "0x2A2AA6   ", itm1, , , , , , -14
    .Add "0x53788C   ", itm1, , , , , , -15
    .Add "0x3D7DA6   ", itm1, , , , , , -16
    .Add "0x9F9F5F   ", itm1, , , , , , -17
    .Add "0x1987D9   ", itm1, , , , , , -18
    .Add "0x3373B8   ", itm1, , , , , , -19
    .Add "0x007FFF   ", itm1, , , , , , -20

    hFont = SendMessage(ExTvw.hWnd, WM_GETFONT, 0, 0)
    GetObjectAPI hFont, LenB(lfTreeView), lfTreeView
    lfTreeView.lfFaceName(1) = 0
    lfTreeView.lfFaceName(2) = 0
    hDC = GetDC(0)

    Set itm1 = .Add("Fonts installed in the system", Nothing, , heYes, , , , 3)
    Set itm2 = .Add("ANSI_CHARSET", itm1, , heYes, , , , -100)
    EnumFonts hDC, ExTvw, itm2, ANSI_CHARSET
    Set itm2 = .Add("BALTIC_CHARSET", itm1, , heYes, , , , -101)
    EnumFonts hDC, ExTvw, itm2, BALTIC_CHARSET
    Set itm2 = .Add("CHINESEBIG5_CHARSET", itm1, , heYes, , , , -102)
    EnumFonts hDC, ExTvw, itm2, CHINESEBIG5_CHARSET
    Set itm2 = .Add("DEFAULT_CHARSET", itm1, , heYes, , , , -103)
    EnumFonts hDC, ExTvw, itm2, DEFAULT_CHARSET
    Set itm2 = .Add("EASTEUROPE_CHARSET", itm1, , heYes, , , , -104)
    EnumFonts hDC, ExTvw, itm2, EASTEUROPE_CHARSET
    Set itm2 = .Add("GB2312_CHARSET", itm1, , heYes, , , , -105)
    EnumFonts hDC, ExTvw, itm2, GB2312_CHARSET
    Set itm2 = .Add("GREEK_CHARSET", itm1, , heYes, , , , -106)
    EnumFonts hDC, ExTvw, itm2, GREEK_CHARSET
    Set itm2 = .Add("HANGUL_CHARSET", itm1, , heYes, , , , -107)
    EnumFonts hDC, ExTvw, itm2, HANGUL_CHARSET
    Set itm2 = .Add("MAC_CHARSET", itm1, , heYes, , , , -108)
    EnumFonts hDC, ExTvw, itm2, MAC_CHARSET
    Set itm2 = .Add("OEM_CHARSET", itm1, , heYes, , , , -109)
    EnumFonts hDC, ExTvw, itm2, OEM_CHARSET
    Set itm2 = .Add("RUSSIAN_CHARSET", itm1, , heYes, , , , -110)
    EnumFonts hDC, ExTvw, itm2, RUSSIAN_CHARSET
    Set itm2 = .Add("SHIFTJIS_CHARSET", itm1, , heYes, , , , -111)
    EnumFonts hDC, ExTvw, itm2, SHIFTJIS_CHARSET
    Set itm2 = .Add("SYMBOL_CHARSET", itm1, , heYes, , , , -112)
    EnumFonts hDC, ExTvw, itm2, SYMBOL_CHARSET
    Set itm2 = .Add("TURKISH_CHARSET", itm1, , heYes, , , , -113)
    EnumFonts hDC, ExTvw, itm2, TURKISH_CHARSET

    itm1.SortSubItems , , , , True

    ReleaseDC 0, hDC
  End With
  ExTvw.MousePointer = ExTVwLibUCtl.MousePointerConstants.mpDefault
End Sub

' subclasses this Form
Private Sub Subclass()
  Const NF_REQUERY = 4
  Const WM_NOTIFYFORMAT = &H55

  #If UseSubClassing Then
    If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
      Debug.Print "Subclassing failed!"
    End If
    ' tell the control to negotiate the correct format with the form
    SendMessageAsLong ExTvw.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
