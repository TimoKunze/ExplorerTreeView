VERSION 5.00
Object = "{1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5}#2.6#0"; "ExTVwU.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "ExplorerTreeView 2.6 - Multi Column Sample"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
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
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picContainer 
      BorderStyle     =   0  'Kein
      Height          =   5700
      Left            =   0
      ScaleHeight     =   380
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4455
      Begin ExTVwLibUCtl.ExplorerTreeView ExTvwU 
         Height          =   5700
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4455
         _cx             =   7858
         _cy             =   10054
         AllowDragDrop   =   0   'False
         AllowLabelEditing=   0   'False
         AlwaysShowSelection=   -1  'True
         Appearance      =   0
         AutoHScroll     =   -1  'True
         AutoHScrollPixelsPerSecond=   150
         AutoHScrollRedrawInterval=   15
         BackColor       =   -2147483643
         BlendSelectedItemsIcons=   0   'False
         BorderStyle     =   0
         BuiltInStateImages=   0
         CaretChangedDelayTime=   500
         DisabledEvents  =   263935
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
         Indent          =   14
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
         ScrollBars      =   1
         ShowStateImages =   0   'False
         ShowToolTips    =   -1  'True
         SingleExpand    =   0
         SortOrder       =   0
         SupportOLEDragImages=   -1  'True
         TreeViewStyle   =   1
         UseSystemFont   =   -1  'True
      End
   End
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
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Type HDITEM
    mask As Long
    cxy As Long
    pszText As Long
    hbm As Long
    cchTextMax As Long
    fmt As Long
    lParam As Long
    iImage As Long
    iOrder As Long
    type As Long
    pvFilter As Long
  End Type

  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type

  Private Type NMHDR
    hWndFrom As Long
    idfrom As Long
    code As Long
  End Type

  Private Type NMHEADER
    hdr As NMHDR
    iItem As Long
    iButton As Long
    pitem As Long
  End Type

  Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
  End Type


  Private propScrollBarVisible As Boolean
  Private propSmallScrollChange As Long

  Private bComctl32Version600OrNewer As Boolean
  Private hImgLst As Long
  Private hTheme As Long
  Private hWndHeader As Long
  Private themeableOS As Boolean


  Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal cb As Long)
  Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
  Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pRect As RECT, ByVal pClipRect As Long) As Long
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
  Private Declare Function GetScrollInfo Lib "user32.dll" (ByVal hWnd As Long, ByVal fnBar As Long, ByRef lpsi As SCROLLINFO) As Long
  Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function InvalidateRectAsLong Lib "user32.dll" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
  Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
  Private Declare Function SetScrollInfo Lib "user32.dll" (ByVal hWnd As Long, ByVal fnBar As Long, ByRef lpsi As SCROLLINFO, ByVal fRedraw As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
  Private Declare Function ShowScrollBar Lib "user32.dll" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMainPicContainer
      lRet = HandleMessage_PicContainer(hWnd, uMsg, wParam, lParam, bCallDefProc)
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

Private Function HandleMessage_PicContainer(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const COLOR_WINDOW = 5
  Const HDI_WIDTH = &H1
  Const HDN_FIRST = (-300)
  Const HDN_ITEMCHANGED = (HDN_FIRST - 1)
  Const HDN_ITEMCHANGING = (HDN_FIRST - 0)
  Const HDN_TRACK = (HDN_FIRST - 8)
  Const NFR_UNICODE = 2
  Const SB_ENDSCROLL = 8
  Const SB_LEFT = 6
  Const SB_LINELEFT = 0
  Const SB_LINERIGHT = 1
  Const SB_PAGELEFT = 2
  Const SB_PAGERIGHT = 3
  Const SB_RIGHT = 7
  Const SB_THUMBTRACK = 5
  Const WM_HSCROLL = &H114
  Const WM_NOTIFY = &H4E
  Const WM_NOTIFYFORMAT = &H55
  Const WM_USER = &H400
  Const OCM__BASE = WM_USER + &H1C00
  Const SB_HORZ = 0
  Const SIF_TRACKPOS = &H10
  Dim currentScrollValue As Long
  Dim hdi As HDITEM
  Dim headerNotificationDetails As NMHEADER
  Dim notificationDetails As NMHDR
  Dim scrInfo As SCROLLINFO
  Dim scrollCode As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_NOTIFYFORMAT
      If wParam = hWndHeader Then
        lRet = NFR_UNICODE
      Else
        ' this is required to make the control work as expected
        lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)
      End If
      bCallDefProc = False
    Case WM_HSCROLL
      scrollCode = (wParam And &HFFFF&)
      Select Case scrollCode
        Case SB_THUMBTRACK
          scrInfo.cbSize = LenB(scrInfo)
          scrInfo.fMask = SIF_TRACKPOS
          GetScrollInfo picContainer.hWnd, SB_HORZ, scrInfo
          ScrollValue = scrInfo.nTrackPos
          ScollBarScrolled
        Case SB_LEFT
          ScrollValue = MinScrollValue
          ScollBarChanged
        Case SB_RIGHT
          ScrollValue = MaxScrollValue
          ScollBarChanged
        Case SB_LINELEFT
          currentScrollValue = ScrollValue
          If currentScrollValue - SmallScrollChange < MinScrollValue Then
            ScrollValue = MinScrollValue
          Else
            ScrollValue = currentScrollValue - SmallScrollChange
          End If
          ScollBarChanged
        Case SB_LINERIGHT
          currentScrollValue = ScrollValue
          If currentScrollValue + SmallScrollChange > MaxScrollValue Then
            ScrollValue = MaxScrollValue
          Else
            ScrollValue = currentScrollValue + SmallScrollChange
          End If
          ScollBarChanged
        Case SB_PAGELEFT
          ScrollValue = ScrollValue - LargeScrollChange
          ScollBarChanged
        Case SB_PAGERIGHT
          ScrollValue = ScrollValue + LargeScrollChange
          ScollBarChanged
        Case SB_ENDSCROLL
          ScollBarChanged
      End Select
      bCallDefProc = False
    Case WM_NOTIFY
      CopyMemory VarPtr(notificationDetails), lParam, LenB(notificationDetails)
      If notificationDetails.hWndFrom = hWndHeader Then
        Select Case notificationDetails.code
          Case HDN_ITEMCHANGED
            CopyMemory VarPtr(headerNotificationDetails), lParam, LenB(headerNotificationDetails)
            CopyMemory VarPtr(hdi), headerNotificationDetails.pitem, LenB(hdi)
            If hdi.mask And HDI_WIDTH Then
              ExTvwU.Refresh
              SizeMultiColumnTreeView
            End If
            bCallDefProc = False
        End Select
      End If
  End Select

StdHandler_Ende:
  HandleMessage_PicContainer = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_PicContainer: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub cmdAbout_Click()
  ExTvwU.About
End Sub

Private Sub ExTvwU_CustomDraw(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, itemLevel As Long, textColor As stdole.OLE_COLOR, textBackColor As stdole.OLE_COLOR, ByVal drawStage As ExTVwLibUCtl.CustomDrawStageConstants, ByVal itemState As ExTVwLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExTVwLibUCtl.RECTANGLE, furtherProcessing As ExTVwLibUCtl.CustomDrawReturnValuesConstants)
  Const COLOR_WINDOW = 5
  Const DT_CALCRECT = &H400
  Const DT_END_ELLIPSIS = &H8000&
  Const DT_NOPREFIX = &H800
  Const HDI_WIDTH = &H1
  Const HDM_FIRST = &H1200
  Const HDM_GETITEM = (HDM_FIRST + 11)
  Const HDM_GETITEMCOUNT = (HDM_FIRST + 0)
  Const TRANSPARENT = 1
  Dim cColumns As Long
  Dim cxColumn() As Long
  Dim cxOverall As Long
  Dim hBrBack As Long
  Dim hdi As HDITEM
  Dim i As Long
  Dim rc As RECT
  Dim rcClear As RECT
  Dim rcLabel As RECT
  Dim rcText As RECT
  Dim xOffset As Long

  Select Case drawStage
    Case CustomDrawStageConstants.cdsPrePaint
      If hWndHeader Then
        furtherProcessing = CustomDrawReturnValuesConstants.cdrvNotifyItemDraw
      End If

    Case CustomDrawStageConstants.cdsItemPrePaint
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvNotifyPostPaint

    Case CustomDrawStageConstants.cdsItemPostPaint
      cColumns = SendMessageAsLong(hWndHeader, HDM_GETITEMCOUNT, 0, 0)
      ReDim cxColumn(0 To cColumns - 1) As Long
      hdi.mask = HDI_WIDTH
      For i = 0 To cColumns - 1
        SendMessageAsLong hWndHeader, HDM_GETITEM, i, VarPtr(hdi)
        cxColumn(i) = hdi.cxy
        cxOverall = cxOverall + cxColumn(i)
      Next i

      ' GetRectangle will fail if the item is not yet visible to the user
      On Error GoTo ItemNotYetVisible
      treeItem.GetRectangle irtItem, rcLabel.Left, rcLabel.Top, rcLabel.Right, rcLabel.Bottom
      LSet rcClear = drawingRectangle
      rcClear.Left = rcLabel.Left
      If cxOverall - 2 < rcClear.Right Then
        rcClear.Right = cxOverall - 2
      End If

      rc = rcClear
      rc.Right = ExTvwU.Width
      hBrBack = GetSysColorBrush(COLOR_WINDOW)
      If hBrBack Then
        FillRect hDC, rc, hBrBack
      End If
      hBrBack = CreateSolidBrush(textBackColor)
      If hBrBack Then
        FillRect hDC, rcClear, hBrBack
        DeleteObject hBrBack
      End If

      DrawText hDC, StrPtr(treeItem.Text), -1, rcText, DT_NOPREFIX Or DT_CALCRECT
      If rcLabel.Left + rcText.Right + 4 < cxColumn(0) - 4 Then
        rcLabel.Right = rcLabel.Left + rcText.Right + 4
      Else
        rcLabel.Right = cxColumn(0) - 4
      End If

      SetBkMode hDC, TRANSPARENT
      SetTextColor hDC, textColor

      ' draw main label
      rcText = rcLabel
      InflateRect rcText, -2, -1
      DrawText hDC, StrPtr(treeItem.Text), -1, rcText, DT_NOPREFIX Or DT_END_ELLIPSIS

      ' draw other columns
      xOffset = cxColumn(0)
      For i = 1 To cColumns - 1
        rcText = rcLabel
        rcText.Left = xOffset
        rcText.Right = xOffset + cxColumn(i)
        InflateRect rcText, -2, -1
        DrawText hDC, StrPtr("SubItem " & i), -1, rcText, DT_NOPREFIX Or DT_END_ELLIPSIS
        xOffset = xOffset + cxColumn(i)
      Next i

ItemNotYetVisible:
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvDoDefault
  End Select
End Sub

Private Sub ExTvwU_RecreatedControlWindow(ByVal hWnd As Long)
  InsertItems
End Sub

Private Sub Form_Initialize()
  Const ILC_COLOR24 = &H18
  Const ILC_COLOR32 = &H20
  Const ILC_MASK = &H1
  Const IMAGE_ICON = 1
  Const LR_DEFAULTSIZE = &H40
  Const LR_LOADFROMFILE = &H10
  Dim DLLVerData As DLLVERSIONINFO
  Dim hIcon As Long
  Dim hMod As Long
  Dim iconsDir As String
  Dim iconPath As String

  InitCommonControls

  propSmallScrollChange = 1

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With

  hImgLst = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 13, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
  iconPath = Dir$(iconsDir & "*.ico")
  While iconPath <> ""
    hIcon = LoadImage(0, StrPtr(iconsDir & iconPath), IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst, hIcon
      DestroyIcon hIcon
    End If
    iconPath = Dir$
  Wend
End Sub

Private Sub Form_Load()
  Const HDF_LEFT = 0
  Const HDF_STRING = &H4000
  Const HDI_FORMAT = &H4
  Const HDI_ORDER = &H80
  Const HDI_TEXT = &H2
  Const HDI_WIDTH = &H1
  Const HDS_FULLDRAG = &H80
  Const HDS_HORZ = &H0
  Const HDM_FIRST = &H1200
  Const HDM_INSERTITEM = (HDM_FIRST + 10)
  Const WS_CHILD = &H40000000
  Const WS_VISIBLE = &H10000000
  Const WM_GETFONT = &H31
  Const WM_SETFONT = &H30
  Dim bComctl32Version600OrNewer As Boolean
  Dim DLLVerData As DLLVERSIONINFO
  Dim hdi As HDITEM
  Dim hFont As Long

  If themeableOS Then
    With DLLVerData
      .cbSize = LenB(DLLVerData)
      DllGetVersion_comctl32 DLLVerData
      bComctl32Version600OrNewer = (.dwMajor >= 6)
    End With
    If bComctl32Version600OrNewer And IsAppThemed() Then
      hTheme = OpenThemeData(Me.hWnd, StrPtr("TreeView"))
    End If
  End If

  ' this is required to make the control work as expected
  Subclass

  hFont = SendMessageAsLong(ExTvwU.hWnd, WM_GETFONT, 0, 0)
  hWndHeader = CreateWindowEx(0, StrPtr("SysHeader32"), StrPtr(""), HDS_HORZ Or WS_CHILD Or WS_VISIBLE, 0, 0, 0, 0, picContainer.hWnd, 0, App.hInstance, 0)
  If hWndHeader Then
    ' tell the controls to negotiate the correct format with the form
    SendMessageAsLong hWndHeader, WM_SETFONT, hFont, 1
    With hdi
      .mask = HDI_FORMAT Or HDI_ORDER Or HDI_TEXT Or HDI_WIDTH
      .cxy = ExTvwU.Width / 3
      .fmt = HDF_LEFT Or HDF_STRING
      .cxy = .cxy + 30
      .pszText = StrPtr("Column 1")
      .cchTextMax = Len("Column 1")
      .iOrder = 0
      SendMessageAsLong hWndHeader, HDM_INSERTITEM, 0, VarPtr(hdi)
      .cxy = .cxy - 45
      .pszText = StrPtr("Column 2")
      .cchTextMax = Len("Column 2")
      .iOrder = 1
      SendMessageAsLong hWndHeader, HDM_INSERTITEM, 1, VarPtr(hdi)
      .pszText = StrPtr("Column 3")
      .cchTextMax = Len("Column 3")
      .iOrder = 2
      SendMessageAsLong hWndHeader, HDM_INSERTITEM, 2, VarPtr(hdi)
    End With
  End If

  InsertItems
End Sub

Private Sub Form_Resize()
  If hTheme Then
    picContainer.Move 1, 1, Me.ScaleWidth - 8 - cmdAbout.Width - 8 - 1, Me.ScaleHeight - 2
  Else
    picContainer.Move 2, 2, Me.ScaleWidth - 8 - cmdAbout.Width - 8 - 2, Me.ScaleHeight - 4
  End If
  cmdAbout.Left = Me.ScaleWidth - 8 - cmdAbout.Width
  SizeMultiColumnTreeView
  picContainer_Paint
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(picContainer.hWnd, EnumSubclassID.escidFrmMainPicContainer) Then
    Debug.Print "UnSubclassing failed!"
  End If

  If hTheme Then
    CloseThemeData hTheme
  End If
End Sub

Private Sub picContainer_Paint()
  Const BDR_SUNKENINNER = &H8
  Const BDR_SUNKENOUTER = &H2
  Const BF_BOTTOM = &H8
  Const BF_FLAT = &H4000
  Const BF_LEFT = &H1
  Const BF_RIGHT = &H4
  Const BF_TOP = &H2
  Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
  Const EDGE_SUNKEN = (BDR_SUNKENINNER Or BDR_SUNKENOUTER)
  Dim rc As RECT

  Me.Cls
  If hTheme Then
    rc.Left = picContainer.Left - 1
    rc.Right = picContainer.Left + picContainer.Width + 1
    rc.Top = picContainer.Top - 1
    rc.Bottom = picContainer.Top + picContainer.Height + 1
    DrawThemeBackground hTheme, Me.hDC, 0, 0, rc, 0
  Else
    rc.Left = picContainer.Left - 2
    rc.Right = picContainer.Left + picContainer.Width + 2
    rc.Top = picContainer.Top - 2
    rc.Bottom = picContainer.Top + picContainer.Height + 2
    DrawEdge Me.hDC, rc, EDGE_SUNKEN, BF_RECT
  End If
End Sub


Private Property Get LargeScrollChange() As Long
  Const SB_HORZ = 0
  Const SIF_PAGE = &H2
  Dim scrInfo As SCROLLINFO

  scrInfo.cbSize = LenB(scrInfo)
  scrInfo.fMask = SIF_PAGE
  GetScrollInfo picContainer.hWnd, SB_HORZ, scrInfo
  LargeScrollChange = scrInfo.nPage
End Property

Private Property Let LargeScrollChange(ByVal newValue As Long)
  Const SB_HORZ = 0
  Const SIF_PAGE = &H2
  Const SIF_RANGE = &H1
  Const SIF_POS = &H4
  Const SIF_TRACKPOS = &H10
  Const SIF_ALL = SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS
  Dim scrInfo As SCROLLINFO

  scrInfo.cbSize = LenB(scrInfo)
  scrInfo.fMask = SIF_ALL
  GetScrollInfo picContainer.hWnd, SB_HORZ, scrInfo
  scrInfo.nMax = scrInfo.nMax - scrInfo.nPage + newValue
  scrInfo.nPage = newValue
  scrInfo.fMask = SIF_PAGE Or SIF_RANGE
  SetScrollInfo picContainer.hWnd, SB_HORZ, scrInfo, 1
End Property

Private Property Get MaxScrollValue() As Long
  Const SB_HORZ = 0
  Const SIF_PAGE = &H2
  Const SIF_RANGE = &H1
  Dim scrInfo As SCROLLINFO

  scrInfo.cbSize = LenB(scrInfo)
  scrInfo.fMask = SIF_RANGE Or SIF_PAGE
  GetScrollInfo picContainer.hWnd, SB_HORZ, scrInfo
  MaxScrollValue = scrInfo.nMax - scrInfo.nPage
End Property

Private Property Let MaxScrollValue(ByVal newValue As Long)
  Const SB_HORZ = 0
  Const SIF_RANGE = &H1
  Dim scrInfo As SCROLLINFO

  scrInfo.nMax = newValue + LargeScrollChange
  scrInfo.nMin = MinScrollValue
  scrInfo.fMask = SIF_RANGE
  SetScrollInfo picContainer.hWnd, SB_HORZ, scrInfo, 1
  ScollBarChanged
End Property

Private Property Get MinScrollValue() As Long
  Const SB_HORZ = 0
  Const SIF_RANGE = &H1
  Dim scrInfo As SCROLLINFO

  scrInfo.cbSize = LenB(scrInfo)
  scrInfo.fMask = SIF_RANGE
  GetScrollInfo picContainer.hWnd, SB_HORZ, scrInfo
  MinScrollValue = scrInfo.nMin
End Property

Private Property Let MinScrollValue(ByVal newValue As Long)
  Const SB_HORZ = 0
  Const SIF_RANGE = &H1
  Dim scrInfo As SCROLLINFO

  scrInfo.nMax = MaxScrollValue + LargeScrollChange
  scrInfo.nMin = newValue
  scrInfo.fMask = SIF_RANGE
  SetScrollInfo picContainer.hWnd, SB_HORZ, scrInfo, 1
End Property

Private Property Get ScrollBarVisible() As Boolean
  ScrollBarVisible = propScrollBarVisible
End Property

Private Property Let ScrollBarVisible(ByVal newValue As Boolean)
  Const SB_HORZ = 0

  propScrollBarVisible = newValue
  ShowScrollBar picContainer.hWnd, SB_HORZ, Abs(newValue)
End Property

Private Property Get ScrollValue() As Long
  Const SB_HORZ = 0
  Const SIF_POS = &H4
  Dim scrInfo As SCROLLINFO

  scrInfo.cbSize = LenB(scrInfo)
  scrInfo.fMask = SIF_POS
  GetScrollInfo picContainer.hWnd, SB_HORZ, scrInfo
  ScrollValue = scrInfo.nPos
End Property

Private Property Let ScrollValue(ByVal newValue As Long)
  Const SB_HORZ = 0
  Const SIF_POS = &H4
  Dim scrInfo As SCROLLINFO

  If newValue <> ScrollValue Then
    scrInfo.cbSize = LenB(scrInfo)
    scrInfo.fMask = SIF_POS
    If newValue > MaxScrollValue Then
      newValue = MaxScrollValue
    End If
    If newValue < MinScrollValue Then
      newValue = MinScrollValue
    End If
    scrInfo.nPos = newValue
    SetScrollInfo picContainer.hWnd, SB_HORZ, scrInfo, 1
    ScollBarChanged
  End If
End Property

Private Property Get SmallScrollChange() As Long
  SmallScrollChange = propSmallScrollChange
End Property

Private Property Let SmallScrollChange(ByVal newValue As Long)
  propSmallScrollChange = newValue
End Property


Private Sub InsertItems()
  Dim cImages As Long
  Dim iIcon As Long
  Dim tiU1 As ExTVwLibUCtl.TreeViewItem
  Dim tiU2 As ExTVwLibUCtl.TreeViewItem

  ExTvwU.hImageList(ImageListConstants.ilItems) = hImgLst
  cImages = ImageList_GetImageCount(hImgLst)

  With ExTvwU.TreeItems
    Set tiU1 = .Add("Level 1, Item 1", , , heYes, iIcon)
    iIcon = (iIcon + 1) Mod cImages
    With tiU1.SubItems
      .Add "Level 2, Item 1", , , heNo, iIcon
      iIcon = (iIcon + 1) Mod cImages
      tiU1.Expand
      Set tiU2 = .Add("Level 2, Item 2", , , heYes, iIcon)
      iIcon = (iIcon + 1) Mod cImages
      With tiU2.SubItems
        .Add "Level 3, Item 1", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        .Add "Level 3, Item 2", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        .Add "Level 3, Item 3", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        tiU2.Expand
      End With
      .Add "Level 2, Item 3", , , heNo, iIcon
      iIcon = (iIcon + 1) Mod cImages
      .Add "Level 2, Item 4", , , heNo, iIcon
      iIcon = (iIcon + 1) Mod cImages
      Set tiU2 = .Add("Level 2, Item 5", , , heYes, iIcon)
      iIcon = (iIcon + 1) Mod cImages
      With tiU2.SubItems
        .Add "Level 3, Item 4", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        .Add "Level 3, Item 5", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        tiU2.Expand
      End With
    End With
  End With
End Sub

Private Sub ScollBarChanged()
  SizeMultiColumnTreeView
End Sub

Private Sub ScollBarScrolled()
  SizeMultiColumnTreeView
End Sub

Private Sub SizeMultiColumnTreeView()
  Const HDI_WIDTH = &H1
  Const HDM_FIRST = &H1200
  Const HDM_GETITEM = (HDM_FIRST + 11)
  Const HDM_GETITEMCOUNT = (HDM_FIRST + 0)
  Const HWND_TOP = 0
  Dim cColumns As Long
  Dim clientRectangle As RECT
  Dim cx As Long
  Dim hdi As HDITEM
  Dim i As Long
  Dim scrollX As Long

  GetClientRect picContainer.hWnd, clientRectangle
  If hWndHeader Then
    If ScrollBarVisible Then
      scrollX = -ScrollValue
    End If
    cColumns = SendMessageAsLong(hWndHeader, HDM_GETITEMCOUNT, 0, 0)
    hdi.mask = HDI_WIDTH
    For i = 0 To cColumns - 1
      SendMessageAsLong hWndHeader, HDM_GETITEM, i, VarPtr(hdi)
      cx = cx + hdi.cxy
    Next i
    If cx < clientRectangle.Right - clientRectangle.Left Then
      If ScrollBarVisible Then
        ScrollValue = 0
        scrollX = 0
        ScrollBarVisible = False
        GetClientRect picContainer.hWnd, clientRectangle
      End If
      cx = clientRectangle.Right - clientRectangle.Left
    Else
      If MaxScrollValue <> cx - (clientRectangle.Right - clientRectangle.Left) Then
        MaxScrollValue = cx - (clientRectangle.Right - clientRectangle.Left)
        If scrollX > MaxScrollValue Then
          scrollX = MaxScrollValue
        End If
        LargeScrollChange = clientRectangle.Right - clientRectangle.Left
      End If
      If Not ScrollBarVisible Then
        ScrollBarVisible = True
        GetClientRect picContainer.hWnd, clientRectangle
      End If
    End If
    SetWindowPos hWndHeader, HWND_TOP, scrollX + clientRectangle.Left, clientRectangle.Top, cx, 20, 0
    clientRectangle.Top = clientRectangle.Top + 20
  Else
    cx = clientRectangle.Right - clientRectangle.Left
  End If
  ExTvwU.Move scrollX + clientRectangle.Left, clientRectangle.Top, clientRectangle.Right - clientRectangle.Left - scrollX, clientRectangle.Bottom - clientRectangle.Top
End Sub

' subclasses this Form
Private Sub Subclass()
  Const NF_REQUERY = 4
  Const WM_NOTIFYFORMAT = &H55

  #If UseSubClassing Then
    If Not SubclassWindow(picContainer.hWnd, Me, EnumSubclassID.escidFrmMainPicContainer) Then
      Debug.Print "Subclassing failed!"
    End If
    ' tell the controls to negotiate the correct format with the form
    SendMessageAsLong ExTvwU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
