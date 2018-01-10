VERSION 5.00
Object = "{1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5}#2.6#0"; "ExTVwU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ExplorerTreeView 2.6 - Disabled Items Sample"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
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
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   610
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkEnabled 
      Caption         =   "&Enabled"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   975
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
      Left            =   3908
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin ExTVwLibUCtl.ExplorerTreeView ExTvwU2 
      Height          =   3495
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   3615
      _cx             =   6376
      _cy             =   6165
      AllowDragDrop   =   0   'False
      AllowLabelEditing=   -1  'True
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
      DisabledEvents  =   263423
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
      ScrollBars      =   2
      ShowStateImages =   0   'False
      ShowToolTips    =   -1  'True
      SingleExpand    =   0
      SortOrder       =   0
      SupportOLEDragImages=   -1  'True
      TreeViewStyle   =   1
      UseSystemFont   =   -1  'True
   End
   Begin ExTVwLibUCtl.ExplorerTreeView ExTvwU1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _cx             =   6376
      _cy             =   5530
      AllowDragDrop   =   0   'False
      AllowLabelEditing=   -1  'True
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
      DisabledEvents  =   263423
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
      ScrollBars      =   2
      ShowStateImages =   0   'False
      ShowToolTips    =   -1  'True
      SingleExpand    =   0
      SortOrder       =   0
      SupportOLEDragImages=   -1  'True
      TreeViewStyle   =   1
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


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type


  Private bComctl32Version600OrNewer As Boolean
  Private hImgLst As Long
  Private themeableOS As Boolean


  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (data As DLLVERSIONINFO) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal i As Long) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
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


Private Sub chkEnabled_Click()
  If chkEnabled.Value = CheckBoxConstants.vbChecked Then
    ExTvwU1.CaretItem.ItemData = 0
    ExTvwU1.Refresh
  ElseIf chkEnabled.Value = CheckBoxConstants.vbUnchecked Then
    ExTvwU1.CaretItem.ItemData = 1
    ExTvwU1.Refresh
  End If
End Sub

Private Sub cmdAbout_Click()
  ExTvwU1.About
End Sub

Private Sub ExTvwU1_CaretChanged(ByVal previousCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal newCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal caretChangeReason As ExTVwLibUCtl.CaretChangeCausedByConstants)
  ' Is the new caret disabled?
  Select Case newCaretItem.ItemData
    Case 0
      ' no
      chkEnabled.Value = CheckBoxConstants.vbChecked
    Case 1
      ' yes
      chkEnabled.Value = CheckBoxConstants.vbUnchecked
  End Select
End Sub

Private Sub ExTvwU1_CollapsingItem(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal removingAllSubItems As Boolean, cancelCollapse As Boolean)
  If treeItem.ItemData = 1 Then
    ' this is a disabled item
    cancelCollapse = True
  End If
End Sub

Private Sub ExTvwU1_CustomDraw(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, itemLevel As Long, textColor As stdole.OLE_COLOR, textBackColor As stdole.OLE_COLOR, ByVal drawStage As ExTVwLibUCtl.CustomDrawStageConstants, ByVal itemState As ExTVwLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExTVwLibUCtl.RECTANGLE, furtherProcessing As ExTVwLibUCtl.CustomDrawReturnValuesConstants)
  Const COLOR_3DFACE = 15
  Const COLOR_GRAYTEXT = 17
  Const COLOR_HIGHLIGHT = 13
  Const COLOR_WINDOWTEXT = 8

  Select Case drawStage
    Case CustomDrawStageConstants.cdsPrePaint
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvNotifyItemDraw

    Case CustomDrawStageConstants.cdsItemPrePaint
      If treeItem.ItemData = 1 Then
        ' this is a disabled item
        If itemState And CustomDrawItemStateConstants.cdisSelected Then
        ' use the following line instead of the last line, if you've set AllowDragDrop to True
        'If textBackColor = GetSysColor(COLOR_HIGHLIGHT) Then
          If itemState And CustomDrawItemStateConstants.cdisFocus Then
            textBackColor = GetSysColor(COLOR_3DFACE)
            textColor = GetSysColor(COLOR_WINDOWTEXT)
          End If
          ' use the following code instead of the last If block, if you've set AllowDragDrop to True
          'textBackColor = GetSysColor(COLOR_3DFACE)
          'textColor = GetSysColor(COLOR_WINDOWTEXT)
        Else
          textColor = GetSysColor(COLOR_GRAYTEXT)
        End If

        furtherProcessing = CustomDrawReturnValuesConstants.cdrvNewFont
      End If
  End Select
End Sub

Private Sub ExTvwU1_ExpandingItem(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal expandingPartially As Boolean, cancelExpansion As Boolean)
  If treeItem.ItemData = 1 Then
    ' this is a disabled item
    cancelExpansion = True
  End If
End Sub

Private Sub ExTvwU1_ItemGetInfoTipText(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If treeItem.ItemData = 1 Then
    ' this is a disabled item
    abortToolTip = True
  End If
End Sub

Private Sub ExTvwU1_RecreatedControlWindow(ByVal hWndTreeView As Long)
  InsertItems1
End Sub

Private Sub ExTvwU1_StartingLabelEditing(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, cancelEditing As Boolean)
  If treeItem.ItemData = 1 Then
    ' this is a disabled item
    cancelEditing = True
  End If
End Sub

Private Sub ExTvwU2_CaretChanging(ByVal previousCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal newCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal caretChangeReason As ExTVwLibUCtl.CaretChangeCausedByConstants, cancelCaretChange As Boolean)
  Dim bFirstChild As Boolean

  If newCaretItem.ItemData = 1 Then
    ' this is a disabled item
    ' select its first child instead
    If previousCaretItem Is Nothing Then
      bFirstChild = True
    ElseIf newCaretItem.FirstSubItem.Handle <> previousCaretItem.Handle Then
      bFirstChild = True
    End If

    If bFirstChild Then
      If newCaretItem.FirstSubItem Is Nothing Then
        ' select the next sibling item
        If Not (newCaretItem.NextSiblingItem Is Nothing) Then
          Set ExTvwU2.CaretItem = newCaretItem.NextSiblingItem
        End If
      Else
        Set ExTvwU2.CaretItem = newCaretItem.FirstSubItem
      End If
    End If
    cancelCaretChange = True
  End If
End Sub

Private Sub ExTvwU2_CollapsingItem(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal removingAllSubItems As Boolean, cancelCollapse As Boolean)
  If treeItem.ItemData = 1 Then
    ' this is a disabled item
    cancelCollapse = True
  End If
End Sub

Private Sub ExTvwU2_CustomDraw(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, itemLevel As Long, textColor As stdole.OLE_COLOR, textBackColor As stdole.OLE_COLOR, ByVal drawStage As ExTVwLibUCtl.CustomDrawStageConstants, ByVal itemState As ExTVwLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExTVwLibUCtl.RECTANGLE, furtherProcessing As ExTVwLibUCtl.CustomDrawReturnValuesConstants)
  Const COLOR_3DFACE = 15
  Const COLOR_GRAYTEXT = 17
  Const COLOR_HIGHLIGHT = 13
  Const COLOR_WINDOWTEXT = 8

  Select Case drawStage
    Case CustomDrawStageConstants.cdsPrePaint
      furtherProcessing = CustomDrawReturnValuesConstants.cdrvNotifyItemDraw

    Case CustomDrawStageConstants.cdsItemPrePaint
      If treeItem.ItemData = 1 Then
        ' this is a disabled item
        If itemState And CustomDrawItemStateConstants.cdisSelected Then
        ' use the following line instead of the last line, if you've set AllowDragDrop to True
        'If textBackColor = GetSysColor(COLOR_HIGHLIGHT) Then
          ' use the following code, if you've set AllowDragDrop to True
          'textBackColor = GetSysColor(COLOR_3DFACE)
          'textColor = GetSysColor(COLOR_WINDOWTEXT)
        Else
          textColor = GetSysColor(COLOR_GRAYTEXT)
        End If

        furtherProcessing = CustomDrawReturnValuesConstants.cdrvNewFont
      End If
  End Select
End Sub

Private Sub ExTvwU2_ExpandingItem(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal expandingPartially As Boolean, cancelExpansion As Boolean)
  If treeItem.ItemData = 1 Then
    ' this is a disabled item
    cancelExpansion = True
  End If
End Sub

Private Sub ExTvwU2_ItemGetInfoTipText(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  abortToolTip = (treeItem.ItemData = 1)
End Sub

Private Sub ExTvwU2_RecreatedControlWindow(ByVal hWndTreeView As Long)
  InsertItems2
End Sub

Private Sub ExTvwU2_StartingLabelEditing(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, cancelEditing As Boolean)
  If treeItem.ItemData = 1 Then
    ' this is a disabled item
    cancelEditing = True
  End If
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
  ' this is required to make the control work as expected
  Subclass

  InsertItems1
  InsertItems2
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub


Private Sub InsertItems1()
  Dim cImages As Long
  Dim iIcon As Long
  Dim tiU1 As ExTVwLibUCtl.TreeViewItem
  Dim tiU2 As ExTVwLibUCtl.TreeViewItem

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExTvwU1.hWnd, StrPtr("explorer"), 0
  End If

  ExTvwU1.hImageList(ImageListConstants.ilItems) = hImgLst
  cImages = ImageList_GetImageCount(hImgLst)

  With ExTvwU1.TreeItems
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

Private Sub InsertItems2()
  Dim cImages As Long
  Dim iIcon As Long
  Dim tiU1 As ExTVwLibUCtl.TreeViewItem
  Dim tiU2 As ExTVwLibUCtl.TreeViewItem

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExTvwU2.hWnd, StrPtr("explorer"), 0
  End If

  ExTvwU2.hImageList(ImageListConstants.ilItems) = hImgLst
  cImages = ImageList_GetImageCount(hImgLst)

  With ExTvwU2.TreeItems
    Set tiU1 = .Add("Level 1, Item 1", , , heYes, iIcon)
    iIcon = (iIcon + 1) Mod cImages
    With tiU1.SubItems
      .Add "Level 2, Item 1", , , heNo, iIcon
      iIcon = (iIcon + 1) Mod cImages
      tiU1.Expand
      tiU1.ItemData = 1
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
        tiU2.ItemData = 1
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
        tiU2.ItemData = 1
      End With
    End With
  End With
End Sub

' subclasses this Form
Private Sub Subclass()
  Const NF_REQUERY = 4
  Const WM_NOTIFYFORMAT = &H55

  #If UseSubClassing Then
    If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
      Debug.Print "Subclassing failed!"
    End If
    ' tell the controls to negotiate the correct format with the form
    SendMessageAsLong ExTvwU1.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong ExTvwU2.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
