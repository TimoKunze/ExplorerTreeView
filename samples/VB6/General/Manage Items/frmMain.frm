VERSION 5.00
Object = "{1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5}#2.6#0"; "ExTVwU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "ExplorerTreeView 2.6 - Manage Items Sample"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5340
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
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
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
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin ExTVwLibUCtl.ExplorerTreeView ExTvwU 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _cx             =   6376
      _cy             =   8916
      AllowDragDrop   =   -1  'True
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
      DisabledEvents  =   264189
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
      Indent          =   16
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
   Begin VB.Label lblDescr 
      BackStyle       =   0  'Transparent
      Caption         =   "Right-click to add, remove or edit items."
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Menu mnuItemActions 
      Caption         =   "Item Actions"
      Visible         =   0   'False
      Begin VB.Menu mnuAddSubItem 
         Caption         =   "Add sub-item..."
      End
      Begin VB.Menu mnuAddItemAbove 
         Caption         =   "Add item above..."
      End
      Begin VB.Menu mnuAddItemBelow 
         Caption         =   "Add item below..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Edit..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveItem 
         Caption         =   "Remove item"
      End
   End
   Begin VB.Menu mnuTreeViewActions 
      Caption         =   "TreeView Actions"
      Visible         =   0   'False
      Begin VB.Menu mnuAddItem 
         Caption         =   "Add item..."
      End
      Begin VB.Menu mnuRemoveAllItems 
         Caption         =   "Remove all items"
      End
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


  Private activeItem As TreeViewItem
  Private bComctl32Version600OrNewer As Boolean
  Private hImgLst As Long
  Private themeableOS As Boolean


  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (data As DLLVERSIONINFO) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
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


Private Sub cmdAbout_Click()
  ExTvwU.About
End Sub

Private Sub ExTvwU_Click(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    If (hitTestDetails And (HitTestConstants.htItemStateImage Or HitTestConstants.htItemExpando)) = 0 Then
      ' clear caret
      Set ExTvwU.CaretItem = Nothing
    End If
  End If
End Sub

Private Sub ExTvwU_ContextMenu(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants, showDefaultMenu As Boolean)
  Set activeItem = treeItem
  If treeItem Is Nothing Then
    mnuRemoveAllItems.Enabled = (ExTvwU.TreeItems.Count > 0)
    PopupMenu mnuTreeViewActions, , ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels) + ExTvwU.Left, ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels) + ExTvwU.Top
  Else
    PopupMenu mnuItemActions, , ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels) + ExTvwU.Left, ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels) + ExTvwU.Top, mnuEditItem
  End If
End Sub

Private Sub ExTvwU_DblClick(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  Dim frm As frmItemProperties

  If Not (treeItem Is Nothing) Then
    Set frm = New frmItemProperties
    With frm
      .Bold = treeItem.Bold
      .Ghosted = treeItem.Ghosted
      .HasExpando = (treeItem.HasExpando = HasExpandoConstants.heYes)
      .IconIndex = treeItem.IconIndex
      .SelectedIconIndex = treeItem.SelectedIconIndex
      .ItemData = treeItem.ItemData
      .HeightIncrement = treeItem.HeightIncrement
      .Text = treeItem.Text
      .Show FormShowConstants.vbModal, Me
      If Not .Canceled Then
        treeItem.Bold = .Bold
        treeItem.Ghosted = .Ghosted
        treeItem.HasExpando = IIf(.HasExpando, HasExpandoConstants.heYes, HasExpandoConstants.heNo)
        treeItem.IconIndex = .IconIndex
        treeItem.SelectedIconIndex = .SelectedIconIndex
        treeItem.ItemData = .ItemData
        treeItem.HeightIncrement = .HeightIncrement
        treeItem.Text = .Text
      End If
      Unload frm
    End With
    Set frm = Nothing
  End If
End Sub

Private Sub ExTvwU_RecreatedControlWindow(ByVal hWndTreeView As Long)
  PrepareTreeView
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

  PrepareTreeView
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub mnuAddItem_Click()
  Dim frm As frmItemProperties
  Dim treeItem As TreeViewItem

  Set frm = New frmItemProperties
  With frm
    .HeightIncrement = 1
    .IconIndex = 0
    .ItemData = 0
    .SelectedIconIndex = 0
    .Text = "Item " & (ExTvwU.TreeItems.Count + 1) & ", Level 0"
    .Show FormShowConstants.vbModal, Me
    If Not .Canceled Then
      Set treeItem = ExTvwU.TreeItems.Add(.Text, , , IIf(.HasExpando, HasExpandoConstants.heYes, HasExpandoConstants.heNo), .IconIndex, .SelectedIconIndex, , .ItemData, , .HeightIncrement)
      treeItem.Bold = .Bold
      treeItem.Ghosted = .Ghosted
    End If
    Unload frm
  End With
  Set frm = Nothing
End Sub

Private Sub mnuAddItemAbove_Click()
  Dim frm As frmItemProperties
  Dim treeItem As TreeViewItem

  Set frm = New frmItemProperties
  With frm
    .HeightIncrement = 1
    .IconIndex = 0
    .ItemData = 0
    .SelectedIconIndex = 0
    .Text = "Item " & (ExTvwU.TreeItems.Count + 1) & ", Level " & activeItem.Level
    .Show FormShowConstants.vbModal, Me
    If Not .Canceled Then
      Set treeItem = ExTvwU.TreeItems.Add(.Text, activeItem.ParentItem, activeItem.PreviousSiblingItem, IIf(.HasExpando, HasExpandoConstants.heYes, HasExpandoConstants.heNo), .IconIndex, .SelectedIconIndex, , .ItemData, , .HeightIncrement)
      treeItem.Bold = .Bold
      treeItem.Ghosted = .Ghosted
      Set ExTvwU.CaretItem = treeItem
    End If
    Unload frm
  End With
  Set frm = Nothing
End Sub

Private Sub mnuAddItemBelow_Click()
  Dim frm As frmItemProperties
  Dim treeItem As TreeViewItem

  Set frm = New frmItemProperties
  With frm
    .HeightIncrement = 1
    .IconIndex = 0
    .ItemData = 0
    .SelectedIconIndex = 0
    .Text = "Item " & (ExTvwU.TreeItems.Count + 1) & ", Level " & activeItem.Level
    .Show FormShowConstants.vbModal, Me
    If Not .Canceled Then
      Set treeItem = ExTvwU.TreeItems.Add(.Text, activeItem.ParentItem, activeItem, IIf(.HasExpando, HasExpandoConstants.heYes, HasExpandoConstants.heNo), .IconIndex, .SelectedIconIndex, , .ItemData, , .HeightIncrement)
      treeItem.Bold = .Bold
      treeItem.Ghosted = .Ghosted
      Set ExTvwU.CaretItem = treeItem
    End If
    Unload frm
  End With
  Set frm = Nothing
End Sub

Private Sub mnuAddSubItem_Click()
  Dim frm As frmItemProperties
  Dim treeItem As TreeViewItem

  Set frm = New frmItemProperties
  With frm
    .HeightIncrement = 1
    .IconIndex = 0
    .ItemData = 0
    .SelectedIconIndex = 0
    .Text = "Item " & (ExTvwU.TreeItems.Count + 1) & ", Level " & (activeItem.Level + 1)
    .Show FormShowConstants.vbModal, Me
    If Not .Canceled Then
      Set treeItem = ExTvwU.TreeItems.Add(.Text, activeItem, , IIf(.HasExpando, HasExpandoConstants.heYes, HasExpandoConstants.heNo), .IconIndex, .SelectedIconIndex, , .ItemData, , .HeightIncrement)
      treeItem.Bold = .Bold
      treeItem.Ghosted = .Ghosted
      Set ExTvwU.CaretItem = treeItem
    End If
    Unload frm
  End With
  Set frm = Nothing
End Sub

Private Sub mnuEditItem_Click()
  Dim frm As frmItemProperties

  Set frm = New frmItemProperties
  With frm
    .Bold = activeItem.Bold
    .Ghosted = activeItem.Ghosted
    .HasExpando = (activeItem.HasExpando = HasExpandoConstants.heYes)
    .IconIndex = activeItem.IconIndex
    .SelectedIconIndex = activeItem.SelectedIconIndex
    .ItemData = activeItem.ItemData
    .HeightIncrement = activeItem.HeightIncrement
    .Text = activeItem.Text
    .Show FormShowConstants.vbModal, Me
    If Not .Canceled Then
      activeItem.Bold = .Bold
      activeItem.Ghosted = .Ghosted
      activeItem.HasExpando = IIf(.HasExpando, HasExpandoConstants.heYes, HasExpandoConstants.heNo)
      activeItem.IconIndex = .IconIndex
      activeItem.SelectedIconIndex = .SelectedIconIndex
      activeItem.ItemData = .ItemData
      activeItem.HeightIncrement = .HeightIncrement
      activeItem.Text = .Text
    End If
    Unload frm
  End With
  Set frm = Nothing
End Sub

Private Sub mnuRemoveAllItems_Click()
  ExTvwU.TreeItems.RemoveAll
  Set activeItem = Nothing
End Sub

Private Sub mnuRemoveItem_Click()
  ExTvwU.TreeItems.Remove activeItem.Handle
  Set activeItem = Nothing
End Sub


Private Sub PrepareTreeView()
  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExTvwU.hWnd, StrPtr("explorer"), 0
  End If

  ExTvwU.hImageList(ImageListConstants.ilItems) = hImgLst
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
    SendMessageAsLong ExTvwU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
