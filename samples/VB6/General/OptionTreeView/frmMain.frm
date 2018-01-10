VERSION 5.00
Object = "{1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5}#2.6#0"; "ExTVwU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ExplorerTreeView 2.6 - OptionTreeView Sample"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
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
   ScaleHeight     =   307
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   346
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
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin ExTVwLibUCtl.ExplorerTreeView ExTvwU 
      Align           =   3  'Links ausrichten
      Height          =   4605
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _cx             =   6376
      _cy             =   8123
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
      DisabledEvents  =   264191
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
      ShowStateImages =   -1  'True
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
  Private hStateImgLst As Long
  Private themeableOS As Boolean


  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hicon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (data As DLLVERSIONINFO) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hicon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hicon As Long) As Long
  Private Declare Function ImageList_SetImageCount Lib "comctl32.dll" (ByVal himl As Long, ByVal uNewCount As Long) As Long
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

Private Sub ExTvwU_ItemStateImageChanging(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal previousStateImageIndex As Long, newStateImageIndex As Long, ByVal causedBy As ExTVwLibUCtl.StateImageChangeCausedByConstants, cancelChange As Boolean)
  Dim col As ExTVwLibUCtl.TreeViewItems
  Dim itm As ExTVwLibUCtl.TreeViewItem

  If causedBy <> StateImageChangeCausedByConstants.siccbCode Then
    Select Case treeItem.ItemData
      Case 1
        ' CheckBox
        If previousStateImageIndex = 1 Then
          ' check the item
          newStateImageIndex = 2
        Else
          ' un-check the item
          newStateImageIndex = 1
        End If
      Case 2
        ' OptionButton
        If previousStateImageIndex = 3 Then
          ' uncheck all items that have the same parent
          Set col = treeItem.ParentItem.SubItems
          col.FilterType(FilteredPropertyConstants.fpStateImageIndex) = FilterTypeConstants.ftIncluding
          col.Filter(FilteredPropertyConstants.fpStateImageIndex) = Array(4)

          For Each itm In col
            itm.StateImageIndex = 3
          Next itm
          ' check the item
          newStateImageIndex = 4
        Else
          ' leave the item checked
          newStateImageIndex = 4
        End If
    End Select
  End If
End Sub

Private Sub ExTvwU_RecreatedControlWindow(ByVal hWndTreeView As Long)
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
  Dim hicon As Long
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
  hStateImgLst = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 5, 0)
  ImageList_SetImageCount hStateImgLst, 5
  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
  iconPath = Dir$(iconsDir & "*.ico")
  While iconPath <> ""
    hicon = LoadImage(0, StrPtr(iconsDir & iconPath), IMAGE_ICON, 16, 16, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hicon Then
      Select Case LCase$(iconPath)
        Case "empty.ico"
          ImageList_ReplaceIcon hStateImgLst, 0, hicon
        Case "unchecked.ico"
          ImageList_ReplaceIcon hStateImgLst, 1, hicon
        Case "checked.ico"
          ImageList_ReplaceIcon hStateImgLst, 2, hicon
        Case "option unchecked.ico"
          ImageList_ReplaceIcon hStateImgLst, 3, hicon
        Case "option checked.ico"
          ImageList_ReplaceIcon hStateImgLst, 4, hicon
        Case Else
          ImageList_AddIcon hImgLst, hicon
      End Select
      DestroyIcon hicon
    End If
    iconPath = Dir$
  Wend
End Sub

Private Sub Form_Load()
  ' this is required to make the control work as expected
  Subclass

  InsertItems
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
  If hStateImgLst Then ImageList_Destroy hStateImgLst
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub


Private Sub InsertItems()
  Dim itm As ExTVwLibUCtl.TreeViewItem

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExTvwU.hWnd, StrPtr("explorer"), 0
  End If

  ExTvwU.hImageList(ImageListConstants.ilItems) = hImgLst
  ExTvwU.hImageList(ImageListConstants.ilState) = hStateImgLst
  With ExTvwU.TreeItems
    Set itm = .Add("Check 1", Nothing, , heYes, 0, , , 1)
    itm.StateImageIndex = 1
    .Add("Option 1", itm, , heNo, 1, , , 2).StateImageIndex = 3
    itm.Expand
    .Add("Option 2", itm, , heNo, 2, , , 2).StateImageIndex = 3
    .Add("Option 3", itm, , heNo, 3, , , 2).StateImageIndex = 3
    Set itm = .Add("Check 2", Nothing, , heYes, 4, , , 1)
    itm.StateImageIndex = 1
    .Add("Option 1", itm, , heNo, 5, , , 2).StateImageIndex = 3
    itm.Expand
    .Add("Option 2", itm, , heNo, 6, , , 2).StateImageIndex = 3
    .Add("Option 3", itm, , heNo, 7, , , 2).StateImageIndex = 3
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
    ' tell the control to negotiate the correct format with the form
    SendMessageAsLong ExTvwU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
