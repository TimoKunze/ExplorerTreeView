VERSION 5.00
Object = "{1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5}#2.6#0"; "ExTVwU.ocx"
Object = "{0725EC44-6E15-4E66-A561-A94FDA7E5DCB}#2.6#0"; "ExTVwA.ocx"
Begin VB.Form frmMain 
   Caption         =   "ExplorerTreeView 2.6 - Events Sample"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
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
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkLog 
      Caption         =   "&Log"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   5400
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
      Left            =   4830
      TabIndex        =   2
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox txtLog 
      Height          =   4815
      Left            =   3720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
   Begin ExTVwLibUCtl.ExplorerTreeView ExTVwU 
      Align           =   3  'Links ausrichten
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _cx             =   6376
      _cy             =   10186
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
      DisabledEvents  =   0
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
      RegisterForOLEDragDrop=   -1  'True
      RichToolTips    =   0   'False
      RightToLeft     =   0
      ScrollBars      =   2
      ShowStateImages =   -1  'True
      ShowToolTips    =   -1  'True
      SingleExpand    =   0
      SortOrder       =   0
      SupportOLEDragImages=   -1  'True
      TreeViewStyle   =   3
      UseSystemFont   =   -1  'True
   End
   Begin ExTVwLibACtl.ExplorerTreeView ExTVwA 
      Align           =   4  'Rechts ausrichten
      Height          =   5775
      Left            =   8265
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      _cx             =   6376
      _cy             =   10186
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
      DisabledEvents  =   0
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
      RegisterForOLEDragDrop=   -1  'True
      RichToolTips    =   0   'False
      RightToLeft     =   0
      ScrollBars      =   2
      ShowStateImages =   -1  'True
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


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type


  Private bComctl32Version600OrNewer As Boolean
  Private bLog As Boolean
  Private hImgLst As Long
  Private objActiveCtl As Object
  Private themeableOS As Boolean


  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (data As DLLVERSIONINFO) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
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


Private Sub chkLog_Click()
  bLog = (chkLog.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  objActiveCtl.About
End Sub

Private Sub ExTVwA_AbortedDrag()
  AddLogEntry "ExTVwA_AbortedDrag"
  Set ExTVwA.DropHilitedItem = Nothing
End Sub

Private Sub ExTVwA_CaretChanged(ByVal previousCaretItem As ExTVwLibACtl.ITreeViewItem, ByVal newCaretItem As ExTVwLibACtl.ITreeViewItem, ByVal caretChangeReason As ExTVwLibACtl.CaretChangeCausedByConstants)
  Dim str As String

  str = "ExTVwA_CaretChanged: previousCaretItem="
  If previousCaretItem Is Nothing Then
    str = str & "Nothing, newCaretItem="
  Else
    str = str & previousCaretItem.Text & ", newCaretItem="
  End If
  If newCaretItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newCaretItem.Text
  End If
  str = str & ", caretChangeReason=" & caretChangeReason

  AddLogEntry str
End Sub

Private Sub ExTVwA_CaretChanging(ByVal previousCaretItem As ExTVwLibACtl.ITreeViewItem, ByVal newCaretItem As ExTVwLibACtl.ITreeViewItem, ByVal caretChangeReason As ExTVwLibACtl.CaretChangeCausedByConstants, cancelCaretChange As Boolean)
  Dim str As String

  str = "ExTVwA_CaretChanging: previousCaretItem="
  If previousCaretItem Is Nothing Then
    str = str & "Nothing, newCaretItem="
  Else
    str = str & previousCaretItem.Text & ", newCaretItem="
  End If
  If newCaretItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newCaretItem.Text
  End If
  str = str & ", caretChangeReason=" & caretChangeReason & ", cancelCaretChange=" & cancelCaretChange

  AddLogEntry str
End Sub

Private Sub ExTVwA_ChangedSortOrder(ByVal previousSortOrder As ExTVwLibACtl.SortOrderConstants, ByVal newSortOrder As ExTVwLibACtl.SortOrderConstants)
  AddLogEntry "ExTVwA_ChangedSortOrder: previousSortOrder=" & previousSortOrder & ", newSortOrder=" & newSortOrder
End Sub

Private Sub ExTVwA_ChangingSortOrder(ByVal previousSortOrder As ExTVwLibACtl.SortOrderConstants, ByVal newSortOrder As ExTVwLibACtl.SortOrderConstants, cancelChange As Boolean)
  AddLogEntry "ExTVwA_ChangingSortOrder: previousSortOrder=" & previousSortOrder & ", newSortOrder=" & newSortOrder & ", cancelChange=" & cancelChange
End Sub

Private Sub ExTVwA_Click(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_Click: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_Click: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwA_CollapsedItem(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal removedAllSubItems As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_CollapsedItem: treeItem=Nothing, removedAllSubItems=" & removedAllSubItems
  Else
    AddLogEntry "ExTVwA_CollapsedItem: treeItem=" & treeItem.Text & ", removedAllSubItems=" & removedAllSubItems
  End If
End Sub

Private Sub ExTVwA_CollapsingItem(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal removingAllSubItems As Boolean, cancelCollapse As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_CollapsingItem: treeItem=Nothing, removingAllSubItems=" & removingAllSubItems & ", cancelCollapse=" & cancelCollapse
  Else
    AddLogEntry "ExTVwA_CollapsingItem: treeItem=" & treeItem.Text & ", removingAllSubItems=" & removingAllSubItems & ", cancelCollapse=" & cancelCollapse
  End If
End Sub

Private Sub ExTVwA_CompareItems(ByVal FirstItem As ExTVwLibACtl.ITreeViewItem, ByVal secondItem As ExTVwLibACtl.ITreeViewItem, result As ExTVwLibACtl.CompareResultConstants)
  Dim str As String

  str = "ExTVwA_CompareItems: firstItem="
  If FirstItem Is Nothing Then
    str = str & "Nothing, secondItem="
  Else
    str = str & FirstItem.Text & ", secondItem="
  End If
  If secondItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & secondItem.Text
  End If
  str = str & ", result=" & result

  AddLogEntry str
End Sub

Private Sub ExTVwA_ContextMenu(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants, showDefaultMenu As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ContextMenu: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails & ", showDefaultMenu=" & showDefaultMenu
  Else
    AddLogEntry "ExTVwA_ContextMenu: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails & ", showDefaultMenu=" & showDefaultMenu
  End If
End Sub

Private Sub ExTVwA_CreatedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "ExTVwA_CreatedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub ExTVwA_CustomDraw(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, itemLevel As Long, textColor As stdole.OLE_COLOR, textBackColor As stdole.OLE_COLOR, ByVal drawStage As ExTVwLibACtl.CustomDrawStageConstants, ByVal itemState As ExTVwLibACtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExTVwLibACtl.RECTANGLE, furtherProcessing As ExTVwLibACtl.CustomDrawReturnValuesConstants)
'  If treeItem Is Nothing Then
'    addLogEntry "ExTVwA_CustomDraw: treeItem=Nothing, itemLevel=" & itemLevel & ", textColor=0x" & Hex(textColor) & ", textBackColor=0x" & Hex(textBackColor) & ", drawStage=" & drawStage & ", itemState=" & itemState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  Else
'    addLogEntry "ExTVwA_CustomDraw: treeItem=" & treeItem.Text & ", itemLevel=" & itemLevel & ", textColor=0x" & Hex(textColor) & ", textBackColor=0x" & Hex(textBackColor) & ", drawStage=" & drawStage & ", itemState=" & itemState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  End If
End Sub

Private Sub ExTVwA_DblClick(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_DblClick: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_DblClick: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ExTVwA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ExTVwA_DestroyedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "ExTVwA_DestroyedEditControlWindow: hWndEdit=" & hWndEdit
End Sub

Private Sub ExTVwA_DragDrop(Source As Control, X As Single, Y As Single)
  AddLogEntry "ExTVwA_DragDrop"
End Sub

Private Sub ExTVwA_DragMouseMove(dropTarget As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants, autoExpandItem As Boolean, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  If dropTarget Is Nothing Then
    AddLogEntry "ExTVwA_DragMouseMove: dropTarget=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails & ", autoExpandItem=" & autoExpandItem & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
  Else
    AddLogEntry "ExTVwA_DragMouseMove: dropTarget=" & dropTarget.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails & ", autoExpandItem=" & autoExpandItem & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
  End If
End Sub

Private Sub ExTVwA_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  AddLogEntry "ExTVwA_DragOver"
End Sub

Private Sub ExTVwA_Drop(ByVal dropTarget As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If dropTarget Is Nothing Then
    AddLogEntry "ExTVwA_Drop: dropTarget=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_Drop: dropTarget=" & dropTarget.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails
  End If
  Set ExTVwA.DropHilitedItem = Nothing
End Sub

Private Sub ExTVwA_EditChange()
  AddLogEntry "ExTVwA_EditChange - " & ExTVwA.EditText
End Sub

Private Sub ExTVwA_EditClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwA_EditClick: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwA_EditContextMenu(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, showDefaultMenu As Boolean)
  AddLogEntry "ExTVwA_EditContextMenu: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub ExTVwA_EditDblClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwA_EditDblClick: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwA_EditGotFocus()
  AddLogEntry "ExTVwA_EditGotFocus"
End Sub

Private Sub ExTVwA_EditKeyDown(keyCode As Integer, ByVal Shift As Integer)
  AddLogEntry "ExTVwA_EditKeyDown: keyCode=" & keyCode & ", shift=" & Shift
End Sub

Private Sub ExTVwA_EditKeyPress(keyAscii As Integer)
  AddLogEntry "ExTVwA_EditKeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ExTVwA_EditKeyUp(keyCode As Integer, ByVal Shift As Integer)
  AddLogEntry "ExTVwA_EditKeyUp: keyCode=" & keyCode & ", shift=" & Shift
End Sub

Private Sub ExTVwA_EditLostFocus()
  AddLogEntry "ExTVwA_EditLostFocus"
End Sub

Private Sub ExTVwA_EditMClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwA_EditMClick: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwA_EditMDblClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwA_EditMDblClick: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwA_EditMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwA_EditMouseDown: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwA_EditMouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwA_EditMouseEnter: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwA_EditMouseHover(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwA_EditMouseHover: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwA_EditMouseLeave(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwA_EditMouseLeave: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwA_EditMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwA_EditMouseMove: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwA_EditMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwA_EditMouseUp: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwA_EditMouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As ExTVwLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "ExTVwA_EditMouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub ExTVwA_EditRClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwA_EditRClick: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwA_EditRDblClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwA_EditRDblClick: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwA_EditXClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExTVwA_EditXClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExTVwA_EditXDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExTVwA_EditXDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExTVwA_ExpandedItem(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal expandedPartially As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ExpandedItem: treeItem=Nothing, expandedPartially=" & expandedPartially
  Else
    AddLogEntry "ExTVwA_ExpandedItem: treeItem=" & treeItem.Text & ", expandedPartially=" & expandedPartially
  End If
End Sub

Private Sub ExTVwA_ExpandingItem(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal expandingPartially As Boolean, cancelExpansion As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ExpandingItem: treeItem=Nothing, expandingPartially=" & expandingPartially & ", cancelExpansion=" & cancelExpansion
  Else
    AddLogEntry "ExTVwA_ExpandingItem: treeItem=" & treeItem.Text & ", expandingPartially=" & expandingPartially & ", cancelExpansion=" & cancelExpansion
  End If
End Sub

Private Sub ExTVwA_FreeItemData(ByVal treeItem As ExTVwLibACtl.ITreeViewItem)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_FreeItemData: treeItem=Nothing"
  Else
    AddLogEntry "ExTVwA_FreeItemData: treeItem=" & treeItem.Text
  End If
End Sub

Private Sub ExTVwA_GotFocus()
  AddLogEntry "ExTVwA_GotFocus"
  Set objActiveCtl = ExTVwA
End Sub

Private Sub ExTVwA_IncrementalSearchStringChanging(ByVal currentSearchString As String, ByVal keyCodeOfCharToBeAdded As Integer, cancelChange As Boolean)
  AddLogEntry "ExTVwA_IncrementalSearchStringChanging: currentSearchString=" & currentSearchString & ", keyCodeOfCharToBeAdded=" & keyCodeOfCharToBeAdded & ", cancelChange=" & cancelChange
End Sub

Private Sub ExTVwA_InsertedItem(ByVal treeItem As ExTVwLibACtl.ITreeViewItem)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_InsertedItem: treeItem=Nothing"
  Else
    AddLogEntry "ExTVwA_InsertedItem: treeItem=" & treeItem.Text
  End If
End Sub

Private Sub ExTVwA_InsertingItem(ByVal treeItem As ExTVwLibACtl.IVirtualTreeViewItem, cancelInsertion As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_InsertingItem: treeItem=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ExTVwA_InsertingItem: treeItem=" & treeItem.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ExTVwA_ItemAsynchronousDrawFailed(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, imageDetails As ExTVwLibACtl.FAILEDIMAGEDETAILS, ByVal failureReason As ExTVwLibACtl.ImageDrawingFailureReasonConstants, furtherProcessing As ExTVwLibACtl.FailedAsyncDrawReturnValuesConstants, newImageToDraw As Long)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ItemAsynchronousDrawFailed: treeItem=Nothing, failureReason=" & failureReason & ", furtherProcessing=" & furtherProcessing & ", newImageToDraw=" & newImageToDraw
  Else
    AddLogEntry "ExTVwA_ItemAsynchronousDrawFailed: treeItem=" & treeItem.Text & ", failureReason=" & failureReason & ", furtherProcessing=" & furtherProcessing & ", newImageToDraw=" & newImageToDraw
  End If
  AddLogEntry "     imageDetails.hImageList=0x" & Hex(imageDetails.hImageList)
  AddLogEntry "     imageDetails.hDC=0x" & Hex(imageDetails.hDC)
  AddLogEntry "     imageDetails.IconIndex=" & imageDetails.IconIndex
  AddLogEntry "     imageDetails.OverlayIndex=" & imageDetails.OverlayIndex
  AddLogEntry "     imageDetails.DrawingStyle=" & imageDetails.DrawingStyle
  AddLogEntry "     imageDetails.DrawingEffects=" & imageDetails.DrawingEffects
  AddLogEntry "     imageDetails.BackColor=0x" & Hex(imageDetails.BackColor)
  AddLogEntry "     imageDetails.ForeColor=0x" & Hex(imageDetails.ForeColor)
  AddLogEntry "     imageDetails.EffectColor=0x" & Hex(imageDetails.EffectColor)
End Sub

Private Sub ExTVwA_ItemBeginDrag(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ItemBeginDrag: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_ItemBeginDrag: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If

  On Error Resume Next
  ExTVwA.BeginDrag ExTVwA.CreateItemContainer(treeItem), -1
End Sub

Private Sub ExTVwA_ItemBeginRDrag(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ItemBeginRDrag: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_ItemBeginRDrag: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwA_ItemGetDisplayInfo(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal requestedInfo As ExTVwLibACtl.RequestedInfoConstants, IconIndex As Long, SelectedIconIndex As Long, ExpandedIconIndex As Long, HasExpando As Boolean, ByVal maxItemTextLength As Long, itemText As String, dontAskAgain As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ItemGetDisplayInfo: treeItem=Nothing, requestedInfo=" & requestedInfo & ", iconIndex=" & IconIndex & ", selectedIconIndex=" & SelectedIconIndex & ", expandedIconIndex=" & ExpandedIconIndex & ", hasExpando=" & HasExpando & ", maxItemTextLength=" & maxItemTextLength & ", itemText=" & itemText & ", dontAskAgain=" & dontAskAgain
  Else
    AddLogEntry "ExTVwA_ItemGetDisplayInfo: treeItem=0x" & Hex(treeItem.Handle) & ", requestedInfo=" & requestedInfo & ", iconIndex=" & IconIndex & ", selectedIconIndex=" & SelectedIconIndex & ", expandedIconIndex=" & ExpandedIconIndex & ", hasExpando=" & HasExpando & ", maxItemTextLength=" & maxItemTextLength & ", itemText=" & itemText & ", dontAskAgain=" & dontAskAgain
  End If
End Sub

Private Sub ExTVwA_ItemGetInfoTipText(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ItemGetInfoTipText: treeItem=Nothing, maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
  Else
    AddLogEntry "ExTVwA_ItemGetInfoTipText: treeItem=" & treeItem.Text & ", maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
    If infoTipText <> "" Then
      infoTipText = infoTipText & vbNewLine
    End If
    infoTipText = infoTipText & "ID: " & treeItem.ID & vbNewLine & "ItemData: " & treeItem.ItemData
  End If
End Sub

Private Sub ExTVwA_ItemMouseEnter(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ItemMouseEnter: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_ItemMouseEnter: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwA_ItemMouseLeave(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ItemMouseLeave: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_ItemMouseLeave: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwA_ItemSelectionChanged(ByVal treeItem As ExTVwLibACtl.ITreeViewItem)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ItemSelectionChanged: treeItem=Nothing"
  Else
    AddLogEntry "ExTVwA_ItemSelectionChanged: treeItem=" & treeItem.Text
  End If
End Sub

Private Sub ExTVwA_ItemSelectionChanging(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, cancelChange As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ItemSelectionChanging: treeItem=Nothing, cancelChange=" & cancelChange
  Else
    AddLogEntry "ExTVwA_ItemSelectionChanging: treeItem=" & treeItem.Text & ", cancelChange=" & cancelChange
  End If
End Sub

Private Sub ExTVwA_ItemSetText(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal itemText As String)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ItemSetText: treeItem=Nothing, itemText=" & itemText
  Else
    AddLogEntry "ExTVwA_ItemSetText: treeItem=0x" & Hex(treeItem.Handle) & ", itemText=" & itemText
  End If
End Sub

Private Sub ExTVwA_ItemStateImageChanged(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal previousStateImageIndex As Long, ByVal newStateImageIndex As Long, ByVal causedBy As ExTVwLibACtl.StateImageChangeCausedByConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ItemStateImageChanged: treeItem=Nothing, previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy
  Else
    AddLogEntry "ExTVwA_ItemStateImageChanged: treeItem=" & treeItem.Text & ", previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy
  End If
End Sub

Private Sub ExTVwA_ItemStateImageChanging(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal previousStateImageIndex As Long, newStateImageIndex As Long, ByVal causedBy As ExTVwLibACtl.StateImageChangeCausedByConstants, cancelChange As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_ItemStateImageChanging: treeItem=Nothing, previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy & ", cancelChange=" & cancelChange
  Else
    AddLogEntry "ExTVwA_ItemStateImageChanging: treeItem=" & treeItem.Text & ", previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy & ", cancelChange=" & cancelChange
  End If
End Sub

Private Sub ExTVwA_KeyDown(keyCode As Integer, ByVal Shift As Integer)
  AddLogEntry "ExTVwA_KeyDown: keyCode=" & keyCode & ", shift=" & Shift
  If keyCode = KeyCodeConstants.vbKeyF2 Then
    ExTVwA.CaretItem.StartLabelEditing
  End If
  If keyCode = KeyCodeConstants.vbKeyEscape Then
    If Not (ExTVwA.DraggedItems Is Nothing) Then ExTVwA.EndDrag True
  End If
End Sub

Private Sub ExTVwA_KeyPress(keyAscii As Integer)
  AddLogEntry "ExTVwA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ExTVwA_KeyUp(keyCode As Integer, ByVal Shift As Integer)
  AddLogEntry "ExTVwA_KeyUp: keyCode=" & keyCode & ", shift=" & Shift
End Sub

Private Sub ExTVwA_LostFocus()
  AddLogEntry "ExTVwA_LostFocus"
End Sub

Private Sub ExTVwA_MClick(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_MClick: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_MClick: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwA_MDblClick(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_MDblClick: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_MDblClick: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwA_MouseDown(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_MouseDown: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_MouseDown: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwA_MouseEnter(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_MouseEnter: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_MouseEnter: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwA_MouseHover(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_MouseHover: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_MouseHover: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwA_MouseLeave(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_MouseLeave: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_MouseLeave: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwA_MouseMove(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
'  If treeItem Is Nothing Then
'    AddLogEntry "ExTVwA_MouseMove: treeItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
'  Else
'    AddLogEntry "ExTVwA_MouseMove: treeItem=" & treeItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
'  End If
End Sub

Private Sub ExTVwA_MouseUp(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_MouseUp: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_MouseUp: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
  If Button = MouseButtonConstants.vbLeftButton Then
    If Not (ExTVwA.DraggedItems Is Nothing) Then
      ' Are we within the client area?
      If ((HitTestConstants.htAbove Or HitTestConstants.htBelow Or HitTestConstants.htToLeft Or HitTestConstants.htToRight) And hitTestDetails) = 0 Then
        ' yes
        ExTVwA.EndDrag False
      Else
        ' no
        ExTVwA.EndDrag True
      End If
    End If
  End If
End Sub

Private Sub ExTVwA_MouseWheel(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants, ByVal scrollAxis As ExTVwLibACtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  Dim str As String

  str = "ExTVwA_MouseWheel: treeItem="
  If treeItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & treeItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta

  AddLogEntry str
End Sub

Private Sub ExTVwA_OLECompleteDrag(ByVal data As ExTVwLibACtl.IOLEDataObject, ByVal performedEffect As ExTVwLibACtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "ExTVwA_OLECompleteDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub ExTVwA_OLEDragDrop(ByVal data As ExTVwLibACtl.IOLEDataObject, effect As ExTVwLibACtl.OLEDropEffectConstants, dropTarget As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ExTVwA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ExTVwA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ExTVwA_OLEDragEnter(ByVal data As ExTVwLibACtl.IOLEDataObject, effect As ExTVwLibACtl.OLEDropEffectConstants, dropTarget As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants, autoExpandItem As Boolean, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ExTVwA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails & ", autoExpandItem=" & autoExpandItem & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ExTVwA_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "ExTVwA_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub ExTVwA_OLEDragLeave(ByVal data As ExTVwLibACtl.IOLEDataObject, ByVal dropTarget As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ExTVwA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub ExTVwA_OLEDragLeavePotentialTarget()
  AddLogEntry "ExTVwA_OLEDragLeavePotentialTarget"
End Sub

Private Sub ExTVwA_OLEDragMouseMove(ByVal data As ExTVwLibACtl.IOLEDataObject, effect As ExTVwLibACtl.OLEDropEffectConstants, dropTarget As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants, autoExpandItem As Boolean, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ExTVwA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails & ", autoExpandItem=" & autoExpandItem & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ExTVwA_OLEGiveFeedback(ByVal effect As ExTVwLibACtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "ExTVwA_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub ExTVwA_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal Button As Integer, ByVal Shift As Integer, actionToContinueWith As ExTVwLibACtl.OLEActionToContinueWithConstants)
  AddLogEntry "ExTVwA_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & Button & ", shift=" & Shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub ExTVwA_OLEReceivedNewData(ByVal data As ExTVwLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ExTVwA_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ExTVwA_OLESetData(ByVal data As ExTVwLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ExTVwA_OLESetData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ExTVwA_OLEStartDrag(ByVal data As ExTVwLibACtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "ExTVwA_OLEStartDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub ExTVwA_RClick(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_RClick: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_RClick: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwA_RDblClick(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_RDblClick: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwA_RDblClick: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ExTVwA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertItemsA
End Sub

Private Sub ExTVwA_RemovedItem(ByVal treeItem As ExTVwLibACtl.IVirtualTreeViewItem)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_RemovedItem: treeItem=Nothing"
  Else
    AddLogEntry "ExTVwA_RemovedItem: treeItem=" & treeItem.Text
  End If
End Sub

Private Sub ExTVwA_RemovingItem(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, cancelDeletion As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_RemovingItem: treeItem=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ExTVwA_RemovingItem: treeItem=" & treeItem.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ExTVwA_RenamedItem(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal previousItemText As String, ByVal newItemText As String)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_RenamedItem: treeItem=Nothing, previousItemText=" & previousItemText & ", newItemText=" & newItemText
  Else
    AddLogEntry "ExTVwA_RenamedItem: treeItem=" & treeItem.Text & ", previousItemText=" & previousItemText & ", newItemText=" & newItemText
  End If
End Sub

Private Sub ExTVwA_RenamingItem(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal previousItemText As String, ByVal newItemText As String, cancelRenaming As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_RenamingItem: treeItem=Nothing, previousItemText=" & previousItemText & ", newItemText=" & newItemText & ", cancelRenaming=" & cancelRenaming
  Else
    AddLogEntry "ExTVwA_RenamingItem: treeItem=" & treeItem.Text & ", previousItemText=" & previousItemText & ", newItemText=" & newItemText & ", cancelRenaming=" & cancelRenaming
  End If
End Sub

Private Sub ExTVwA_ResizedControlWindow()
  AddLogEntry "ExTVwA_ResizedControlWindow"
End Sub

Private Sub ExTVwA_SingleExpanding(ByVal previousCaretItem As ExTVwLibACtl.ITreeViewItem, ByVal newCaretItem As ExTVwLibACtl.ITreeViewItem, dontChangePreviousItem As Boolean, dontChangeNewItem As Boolean)
  Dim str As String

  str = "ExTVwA_SingleExpanding: previousCaretItem="
  If previousCaretItem Is Nothing Then
    str = str & "Nothing, newCaretItem="
  Else
    str = str & previousCaretItem.Text & ", newCaretItem="
  End If
  If newCaretItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newCaretItem.Text
  End If
  str = str & ", dontChangePreviousItem=" & dontChangePreviousItem & ", dontChangeNewItem=" & dontChangeNewItem

  AddLogEntry str
End Sub

Private Sub ExTVwA_StartingLabelEditing(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, cancelEditing As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwA_StartingLabelEditing: treeItem=Nothing, cancelEditing=" & cancelEditing
  Else
    AddLogEntry "ExTVwA_StartingLabelEditing: treeItem=" & treeItem.Text & ", cancelEditing=" & cancelEditing
  End If
End Sub

Private Sub ExTVwA_Validate(Cancel As Boolean)
  AddLogEntry "ExTVwA_Validate"
End Sub

Private Sub ExTVwA_XClick(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExTVwA_XClick: treeItem="
  If treeItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & treeItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExTVwA_XDblClick(ByVal treeItem As ExTVwLibACtl.ITreeViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExTVwLibACtl.HitTestConstants)
  Dim str As String

  str = "ExTVwA_XDblClick: treeItem="
  If treeItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & treeItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExTVwU_AbortedDrag()
  AddLogEntry "ExTVwU_AbortedDrag"
  Set ExTVwU.DropHilitedItem = Nothing
End Sub

Private Sub ExTVwU_CaretChanged(ByVal previousCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal newCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal caretChangeReason As ExTVwLibUCtl.CaretChangeCausedByConstants)
  Dim str As String

  str = "ExTVwU_CaretChanged: previousCaretItem="
  If previousCaretItem Is Nothing Then
    str = str & "Nothing, newCaretItem="
  Else
    str = str & previousCaretItem.Text & ", newCaretItem="
  End If
  If newCaretItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newCaretItem.Text
  End If
  str = str & ", caretChangeReason=" & caretChangeReason

  AddLogEntry str
End Sub

Private Sub ExTVwU_CaretChanging(ByVal previousCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal newCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal caretChangeReason As ExTVwLibUCtl.CaretChangeCausedByConstants, cancelCaretChange As Boolean)
  Dim str As String

  str = "ExTVwU_CaretChanging: previousCaretItem="
  If previousCaretItem Is Nothing Then
    str = str & "Nothing, newCaretItem="
  Else
    str = str & previousCaretItem.Text & ", newCaretItem="
  End If
  If newCaretItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newCaretItem.Text
  End If
  str = str & ", caretChangeReason=" & caretChangeReason & ", cancelCaretChange=" & cancelCaretChange

  AddLogEntry str
End Sub

Private Sub ExTVwU_ChangedSortOrder(ByVal previousSortOrder As ExTVwLibUCtl.SortOrderConstants, ByVal newSortOrder As ExTVwLibUCtl.SortOrderConstants)
  AddLogEntry "ExTVwU_ChangedSortOrder: previousSortOrder=" & previousSortOrder & ", newSortOrder=" & newSortOrder
End Sub

Private Sub ExTVwU_ChangingSortOrder(ByVal previousSortOrder As ExTVwLibUCtl.SortOrderConstants, ByVal newSortOrder As ExTVwLibUCtl.SortOrderConstants, cancelChange As Boolean)
  AddLogEntry "ExTVwU_ChangingSortOrder: previousSortOrder=" & previousSortOrder & ", newSortOrder=" & newSortOrder & ", cancelChange=" & cancelChange
End Sub

Private Sub ExTVwU_Click(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_Click: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_Click: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwU_CollapsedItem(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal removedAllSubItems As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_CollapsedItem: treeItem=Nothing, removedAllSubItems=" & removedAllSubItems
  Else
    AddLogEntry "ExTVwU_CollapsedItem: treeItem=" & treeItem.Text & ", removedAllSubItems=" & removedAllSubItems
  End If
End Sub

Private Sub ExTVwU_CollapsingItem(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal removingAllSubItems As Boolean, cancelCollapse As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_CollapsingItem: treeItem=Nothing, removingAllSubItems=" & removingAllSubItems & ", cancelCollapse=" & cancelCollapse
  Else
    AddLogEntry "ExTVwU_CollapsingItem: treeItem=" & treeItem.Text & ", removingAllSubItems=" & removingAllSubItems & ", cancelCollapse=" & cancelCollapse
  End If
End Sub

Private Sub ExTVwU_CompareItems(ByVal FirstItem As ExTVwLibUCtl.ITreeViewItem, ByVal secondItem As ExTVwLibUCtl.ITreeViewItem, result As ExTVwLibUCtl.CompareResultConstants)
  Dim str As String

  str = "ExTVwU_CompareItems: firstItem="
  If FirstItem Is Nothing Then
    str = str & "Nothing, secondItem="
  Else
    str = str & FirstItem.Text & ", secondItem="
  End If
  If secondItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & secondItem.Text
  End If
  str = str & ", result=" & result

  AddLogEntry str
End Sub

Private Sub ExTVwU_ContextMenu(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants, showDefaultMenu As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ContextMenu: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails & ", showDefaultMenu=" & showDefaultMenu
  Else
    AddLogEntry "ExTVwU_ContextMenu: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails & ", showDefaultMenu=" & showDefaultMenu
  End If
End Sub

Private Sub ExTVwU_CreatedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "ExTVwU_CreatedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub ExTVwU_CustomDraw(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, itemLevel As Long, textColor As stdole.OLE_COLOR, textBackColor As stdole.OLE_COLOR, ByVal drawStage As ExTVwLibUCtl.CustomDrawStageConstants, ByVal itemState As ExTVwLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As ExTVwLibUCtl.RECTANGLE, furtherProcessing As ExTVwLibUCtl.CustomDrawReturnValuesConstants)
'  If treeItem Is Nothing Then
'    addLogEntry "ExTVwU_CustomDraw: treeItem=Nothing, itemLevel=" & itemLevel & ", textColor=0x" & Hex(textColor) & ", textBackColor=0x" & Hex(textBackColor) & ", drawStage=" & drawStage & ", itemState=" & itemState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  Else
'    addLogEntry "ExTVwU_CustomDraw: treeItem=" & treeItem.Text & ", itemLevel=" & itemLevel & ", textColor=0x" & Hex(textColor) & ", textBackColor=0x" & Hex(textBackColor) & ", drawStage=" & drawStage & ", itemState=" & itemState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  End If
End Sub

Private Sub ExTVwU_DblClick(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_DblClick: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_DblClick: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ExTVwU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ExTVwU_DestroyedEditControlWindow(ByVal hWndEdit As Long)
  AddLogEntry "ExTVwU_DestroyedEditControlWindow: hWndEdit=0x" & Hex(hWndEdit)
End Sub

Private Sub ExTVwU_DragDrop(Source As Control, X As Single, Y As Single)
  AddLogEntry "ExTVwU_DragDrop"
End Sub

Private Sub ExTVwU_DragMouseMove(dropTarget As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants, autoExpandItem As Boolean, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  If dropTarget Is Nothing Then
    AddLogEntry "ExTVwU_DragMouseMove: dropTarget=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails & ", autoExpandItem=" & autoExpandItem & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
  Else
    AddLogEntry "ExTVwU_DragMouseMove: dropTarget=" & dropTarget.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails & ", autoExpandItem=" & autoExpandItem & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity
  End If
End Sub

Private Sub ExTVwU_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  AddLogEntry "ExTVwU_DragOver"
End Sub

Private Sub ExTVwU_Drop(ByVal dropTarget As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If dropTarget Is Nothing Then
    AddLogEntry "ExTVwU_Drop: dropTarget=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_Drop: dropTarget=" & dropTarget.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails
  End If
  Set ExTVwU.DropHilitedItem = Nothing
End Sub

Private Sub ExTVwU_EditChange()
  AddLogEntry "ExTVwU_EditChange - " & ExTVwU.EditText
End Sub

Private Sub ExTVwU_EditClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwU_EditClick: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwU_EditContextMenu(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, showDefaultMenu As Boolean)
  AddLogEntry "ExTVwU_EditContextMenu: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", showDefaultMenu=" & showDefaultMenu
End Sub

Private Sub ExTVwU_EditDblClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwU_EditDblClick: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwU_EditGotFocus()
  AddLogEntry "ExTVwU_EditGotFocus"
End Sub

Private Sub ExTVwU_EditKeyDown(keyCode As Integer, ByVal Shift As Integer)
  AddLogEntry "ExTVwU_EditKeyDown: keyCode=" & keyCode & ", shift=" & Shift
End Sub

Private Sub ExTVwU_EditKeyPress(keyAscii As Integer)
  AddLogEntry "ExTVwU_EditKeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ExTVwU_EditKeyUp(keyCode As Integer, ByVal Shift As Integer)
  AddLogEntry "ExTVwU_EditKeyUp: keyCode=" & keyCode & ", shift=" & Shift
End Sub

Private Sub ExTVwU_EditLostFocus()
  AddLogEntry "ExTVwU_EditLostFocus"
End Sub

Private Sub ExTVwU_EditMClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwU_EditMClick: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwU_EditMDblClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwU_EditMDblClick: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwU_EditMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwU_EditMouseDown: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwU_EditMouseEnter(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwU_EditMouseEnter: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwU_EditMouseHover(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwU_EditMouseHover: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwU_EditMouseLeave(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwU_EditMouseLeave: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwU_EditMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwU_EditMouseMove: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwU_EditMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwU_EditMouseUp: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwU_EditMouseWheel(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal scrollAxis As ExTVwLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  AddLogEntry "ExTVwU_EditMouseWheel: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta
End Sub

Private Sub ExTVwU_EditRClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwU_EditRClick: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwU_EditRDblClick(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  AddLogEntry "ExTVwU_EditRDblClick: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub ExTVwU_EditXClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExTVwU_EditXClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExTVwU_EditXDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "ExTVwU_EditXDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub ExTVwU_ExpandedItem(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal expandedPartially As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ExpandedItem: treeItem=Nothing, expandedPartially=" & expandedPartially
  Else
    AddLogEntry "ExTVwU_ExpandedItem: treeItem=" & treeItem.Text & ", expandedPartially=" & expandedPartially
  End If
End Sub

Private Sub ExTVwU_ExpandingItem(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal expandingPartially As Boolean, cancelExpansion As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ExpandingItem: treeItem=Nothing, expandingPartially=" & expandingPartially & ", cancelExpansion=" & cancelExpansion
  Else
    AddLogEntry "ExTVwU_ExpandingItem: treeItem=" & treeItem.Text & ", expandingPartially=" & expandingPartially & ", cancelExpansion=" & cancelExpansion
  End If
End Sub

Private Sub ExTVwU_FreeItemData(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_FreeItemData: treeItem=Nothing"
  Else
    AddLogEntry "ExTVwU_FreeItemData: treeItem=" & treeItem.Text
  End If
End Sub

Private Sub ExTVwU_GotFocus()
  AddLogEntry "ExTVwU_GotFocus"
  Set objActiveCtl = ExTVwU
End Sub

Private Sub ExTVwU_IncrementalSearchStringChanging(ByVal currentSearchString As String, ByVal keyCodeOfCharToBeAdded As Integer, cancelChange As Boolean)
  AddLogEntry "ExTVwU_IncrementalSearchStringChanging: currentSearchString=" & currentSearchString & ", keyCodeOfCharToBeAdded=" & keyCodeOfCharToBeAdded & ", cancelChange=" & cancelChange
End Sub

Private Sub ExTVwU_InsertedItem(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_InsertedItem: treeItem=Nothing"
  Else
    AddLogEntry "ExTVwU_InsertedItem: treeItem=" & treeItem.Text
  End If
End Sub

Private Sub ExTVwU_InsertingItem(ByVal treeItem As ExTVwLibUCtl.IVirtualTreeViewItem, cancelInsertion As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_InsertingItem: treeItem=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ExTVwU_InsertingItem: treeItem=" & treeItem.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ExTVwU_ItemAsynchronousDrawFailed(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, imageDetails As ExTVwLibUCtl.FAILEDIMAGEDETAILS, ByVal failureReason As ExTVwLibUCtl.ImageDrawingFailureReasonConstants, furtherProcessing As ExTVwLibUCtl.FailedAsyncDrawReturnValuesConstants, newImageToDraw As Long)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ItemAsynchronousDrawFailed: treeItem=Nothing, failureReason=" & failureReason & ", furtherProcessing=" & furtherProcessing & ", newImageToDraw=" & newImageToDraw
  Else
    AddLogEntry "ExTVwU_ItemAsynchronousDrawFailed: treeItem=" & treeItem.Text & ", failureReason=" & failureReason & ", furtherProcessing=" & furtherProcessing & ", newImageToDraw=" & newImageToDraw
  End If
  AddLogEntry "     imageDetails.hImageList=0x" & Hex(imageDetails.hImageList)
  AddLogEntry "     imageDetails.hDC=0x" & Hex(imageDetails.hDC)
  AddLogEntry "     imageDetails.IconIndex=" & imageDetails.IconIndex
  AddLogEntry "     imageDetails.OverlayIndex=" & imageDetails.OverlayIndex
  AddLogEntry "     imageDetails.DrawingStyle=" & imageDetails.DrawingStyle
  AddLogEntry "     imageDetails.DrawingEffects=" & imageDetails.DrawingEffects
  AddLogEntry "     imageDetails.BackColor=0x" & Hex(imageDetails.BackColor)
  AddLogEntry "     imageDetails.ForeColor=0x" & Hex(imageDetails.ForeColor)
  AddLogEntry "     imageDetails.EffectColor=0x" & Hex(imageDetails.EffectColor)
End Sub

Private Sub ExTVwU_ItemBeginDrag(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ItemBeginDrag: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_ItemBeginDrag: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If

  On Error Resume Next
  ExTVwU.BeginDrag ExTVwU.CreateItemContainer(treeItem), -1
End Sub

Private Sub ExTVwU_ItemBeginRDrag(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ItemBeginRDrag: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_ItemBeginRDrag: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwU_ItemGetDisplayInfo(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal requestedInfo As ExTVwLibUCtl.RequestedInfoConstants, IconIndex As Long, SelectedIconIndex As Long, ExpandedIconIndex As Long, HasExpando As Boolean, ByVal maxItemTextLength As Long, itemText As String, dontAskAgain As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ItemGetDisplayInfo: treeItem=Nothing, requestedInfo=" & requestedInfo & ", iconIndex=" & IconIndex & ", selectedIconIndex=" & SelectedIconIndex & ", expandedIconIndex=" & ExpandedIconIndex & ", hasExpando=" & HasExpando & ", maxItemTextLength=" & maxItemTextLength & ", itemText=" & itemText & ", dontAskAgain=" & dontAskAgain
  Else
    AddLogEntry "ExTVwU_ItemGetDisplayInfo: treeItem=0x" & Hex(treeItem.Handle) & ", requestedInfo=" & requestedInfo & ", iconIndex=" & IconIndex & ", selectedIconIndex=" & SelectedIconIndex & ", expandedIconIndex=" & ExpandedIconIndex & ", hasExpando=" & HasExpando & ", maxItemTextLength=" & maxItemTextLength & ", itemText=" & itemText & ", dontAskAgain=" & dontAskAgain
  End If
End Sub

Private Sub ExTVwU_ItemGetInfoTipText(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ItemGetInfoTipText: treeItem=Nothing, maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
  Else
    AddLogEntry "ExTVwU_ItemGetInfoTipText: treeItem=" & treeItem.Text & ", maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
    If infoTipText <> "" Then
      infoTipText = infoTipText & vbNewLine
    End If
    infoTipText = infoTipText & "ID: " & treeItem.ID & vbNewLine & "ItemData: " & treeItem.ItemData
  End If
End Sub

Private Sub ExTVwU_ItemMouseEnter(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ItemMouseEnter: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_ItemMouseEnter: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwU_ItemMouseLeave(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ItemMouseLeave: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_ItemMouseLeave: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwU_ItemSelectionChanged(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ItemSelectionChanged: treeItem=Nothing"
  Else
    AddLogEntry "ExTVwU_ItemSelectionChanged: treeItem=" & treeItem.Text
  End If
End Sub

Private Sub ExTVwU_ItemSelectionChanging(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, cancelChange As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ItemSelectionChanging: treeItem=Nothing, cancelChange=" & cancelChange
  Else
    AddLogEntry "ExTVwU_ItemSelectionChanging: treeItem=" & treeItem.Text & ", cancelChange=" & cancelChange
  End If
End Sub

Private Sub ExTVwU_ItemSetText(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal itemText As String)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ItemSetText: treeItem=Nothing, itemText=" & itemText
  Else
    AddLogEntry "ExTVwU_ItemSetText: treeItem=0x" & Hex(treeItem.Handle) & ", itemText=" & itemText
  End If
End Sub

Private Sub ExTVwU_ItemStateImageChanged(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal previousStateImageIndex As Long, ByVal newStateImageIndex As Long, ByVal causedBy As ExTVwLibUCtl.StateImageChangeCausedByConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ItemStateImageChanged: treeItem=Nothing, previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy
  Else
    AddLogEntry "ExTVwU_ItemStateImageChanged: treeItem=" & treeItem.Text & ", previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy
  End If
End Sub

Private Sub ExTVwU_ItemStateImageChanging(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal previousStateImageIndex As Long, newStateImageIndex As Long, ByVal causedBy As ExTVwLibUCtl.StateImageChangeCausedByConstants, cancelChange As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_ItemStateImageChanging: treeItem=Nothing, previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy & ", cancelChange=" & cancelChange
  Else
    AddLogEntry "ExTVwU_ItemStateImageChanging: treeItem=" & treeItem.Text & ", previousStateImageIndex=" & previousStateImageIndex & ", newStateImageIndex=" & newStateImageIndex & ", causedBy=" & causedBy & ", cancelChange=" & cancelChange
  End If
End Sub

Private Sub ExTVwU_KeyDown(keyCode As Integer, ByVal Shift As Integer)
  AddLogEntry "ExTVwU_KeyDown: keyCode=" & keyCode & ", shift=" & Shift
  If keyCode = KeyCodeConstants.vbKeyF2 Then
    ExTVwU.CaretItem.StartLabelEditing
  End If
  If keyCode = KeyCodeConstants.vbKeyEscape Then
    If Not (ExTVwU.DraggedItems Is Nothing) Then ExTVwU.EndDrag True
  End If
End Sub

Private Sub ExTVwU_KeyPress(keyAscii As Integer)
  AddLogEntry "ExTVwU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ExTVwU_KeyUp(keyCode As Integer, ByVal Shift As Integer)
  AddLogEntry "ExTVwU_KeyUp: keyCode=" & keyCode & ", shift=" & Shift
End Sub

Private Sub ExTVwU_LostFocus()
  AddLogEntry "ExTVwU_LostFocus"
End Sub

Private Sub ExTVwU_MClick(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_MClick: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_MClick: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwU_MDblClick(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_MDblClick: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_MDblClick: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwU_MouseDown(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_MouseDown: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_MouseDown: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwU_MouseEnter(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_MouseEnter: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_MouseEnter: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwU_MouseHover(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_MouseHover: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_MouseHover: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwU_MouseLeave(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_MouseLeave: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_MouseLeave: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwU_MouseMove(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
'  If treeItem Is Nothing Then
'    AddLogEntry "ExTVwU_MouseMove: treeItem=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
'  Else
'    AddLogEntry "ExTVwU_MouseMove: treeItem=" & treeItem.Text & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails
'  End If
End Sub

Private Sub ExTVwU_MouseUp(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_MouseUp: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_MouseUp: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If

  If Button = MouseButtonConstants.vbLeftButton Then
    If Not (ExTVwU.DraggedItems Is Nothing) Then
      ' Are we within the client area?
      If ((HitTestConstants.htAbove Or HitTestConstants.htBelow Or HitTestConstants.htToLeft Or HitTestConstants.htToRight) And hitTestDetails) = 0 Then
        ' yes
        ExTVwU.EndDrag False
      Else
        ' no
        ExTVwU.EndDrag True
      End If
    End If
  End If
End Sub

Private Sub ExTVwU_MouseWheel(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants, ByVal scrollAxis As ExTVwLibUCtl.ScrollAxisConstants, ByVal wheelDelta As Integer)
  Dim str As String

  str = "ExTVwU_MouseWheel: treeItem="
  If treeItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & treeItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", scrollAxis=" & scrollAxis & ", wheelDelta=" & wheelDelta

  AddLogEntry str
End Sub

Private Sub ExTVwU_OLECompleteDrag(ByVal data As ExTVwLibUCtl.IOLEDataObject, ByVal performedEffect As ExTVwLibUCtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "ExTVwU_OLECompleteDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub ExTVwU_OLEDragDrop(ByVal data As ExTVwLibUCtl.IOLEDataObject, effect As ExTVwLibUCtl.OLEDropEffectConstants, dropTarget As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ExTVwU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ExTVwU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ExTVwU_OLEDragEnter(ByVal data As ExTVwLibUCtl.IOLEDataObject, effect As ExTVwLibUCtl.OLEDropEffectConstants, dropTarget As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants, autoExpandItem As Boolean, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ExTVwU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails & ", autoExpandItem=" & autoExpandItem & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ExTVwU_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "ExTVwU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub ExTVwU_OLEDragLeave(ByVal data As ExTVwLibUCtl.IOLEDataObject, ByVal dropTarget As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ExTVwU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub ExTVwU_OLEDragLeavePotentialTarget()
  AddLogEntry "ExTVwU_OLEDragLeavePotentialTarget"
End Sub

Private Sub ExTVwU_OLEDragMouseMove(ByVal data As ExTVwLibUCtl.IOLEDataObject, effect As ExTVwLibUCtl.OLEDropEffectConstants, dropTarget As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants, autoExpandItem As Boolean, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim files() As String
  Dim str As String

  str = "ExTVwU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", yToItemTop=" & yToItemTop & ", hitTestDetails=" & hitTestDetails & ", autoExpandItem=" & autoExpandItem & ", autoHScrollVelocity=" & autoHScrollVelocity & ", autoVScrollVelocity=" & autoVScrollVelocity

  AddLogEntry str
End Sub

Private Sub ExTVwU_OLEGiveFeedback(ByVal effect As ExTVwLibUCtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "ExTVwU_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub ExTVwU_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal Button As Integer, ByVal Shift As Integer, actionToContinueWith As ExTVwLibUCtl.OLEActionToContinueWithConstants)
  AddLogEntry "ExTVwU_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & Button & ", shift=" & Shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub ExTVwU_OLEReceivedNewData(ByVal data As ExTVwLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ExTVwU_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ExTVwU_OLESetData(ByVal data As ExTVwLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ExTVwU_OLESetData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ExTVwU_OLEStartDrag(ByVal data As ExTVwLibUCtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "ExTVwU_OLEStartDrag: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub ExTVwU_RClick(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_RClick: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_RClick: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwU_RDblClick(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_RDblClick: treeItem=Nothing, button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  Else
    AddLogEntry "ExTVwU_RDblClick: treeItem=" & treeItem.Text & ", button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=" & hitTestDetails
  End If
End Sub

Private Sub ExTVwU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ExTVwU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertItemsU
End Sub

Private Sub ExTVwU_RemovedItem(ByVal treeItem As ExTVwLibUCtl.IVirtualTreeViewItem)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_RemovedItem: treeItem=Nothing"
  Else
    AddLogEntry "ExTVwU_RemovedItem: treeItem=" & treeItem.Text
  End If
End Sub

Private Sub ExTVwU_RemovingItem(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, cancelDeletion As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_RemovingItem: treeItem=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ExTVwU_RemovingItem: treeItem=" & treeItem.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ExTVwU_RenamedItem(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal previousItemText As String, ByVal newItemText As String)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_RenamedItem: treeItem=Nothing, previousItemText=" & previousItemText & ", newItemText=" & newItemText
  Else
    AddLogEntry "ExTVwU_RenamedItem: treeItem=" & treeItem.Text & ", previousItemText=" & previousItemText & ", newItemText=" & newItemText
  End If
End Sub

Private Sub ExTVwU_RenamingItem(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal previousItemText As String, ByVal newItemText As String, cancelRenaming As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_RenamingItem: treeItem=Nothing, previousItemText=" & previousItemText & ", newItemText=" & newItemText & ", cancelRenaming=" & cancelRenaming
  Else
    AddLogEntry "ExTVwU_RenamingItem: treeItem=" & treeItem.Text & ", previousItemText=" & previousItemText & ", newItemText=" & newItemText & ", cancelRenaming=" & cancelRenaming
  End If
End Sub

Private Sub ExTVwU_ResizedControlWindow()
  AddLogEntry "ExTVwU_ResizedControlWindow"
End Sub

Private Sub ExTVwU_SingleExpanding(ByVal previousCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal newCaretItem As ExTVwLibUCtl.ITreeViewItem, dontChangePreviousItem As Boolean, dontChangeNewItem As Boolean)
  Dim str As String

  str = "ExTVwU_SingleExpanding: previousCaretItem="
  If previousCaretItem Is Nothing Then
    str = str & "Nothing, newCaretItem="
  Else
    str = str & previousCaretItem.Text & ", newCaretItem="
  End If
  If newCaretItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newCaretItem.Text
  End If
  str = str & ", dontChangePreviousItem=" & dontChangePreviousItem & ", dontChangeNewItem=" & dontChangeNewItem

  AddLogEntry str
End Sub

Private Sub ExTVwU_StartingLabelEditing(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, cancelEditing As Boolean)
  If treeItem Is Nothing Then
    AddLogEntry "ExTVwU_StartingLabelEditing: treeItem=Nothing, cancelEditing=" & cancelEditing
  Else
    AddLogEntry "ExTVwU_StartingLabelEditing: treeItem=" & treeItem.Text & ", cancelEditing=" & cancelEditing
  End If
End Sub

Private Sub ExTVwU_XClick(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExTVwU_XClick: treeItem="
  If treeItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & treeItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExTVwU_XDblClick(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  Dim str As String

  str = "ExTVwU_XDblClick: treeItem="
  If treeItem Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & treeItem.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub ExTVwU_Validate(Cancel As Boolean)
  AddLogEntry "ExTVwU_Validate"
End Sub

Private Sub Form_Click()
  '
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
  chkLog.Value = CheckBoxConstants.vbChecked

  ' this is required to make the control work as expected
  Subclass

  InsertItemsA
  InsertItemsU
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    cmdAbout.Left = Me.ScaleWidth / 2 - cmdAbout.Width / 2
    cmdAbout.Top = Me.ScaleHeight - cmdAbout.Height - 8
    txtLog.Width = Me.ScaleWidth - txtLog.Left - ExTVwA.Width - 8
    txtLog.Height = Me.ScaleHeight - txtLog.Top - cmdAbout.Height - 16
    chkLog.Top = cmdAbout.Top + 8
  End If
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long
  Static cLines As Long
  Static oldtxt As String

  If bLog Then
    cLines = cLines + 1
    If cLines > 50 Then
      ' delete the first line
      pos = InStr(oldtxt, vbNewLine)
      oldtxt = Mid(oldtxt, pos + Len(vbNewLine))
      cLines = cLines - 1
    End If
    oldtxt = oldtxt & (txt & vbNewLine)

    txtLog.Text = oldtxt
    txtLog.SelStart = Len(oldtxt)
  End If
End Sub

Private Sub InsertItemsA()
  Dim cImages As Long
  Dim iIcon As Long
  Dim itm1 As ExTVwLibACtl.TreeViewItem
  Dim itm2 As ExTVwLibACtl.TreeViewItem

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExTVwA.hWnd, StrPtr("explorer"), 0
  End If

  ExTVwA.hImageList(ImageListConstants.ilItems) = hImgLst
  cImages = ImageList_GetImageCount(hImgLst)

  With ExTVwA.TreeItems
    Set itm1 = .Add("Level 1, Item 1", Nothing, , heYes, iIcon)
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 2, Item 1", itm1, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    itm1.Expand
    Set itm2 = .Add("Level 2, Item 2", itm1, , heYes, iIcon)
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 3, Item 1", itm2, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 3, Item 2", itm2, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 3, Item 3", itm2, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    itm2.Expand
    .Add "Level 2, Item 3", itm1, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 2, Item 4", itm1, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    Set itm2 = .Add("Level 2, Item 5", itm1, , heYes, iIcon)
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 3, Item 4", itm2, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 3, Item 5", itm2, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    itm2.Expand

    Set itm1 = .Add("Level 1, Item 1", Nothing, , heYes, iIcon)
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 2, Item 1", itm1, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    itm1.Expand
    Set itm2 = .Add("Level 2, Item 2", itm1, , heYes, iIcon)
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 3, Item 1", itm2, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 3, Item 2", itm2, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 3, Item 3", itm2, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    itm2.Expand
    .Add "Level 2, Item 3", itm1, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 2, Item 4", itm1, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    Set itm2 = .Add("Level 2, Item 5", itm1, , heYes, iIcon)
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 3, Item 4", itm2, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    .Add "Level 3, Item 5", itm2, , heNo, iIcon
    iIcon = (iIcon + 1) Mod cImages
    itm2.Expand
  End With
End Sub

Private Sub InsertItemsU()
  Dim cImages As Long
  Dim iIcon As Long
  Dim itm1 As ExTVwLibUCtl.TreeViewItem
  Dim itm2 As ExTVwLibUCtl.TreeViewItem

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ExTVwU.hWnd, StrPtr("explorer"), 0
  End If

  ExTVwU.hImageList(ImageListConstants.ilItems) = hImgLst
  cImages = ImageList_GetImageCount(hImgLst)

  With ExTVwU.TreeItems
    Set itm1 = .Add("Level 1, Item 1", , , heYes, iIcon)
    iIcon = (iIcon + 1) Mod cImages
    With itm1.SubItems
      .Add "Level 2, Item 1", , , heNo, iIcon
      iIcon = (iIcon + 1) Mod cImages
      itm1.Expand
      Set itm2 = .Add("Level 2, Item 2", , , heYes, iIcon)
      iIcon = (iIcon + 1) Mod cImages
      With itm2.SubItems
        .Add "Level 3, Item 1", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        .Add "Level 3, Item 2", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        .Add "Level 3, Item 3", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        itm2.Expand
      End With
      .Add "Level 2, Item 3", , , heNo, iIcon
      iIcon = (iIcon + 1) Mod cImages
      .Add "Level 2, Item 4", , , heNo, iIcon
      iIcon = (iIcon + 1) Mod cImages
      Set itm2 = .Add("Level 2, Item 5", , , heYes, iIcon)
      iIcon = (iIcon + 1) Mod cImages
      With itm2.SubItems
        .Add "Level 3, Item 4", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        .Add "Level 3, Item 5", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        itm2.Expand
      End With
    End With

    Set itm1 = .Add("Level 1, Item 1", , , heYes, iIcon)
    iIcon = (iIcon + 1) Mod cImages
    With itm1.SubItems
      .Add "Level 2, Item 1", , , heNo, iIcon
      iIcon = (iIcon + 1) Mod cImages
      itm1.Expand
      Set itm2 = .Add("Level 2, Item 2", , , heYes, iIcon)
      iIcon = (iIcon + 1) Mod cImages
      With itm2.SubItems
        .Add "Level 3, Item 1", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        .Add "Level 3, Item 2", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        .Add "Level 3, Item 3", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        itm2.Expand
      End With
      .Add "Level 2, Item 3", , , heNo, iIcon
      iIcon = (iIcon + 1) Mod cImages
      .Add "Level 2, Item 4", , , heNo, iIcon
      iIcon = (iIcon + 1) Mod cImages
      Set itm2 = .Add("Level 2, Item 5", , , heYes, iIcon)
      iIcon = (iIcon + 1) Mod cImages
      With itm2.SubItems
        .Add "Level 3, Item 4", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        .Add "Level 3, Item 5", , , heNo, iIcon
        iIcon = (iIcon + 1) Mod cImages
        itm2.Expand
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
    SendMessageAsLong ExTVwU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong ExTVwA.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
