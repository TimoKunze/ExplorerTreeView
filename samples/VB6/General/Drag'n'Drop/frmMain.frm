VERSION 5.00
Object = "{1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5}#2.6#0"; "ExTVwU.ocx"
Begin VB.Form frmMain 
   Caption         =   "ExplorerTreeView 2.6 - Drag'n'Drop Sample"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5145
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
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   343
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkOLEDragDrop 
      Caption         =   "&OLE Drag'n'Drop"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   3480
      Value           =   1  'Aktiviert
      Width           =   1575
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
      Left            =   3600
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin ExTVwLibUCtl.ExplorerTreeView ExTVwU 
      Align           =   3  'Links ausrichten
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _cx             =   5953
      _cy             =   7011
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
      DisabledEvents  =   264182
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
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuAeroDragImages 
         Caption         =   "&Aero OLE Drag Images"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "dummy"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy Item"
      End
      Begin VB.Menu mnuMove 
         Caption         =   "&Move Item"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "C&ancel"
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


  Private Enum ChosenMenuItemConstants
    ciCopy
    ciMove
    ciCancel
  End Enum


  Private Const CF_DIBV5 = 17
  Private Const CF_OEMTEXT = 7
  Private Const CF_UNICODETEXT = 13

  Private Const strCLSID_RecycleBin = "{645FF040-5081-101B-9F08-00AA002F954E}"


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
  End Type

  Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
  End Type


  Private CFSTR_HTML As Long
  Private CFSTR_HTML2 As Long
  Private CFSTR_LOGICALPERFORMEDDROPEFFECT As Long
  Private CFSTR_MOZILLAHTMLCONTEXT As Long
  Private CFSTR_PERFORMEDDROPEFFECT As Long
  Private CFSTR_TARGETCLSID As Long
  Private CFSTR_TEXTHTML As Long

  Private bComctl32Version600OrNewer As Boolean
  Private bRightDrag As Boolean
  Private chosenMenuItem As ChosenMenuItemConstants
  Private CLSID_RecycleBin As GUID
  Private draggedPicture As StdPicture
  Private hImgLst As Long
  Private hLastDropTarget As OLE_HANDLE
  Private ToolTip As clsToolTip
  Private usingThemes As Boolean


  Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal pString As Long, CLSID As GUID) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hinst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function RegisterClipboardFormat Lib "user32.dll" Alias "RegisterClipboardFormatW" (ByVal lpString As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long
  Private Declare Function SHGetFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal CSIDL As Long, ByVal hToken As Long, ByVal reserved As Long, pIDLToDesktop As Long) As Long


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
  ExTVwU.About
End Sub

Private Sub ExTVwU_AbortedDrag()
  Set ExTVwU.DropHilitedItem = Nothing
  ExTVwU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, Nothing
  hLastDropTarget = 0
End Sub

Private Sub ExTVwU_DragMouseMove(dropTarget As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants, autoExpandItem As Boolean, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim hNewDropTarget As OLE_HANDLE
  Dim lastInsertMarkRelativePosition As InsertMarkPositionConstants
  Dim newInsertMarkRelativePosition As InsertMarkPositionConstants
  Dim treeItem As TreeViewItem

  If dropTarget Is Nothing Then
    hNewDropTarget = 0
  Else
    hNewDropTarget = dropTarget.Handle
  End If
  If hNewDropTarget <> hLastDropTarget Then
    Set ExTVwU.DropHilitedItem = dropTarget
    hLastDropTarget = hNewDropTarget
  End If

  newInsertMarkRelativePosition = CalculateInsertMarkPosition(dropTarget, yToItemTop)
  ExTVwU.GetInsertMarkPosition lastInsertMarkRelativePosition, treeItem
  If (newInsertMarkRelativePosition <> lastInsertMarkRelativePosition) Or Not IsSameItem(treeItem, dropTarget) Then
    ExTVwU.SetInsertMarkPosition newInsertMarkRelativePosition, dropTarget
  End If
End Sub

Private Sub ExTVwU_Drop(ByVal dropTarget As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  Dim insertAfter As TreeViewItem
  Dim insertionMark As InsertMarkPositionConstants
  Dim itm As TreeViewItem
  Dim parentItem As TreeViewItem

  If bRightDrag Then
    x = Me.ScaleX(x, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
    y = Me.ScaleY(y, ScaleModeConstants.vbTwips, ScaleModeConstants.vbPixels)
    Me.PopupMenu mnuPopup, , x, y, mnuMove
    Select Case chosenMenuItem
      Case ChosenMenuItemConstants.ciCancel
        ExTVwU_AbortedDrag
        Exit Sub
      Case ChosenMenuItemConstants.ciCopy
        MsgBox "ToDo: Copy the item.", VbMsgBoxStyle.vbInformation, "Command not implemented"
        ExTVwU_AbortedDrag
        Exit Sub
      Case ChosenMenuItemConstants.ciMove
        ' fall through
    End Select
  End If

  ' move the item

  ExTVwU.GetInsertMarkPosition insertionMark, itm
  Select Case insertionMark
    Case InsertMarkPositionConstants.impAfter
      Set parentItem = dropTarget.parentItem
      Set insertAfter = dropTarget
    Case InsertMarkPositionConstants.impBefore
      Set parentItem = dropTarget.parentItem
      Set insertAfter = dropTarget.PreviousSiblingItem
    Case InsertMarkPositionConstants.impNowhere
      Set parentItem = dropTarget
  End Select

  If Not (parentItem Is Nothing) Then
    parentItem.HasExpando = HasExpandoConstants.heYes
  End If
  For Each itm In ExTVwU.DraggedItems
    ' If we move an item together with its sub-items, it may happen that we move the parent before
    ' the children, which will raise an error when the children shall be moved. It's the easiest
    ' solution to ignore this error.
    ' The same goes for the situation that the 'DraggedItems' object contains items that can't be
    ' moved to the target (e. g. because they're parent items of the target).
    On Error Resume Next
    If insertionMark = InsertMarkPositionConstants.impNowhere Then
      itm.Move parentItem, InsertAfterConstants.iaLast
    Else
      itm.Move parentItem, insertAfter
    End If
  Next itm

  Set ExTVwU.DropHilitedItem = Nothing
  ExTVwU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, Nothing
  hLastDropTarget = 0
End Sub

Private Sub ExTVwU_ItemBeginDrag(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  bRightDrag = False
  hLastDropTarget = 0
  If chkOLEDragDrop.Value = CheckBoxConstants.vbChecked Then
    ExTVwU.OLEDrag , , , ExTVwU.CreateItemContainer(treeItem)
  Else
    ExTVwU.BeginDrag ExTVwU.CreateItemContainer(treeItem), -1
  End If
End Sub

Private Sub ExTVwU_ItemBeginRDrag(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  bRightDrag = True
  hLastDropTarget = 0
  If chkOLEDragDrop.Value = CheckBoxConstants.vbChecked Then
    ExTVwU.OLEDrag , , , ExTVwU.CreateItemContainer(treeItem)
  Else
    ExTVwU.BeginDrag ExTVwU.CreateItemContainer(treeItem), -1
  End If
End Sub

Private Sub ExTVwU_KeyDown(keyCode As Integer, ByVal Shift As Integer)
  If keyCode = KeyCodeConstants.vbKeyEscape Then
    If Not (ExTVwU.DraggedItems Is Nothing) Then ExTVwU.EndDrag True
  End If
End Sub

Private Sub ExTVwU_MouseUp(ByVal treeItem As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  If chkOLEDragDrop.Value = CheckBoxConstants.vbUnchecked Then
    If ((Button = MouseButtonConstants.vbLeftButton) And (Not bRightDrag)) Or ((Button = MouseButtonConstants.vbRightButton) And bRightDrag) Then
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
  End If
End Sub

Private Sub ExTVwU_OLECompleteDrag(ByVal Data As ExTVwLibUCtl.IOLEDataObject, ByVal performedEffect As ExTVwLibUCtl.OLEDropEffectConstants)
  Dim bytArray() As Byte
  Dim CLSIDOfTarget As GUID

  If Data.GetFormat(CFSTR_TARGETCLSID) Then
    bytArray = Data.GetData(CFSTR_TARGETCLSID)
    CopyMemory VarPtr(CLSIDOfTarget), VarPtr(bytArray(LBound(bytArray))), LenB(CLSIDOfTarget)
    If IsEqualGUID(CLSIDOfTarget, CLSID_RecycleBin) Then
      ' remove the dragged items
      ExTVwU.DraggedItems.RemoveAll True
    End If
  End If
End Sub

Private Sub ExTVwU_OLEDragDrop(ByVal Data As ExTVwLibUCtl.IOLEDataObject, Effect As ExTVwLibUCtl.OLEDropEffectConstants, dropTarget As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  Dim bytArray() As Byte
  Dim files() As String
  Dim i As Long
  Dim insertAfter As Variant
  Dim insertionMark As InsertMarkPositionConstants
  Dim itm As TreeViewItem
  Dim parentItem As TreeViewItem
  Dim str As String

  If Data.GetFormat(vbCFDIB) Then
    Set draggedPicture = Data.GetData(vbCFDIB)
  Else
    If Data.GetFormat(CF_UNICODETEXT) Then
      str = Data.GetData(CF_UNICODETEXT)
    ElseIf Data.GetFormat(vbCFText) Then
      str = Data.GetData(vbCFText)
    ElseIf Data.GetFormat(CF_OEMTEXT) Then
      str = Data.GetData(CF_OEMTEXT)
    End If
    str = Replace$(str, vbNewLine, "")

    If str <> "" Then
      ' insert a new item

      ExTVwU.GetInsertMarkPosition insertionMark, itm
      Select Case insertionMark
        Case InsertMarkPositionConstants.impAfter
          Set parentItem = dropTarget.parentItem
          Set insertAfter = dropTarget
        Case InsertMarkPositionConstants.impBefore
          Set parentItem = dropTarget.parentItem
          Set insertAfter = dropTarget.PreviousSiblingItem
        Case InsertMarkPositionConstants.impNowhere
          Set parentItem = dropTarget
      End Select

      If Not (parentItem Is Nothing) Then
        parentItem.HasExpando = HasExpandoConstants.heYes
      End If
      If insertionMark = InsertMarkPositionConstants.impNowhere Then insertAfter = InsertAfterConstants.iaLast

      ExTVwU.TreeItems.Add str, parentItem, insertAfter, , 1, 1

    ElseIf Data.GetFormat(vbCFFiles) Then
      ' insert a new item for each file/folder

      ExTVwU.GetInsertMarkPosition insertionMark, itm
      Select Case insertionMark
        Case InsertMarkPositionConstants.impAfter
          Set parentItem = dropTarget.parentItem
          Set insertAfter = dropTarget
        Case InsertMarkPositionConstants.impBefore
          Set parentItem = dropTarget.parentItem
          Set insertAfter = dropTarget.PreviousSiblingItem
        Case InsertMarkPositionConstants.impNowhere
          Set parentItem = dropTarget
      End Select

      If Not (parentItem Is Nothing) Then
        parentItem.HasExpando = HasExpandoConstants.heYes
      End If
      If insertionMark = InsertMarkPositionConstants.impNowhere Then insertAfter = InsertAfterConstants.iaLast

      On Error GoTo NoFiles
      files = Data.GetData(vbCFFiles)
      For i = LBound(files) To UBound(files)
        Set insertAfter = ExTVwU.TreeItems.Add(files(i), parentItem, insertAfter, , 1, 1)
      Next i
NoFiles:
    End If
  End If

  Set ExTVwU.DropHilitedItem = Nothing
  ExTVwU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, Nothing
  hLastDropTarget = 0

  ' Be careful with odeMove!! Some sources delete the data themselves if the Move effect was
  ' chosen. So you may lose data!
  'If shift And vbShiftMask Then effect = ExTVwLibUCtl.OLEDropEffectConstants.odeMove
  If Shift And vbCtrlMask Then Effect = ExTVwLibUCtl.OLEDropEffectConstants.odeCopy
  If Shift And vbAltMask Then Effect = ExTVwLibUCtl.OLEDropEffectConstants.odeLink

'  If Data.GetFormat(vbCFText) Then MsgBox "Dropped Text: " & Data.GetData(vbCFText)
'  If Data.GetFormat(CF_OEMTEXT) Then MsgBox "Dropped OEM-Text: " & Data.GetData(7)
'  If Data.GetFormat(CF_UNICODETEXT) Then MsgBox "Dropped Unicode-Text: " & Data.GetData(13)
'  If Data.GetFormat(vbCFRTF) Then MsgBox "Dropped RTF-Text: " & Data.GetData(vbCFRTF)
'  If Data.GetFormat(CFSTR_HTML) Then MsgBox "Dropped HTML: " & Data.GetData(CFSTR_HTML)
'  If Data.GetFormat(CFSTR_HTML2) Then MsgBox "Dropped HTML2: " & Data.GetData(CFSTR_HTML2)
'  If Data.GetFormat(CFSTR_MOZILLAHTMLCONTEXT) Then MsgBox "Dropped text/_moz_htmlcontext: " & Data.GetData(CFSTR_MOZILLAHTMLCONTEXT)
'  If Data.GetFormat(CFSTR_TEXTHTML) Then MsgBox "Dropped text/html: " & Data.GetData(CFSTR_TEXTHTML)
'  If data.GetFormat(vbCFFiles) Then
'    str = ""
'    For Each file In data.Files
'      str = str & file & vbNewLine
'    Next file
'    MsgBox "Dropped files:" & vbNewLine & str
'  End If
'
'  ReDim bytArray(1 To LenB(CLSID_RecycleBin)) As Byte
'  CopyMemory VarPtr(bytArray(LBound(bytArray))), VarPtr(CLSID_RecycleBin), LenB(CLSID_RecycleBin)
'  Data.SetData CFSTR_TARGETCLSID, bytArray
'
'  ' another way to inform the source of the performed drop effect
'  Data.SetData CFSTR_PERFORMEDDROPEFFECT, Effect
End Sub

Private Sub ExTVwU_OLEDragEnter(ByVal Data As ExTVwLibUCtl.IOLEDataObject, Effect As ExTVwLibUCtl.OLEDropEffectConstants, dropTarget As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants, autoExpandItem As Boolean, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim dropActionDescription As String
  Dim dropDescriptionIcon As DropDescriptionIconConstants
  Dim dropTargetDescription As String
  Dim hNewDropTarget As OLE_HANDLE
  Dim lastInsertMarkRelativePosition As InsertMarkPositionConstants
  Dim newInsertMarkRelativePosition As InsertMarkPositionConstants
  Dim treeItem As TreeViewItem

  If dropTarget Is Nothing Then
    hNewDropTarget = 0
  Else
    dropTargetDescription = dropTarget.Text
    hNewDropTarget = dropTarget.Handle
  End If
  If hNewDropTarget <> hLastDropTarget Then
    Set ExTVwU.DropHilitedItem = dropTarget
    hLastDropTarget = hNewDropTarget
  End If

  newInsertMarkRelativePosition = CalculateInsertMarkPosition(dropTarget, yToItemTop)
  ExTVwU.GetInsertMarkPosition lastInsertMarkRelativePosition, treeItem
  If (newInsertMarkRelativePosition <> lastInsertMarkRelativePosition) Or Not IsSameItem(treeItem, dropTarget) Then
    ExTVwU.SetInsertMarkPosition newInsertMarkRelativePosition, dropTarget
  End If

  Effect = ExTVwLibUCtl.OLEDropEffectConstants.odeCopy
  If Shift And vbShiftMask Then Effect = ExTVwLibUCtl.OLEDropEffectConstants.odeMove
  If Shift And vbCtrlMask Then Effect = ExTVwLibUCtl.OLEDropEffectConstants.odeCopy
  If Shift And vbAltMask Then Effect = ExTVwLibUCtl.OLEDropEffectConstants.odeLink

  If dropTarget Is Nothing Then
    dropDescriptionIcon = DropDescriptionIconConstants.ddiNoDrop
    dropActionDescription = "Cannot place here"
  Else
    Select Case Effect
      Case ExTVwLibUCtl.OLEDropEffectConstants.odeMove
        Select Case newInsertMarkRelativePosition
          Case InsertMarkPositionConstants.impAfter
            dropActionDescription = "Move after %1"
          Case InsertMarkPositionConstants.impBefore
            dropActionDescription = "Move before %1"
          Case Else
            dropActionDescription = "Move to %1"
        End Select
        dropDescriptionIcon = DropDescriptionIconConstants.ddiMove
      Case ExTVwLibUCtl.OLEDropEffectConstants.odeLink
        Select Case newInsertMarkRelativePosition
          Case InsertMarkPositionConstants.impAfter
            dropActionDescription = "Create link after %1"
          Case InsertMarkPositionConstants.impBefore
            dropActionDescription = "Create link before %1"
          Case Else
            dropActionDescription = "Create link in %1"
        End Select
        dropDescriptionIcon = DropDescriptionIconConstants.ddiLink
      Case Else
        Select Case newInsertMarkRelativePosition
          Case InsertMarkPositionConstants.impAfter
            dropActionDescription = "Insert copy after %1"
          Case InsertMarkPositionConstants.impBefore
            dropActionDescription = "Insert copy before %1"
          Case Else
            dropActionDescription = "Copy to %1"
        End Select
        dropDescriptionIcon = DropDescriptionIconConstants.ddiCopy
    End Select
  End If
  Data.SetDropDescription dropTargetDescription, dropActionDescription, dropDescriptionIcon
End Sub

Private Sub ExTVwU_OLEDragLeave(ByVal Data As ExTVwLibUCtl.IOLEDataObject, ByVal dropTarget As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants)
  Set ExTVwU.DropHilitedItem = Nothing
  ExTVwU.SetInsertMarkPosition InsertMarkPositionConstants.impNowhere, Nothing
  hLastDropTarget = 0
End Sub

Private Sub ExTVwU_OLEDragMouseMove(ByVal Data As ExTVwLibUCtl.IOLEDataObject, Effect As ExTVwLibUCtl.OLEDropEffectConstants, dropTarget As ExTVwLibUCtl.ITreeViewItem, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal yToItemTop As Long, ByVal hitTestDetails As ExTVwLibUCtl.HitTestConstants, autoExpandItem As Boolean, autoHScrollVelocity As Long, autoVScrollVelocity As Long)
  Dim dropActionDescription As String
  Dim dropDescriptionIcon As DropDescriptionIconConstants
  Dim dropTargetDescription As String
  Dim hNewDropTarget As OLE_HANDLE
  Dim lastInsertMarkRelativePosition As InsertMarkPositionConstants
  Dim newInsertMarkRelativePosition As InsertMarkPositionConstants
  Dim treeItem As TreeViewItem

  If dropTarget Is Nothing Then
    hNewDropTarget = 0
  Else
    dropTargetDescription = dropTarget.Text
    hNewDropTarget = dropTarget.Handle
  End If
  If hNewDropTarget <> hLastDropTarget Then
    Set ExTVwU.DropHilitedItem = dropTarget
    hLastDropTarget = hNewDropTarget
  End If

  newInsertMarkRelativePosition = CalculateInsertMarkPosition(dropTarget, yToItemTop)
  ExTVwU.GetInsertMarkPosition lastInsertMarkRelativePosition, treeItem
  If (newInsertMarkRelativePosition <> lastInsertMarkRelativePosition) Or Not IsSameItem(treeItem, dropTarget) Then
    ExTVwU.SetInsertMarkPosition newInsertMarkRelativePosition, dropTarget
  End If

  Effect = ExTVwLibUCtl.OLEDropEffectConstants.odeCopy
  If Shift And vbShiftMask Then Effect = ExTVwLibUCtl.OLEDropEffectConstants.odeMove
  If Shift And vbCtrlMask Then Effect = ExTVwLibUCtl.OLEDropEffectConstants.odeCopy
  If Shift And vbAltMask Then Effect = ExTVwLibUCtl.OLEDropEffectConstants.odeLink

  If dropTarget Is Nothing Then
    dropDescriptionIcon = DropDescriptionIconConstants.ddiNoDrop
    dropActionDescription = "Cannot place here"
  Else
    Select Case Effect
      Case ExTVwLibUCtl.OLEDropEffectConstants.odeMove
        Select Case newInsertMarkRelativePosition
          Case InsertMarkPositionConstants.impAfter
            dropActionDescription = "Move after %1"
          Case InsertMarkPositionConstants.impBefore
            dropActionDescription = "Move before %1"
          Case Else
            dropActionDescription = "Move to %1"
        End Select
        dropDescriptionIcon = DropDescriptionIconConstants.ddiMove
      Case ExTVwLibUCtl.OLEDropEffectConstants.odeLink
        Select Case newInsertMarkRelativePosition
          Case InsertMarkPositionConstants.impAfter
            dropActionDescription = "Create link after %1"
          Case InsertMarkPositionConstants.impBefore
            dropActionDescription = "Create link before %1"
          Case Else
            dropActionDescription = "Create link in %1"
        End Select
        dropDescriptionIcon = DropDescriptionIconConstants.ddiLink
      Case Else
        Select Case newInsertMarkRelativePosition
          Case InsertMarkPositionConstants.impAfter
            dropActionDescription = "Insert copy after %1"
          Case InsertMarkPositionConstants.impBefore
            dropActionDescription = "Insert copy before %1"
          Case Else
            dropActionDescription = "Copy to %1"
        End Select
        dropDescriptionIcon = DropDescriptionIconConstants.ddiCopy
    End Select
  End If
  Data.SetDropDescription dropTargetDescription, dropActionDescription, dropDescriptionIcon
End Sub

Private Sub ExTVwU_OLESetData(ByVal Data As ExTVwLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim itm As TreeViewItem
  Dim str As String

  Select Case formatID
    Case vbCFText, CF_OEMTEXT, CF_UNICODETEXT
      For Each itm In ExTVwU.DraggedItems
        str = str & itm.Text & vbNewLine
      Next itm
      Data.SetData formatID, str

    Case vbCFFiles
      ' WARNING! Dragging files to objects like the recycle bin may result in deleted files and
      ' lost data!
'      ReDim files(1 To 2) As String
'      files(1) = "C:\Test1.txt"
'      files(2) = "C:\Test2.txt"
'      Data.SetData formatID, files

    Case vbCFBitmap, vbCFDIB, CF_DIBV5
      Data.SetData formatID, draggedPicture

    Case vbCFPalette
      If Not (draggedPicture Is Nothing) Then
        Data.SetData formatID, draggedPicture.hPal
      End If
  End Select
End Sub

Private Sub ExTVwU_OLEStartDrag(ByVal Data As ExTVwLibUCtl.IOLEDataObject)
  Data.SetData vbCFText
  Data.SetData CF_OEMTEXT
  Data.SetData CF_UNICODETEXT
  Data.SetData vbCFFiles
  If Not (draggedPicture Is Nothing) Then
    Data.SetData vbCFPalette
    Data.SetData vbCFBitmap
    Data.SetData vbCFDIB
    Data.SetData CF_DIBV5
  End If
End Sub

Private Sub ExTVwU_RecreatedControlWindow(ByVal hWndTreeView As Long)
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

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    FreeLibrary hMod
    usingThemes = bComctl32Version600OrNewer
  End If

  CFSTR_HTML = RegisterClipboardFormat(StrPtr("HTML Format"))
  CFSTR_HTML2 = RegisterClipboardFormat(StrPtr("HTML (HyperText Markup Language)"))
  CFSTR_LOGICALPERFORMEDDROPEFFECT = RegisterClipboardFormat(StrPtr("Logical Performed DropEffect"))
  CFSTR_MOZILLAHTMLCONTEXT = RegisterClipboardFormat(StrPtr("text/_moz_htmlcontext"))
  CFSTR_PERFORMEDDROPEFFECT = RegisterClipboardFormat(StrPtr("Performed DropEffect"))
  CFSTR_TARGETCLSID = RegisterClipboardFormat(StrPtr("TargetCLSID"))
  CFSTR_TEXTHTML = RegisterClipboardFormat(StrPtr("text/html"))

  CLSIDFromString StrPtr(strCLSID_RecycleBin), CLSID_RecycleBin

  hImgLst = ImageList_Create(16, 16, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 14, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
  iconsDir = iconsDir & IIf(bComctl32Version600OrNewer, "16x16x32bpp\", "16x16x8bpp\")
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
  Set ToolTip = New clsToolTip
  With ToolTip
    .Create Me.hWnd
    .Activate
    .AddTool chkOLEDragDrop.hWnd, "Check to start OLE Drag'n'Drop operations instead of normal ones.", , , False
  End With

  ' this is required to make the control work as expected
  Subclass

  InsertItems
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
  Set ToolTip = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub mnuAeroDragImages_Click()
  ExTVwU.OLEDragImageStyle = IIf(Not mnuAeroDragImages.Checked, OLEDragImageStyleConstants.odistAeroIfAvailable, OLEDragImageStyleConstants.odistClassic)
  mnuAeroDragImages.Checked = (ExTVwU.OLEDragImageStyle = OLEDragImageStyleConstants.odistAeroIfAvailable)
End Sub

Private Sub mnuCancel_Click()
  chosenMenuItem = ChosenMenuItemConstants.ciCancel
End Sub

Private Sub mnuCopy_Click()
  chosenMenuItem = ChosenMenuItemConstants.ciCopy
End Sub

Private Sub mnuMove_Click()
  chosenMenuItem = ChosenMenuItemConstants.ciMove
End Sub


Private Function CalculateInsertMarkPosition(ByVal dropTarget As ExTVwLibUCtl.ITreeViewItem, ByVal yToItemTop As Long) As ExTVwLibUCtl.InsertMarkPositionConstants
  Dim cy As Long
  Dim ret As InsertMarkPositionConstants
  Dim yBottom As Long
  Dim yTop As Long

  ret = InsertMarkPositionConstants.impNowhere
  If Not (dropTarget Is Nothing) Then
    dropTarget.GetRectangle irtItem, , yTop, , yBottom
    cy = yBottom - yTop
    If yToItemTop < cy / 3 Then
      ret = InsertMarkPositionConstants.impBefore
    ElseIf yToItemTop > cy * 2 / 3 Then
      ret = InsertMarkPositionConstants.impAfter
    Else
      ret = InsertMarkPositionConstants.impNowhere
    End If
  End If

  CalculateInsertMarkPosition = ret
End Function

Private Sub InsertItems()
  Dim cImages As Long
  Dim iIcon As Long
  Dim tiU1 As ExTVwLibUCtl.TreeViewItem
  Dim tiU2 As ExTVwLibUCtl.TreeViewItem

  If usingThemes Then
    ' for Windows Vista
    SetWindowTheme ExTVwU.hWnd, StrPtr("explorer"), 0
  End If

  ExTVwU.hImageList(ImageListConstants.ilItems) = hImgLst
  cImages = ImageList_GetImageCount(hImgLst)

  With ExTVwU.TreeItems
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

Private Function IsEqualGUID(IID1 As GUID, IID2 As GUID) As Boolean
  Dim Tmp1 As Currency
  Dim Tmp2 As Currency

  If IID1.Data1 = IID2.Data1 Then
    If IID1.Data2 = IID2.Data2 Then
      If IID1.Data3 = IID2.Data3 Then
        CopyMemory VarPtr(Tmp1), VarPtr(IID1.Data4(0)), 8
        CopyMemory VarPtr(Tmp2), VarPtr(IID2.Data4(0)), 8
        IsEqualGUID = (Tmp1 = Tmp2)
      End If
    End If
  End If
End Function

Private Function IsSameItem(item1 As ExTVwLibUCtl.ITreeViewItem, item2 As ExTVwLibUCtl.ITreeViewItem) As Boolean
  If item1 Is Nothing Then
    IsSameItem = (item2 Is Nothing)
  ElseIf item2 Is Nothing Then
    IsSameItem = False
  Else
    IsSameItem = (item1.Handle = item2.Handle)
  End If
End Function

' subclasses this Form
Private Sub Subclass()
  Const NF_REQUERY = 4
  Const WM_NOTIFYFORMAT = &H55

  #If UseSubClassing Then
    If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
      Debug.Print "Subclassing failed!"
    End If
    ' tell the control to negotiate the correct format with the form
    SendMessageAsLong ExTVwU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
