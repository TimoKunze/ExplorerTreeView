VERSION 5.00
Begin VB.Form frmItemProperties 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Edit Item"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
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
   ScaleHeight     =   164
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   377
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtSelectedIconIndex 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   4920
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtIconIndex 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtHeightIncrement 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   4800
      TabIndex        =   12
      Top             =   1320
      Width           =   735
   End
   Begin VB.CheckBox chkHasExpando 
      Caption         =   """+""/""-""-Button (Expando)"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtItemData 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.CheckBox chkGhosted 
      Caption         =   "Ghosted"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Left            =   720
      MaxLength       =   259
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label lblDescr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SelectedIconIndex:"
      Height          =   195
      Index           =   4
      Left            =   3360
      TabIndex        =   7
      Top             =   1050
      Width           =   1410
   End
   Begin VB.Label lblDescr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IconIndex:"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   1050
      Width           =   795
   End
   Begin VB.Label lblDescr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HeightIncrement:"
      Height          =   195
      Index           =   2
      Left            =   3480
      TabIndex        =   11
      Top             =   1410
      Width           =   1260
   End
   Begin VB.Label lblDescr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ItemData:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1410
      Width           =   735
   End
   Begin VB.Label lblDescr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   390
   End
End
Attribute VB_Name = "frmItemProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private propCanceled As Boolean


Private Sub cmdCancel_Click()
  propCanceled = True
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  propCanceled = False
  Me.Hide
End Sub


Public Property Get Bold() As Boolean
  Bold = (chkBold.Value = CheckBoxConstants.vbChecked)
End Property

Public Property Let Bold(ByVal newValue As Boolean)
  chkBold.Value = IIf(newValue, CheckBoxConstants.vbChecked, CheckBoxConstants.vbUnchecked)
End Property

Public Property Get Canceled() As Boolean
  Canceled = propCanceled
End Property

Public Property Get Ghosted() As Boolean
  Ghosted = (chkGhosted.Value = CheckBoxConstants.vbChecked)
End Property

Public Property Let Ghosted(ByVal newValue As Boolean)
  chkGhosted.Value = IIf(newValue, CheckBoxConstants.vbChecked, CheckBoxConstants.vbUnchecked)
End Property

Public Property Get HasExpando() As Boolean
  HasExpando = (chkHasExpando.Value = CheckBoxConstants.vbChecked)
End Property

Public Property Let HasExpando(ByVal newValue As Boolean)
  chkHasExpando.Value = IIf(newValue, CheckBoxConstants.vbChecked, CheckBoxConstants.vbUnchecked)
End Property

Public Property Get HeightIncrement() As Long
  On Error Resume Next
  HeightIncrement = CLng(txtHeightIncrement.Text)
End Property

Public Property Let HeightIncrement(ByVal newValue As Long)
  txtHeightIncrement.Text = newValue
End Property

Public Property Get IconIndex() As Long
  On Error Resume Next
  IconIndex = CLng(txtIconIndex.Text)
End Property

Public Property Let IconIndex(ByVal newValue As Long)
  txtIconIndex.Text = newValue
End Property

Public Property Get ItemData() As Long
  On Error Resume Next
  ItemData = CLng(txtItemData.Text)
End Property

Public Property Let ItemData(ByVal newValue As Long)
  txtItemData.Text = newValue
End Property

Public Property Get SelectedIconIndex() As Long
  On Error Resume Next
  SelectedIconIndex = CLng(txtSelectedIconIndex.Text)
End Property

Public Property Let SelectedIconIndex(ByVal newValue As Long)
  txtSelectedIconIndex.Text = newValue
End Property

Public Property Get Text() As String
  Text = txtText.Text
End Property

Public Property Let Text(ByVal newValue As String)
  txtText.Text = newValue
End Property
