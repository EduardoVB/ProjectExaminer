VERSION 5.00
Begin VB.Form frmSelectItems 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select items"
   ClientHeight    =   5040
   ClientLeft      =   6264
   ClientTop       =   5076
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSelectNone 
      Caption         =   "Select None"
      Height          =   348
      Left            =   1896
      TabIndex        =   4
      Top             =   100
      Width           =   1524
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      Height          =   348
      Left            =   240
      TabIndex        =   3
      Top             =   100
      Width           =   1524
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   348
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   1524
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   348
      Left            =   1896
      TabIndex        =   1
      Top             =   4560
      Width           =   1524
   End
   Begin VB.ListBox lstItems 
      Height          =   3936
      Left            =   96
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   520
      Width           =   3516
   End
End
Attribute VB_Name = "frmSelectItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OKPressed As Boolean

Private mItems() As String
Private mSelected() As Boolean

Public Sub AddItem(nItem As Variant, nItemCaption As Variant, nSelected As Boolean)
    Dim i As Long
    
    i = UBound(mItems) + 1
    ReDim Preserve mItems(i)
    ReDim Preserve mSelected(i)
    mItems(i) = nItem
    lstItems.AddItem nItemCaption
    lstItems.ItemData(lstItems.NewIndex) = i
    If nSelected Then
        lstItems.Selected(lstItems.NewIndex) = True
        mSelected(i) = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    OKPressed = True
    Unload Me
End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Long
    
    For i = 0 To lstItems.ListCount - 1
        lstItems.Selected(i) = True
    Next
End Sub

Private Sub cmdSelectNone_Click()
    Dim i As Long
    
    For i = 0 To lstItems.ListCount - 1
        lstItems.Selected(i) = False
    Next
End Sub

Private Sub Form_Initialize()
    ReDim mItems(0)
    ReDim mSelected(0)
End Sub

Private Sub lstItems_ItemCheck(Item As Integer)
    mSelected(lstItems.ItemData(Item)) = lstItems.Selected(Item)
End Sub

Public Property Get ItemCount() As Long
    ItemCount = UBound(mItems)
End Property

Public Property Get Item(nIndex As Long) As String
    Item = mItems(nIndex)
End Property

Public Property Get Selected(nIndex As Long) As Boolean
    Selected = mSelected(nIndex)
End Property
