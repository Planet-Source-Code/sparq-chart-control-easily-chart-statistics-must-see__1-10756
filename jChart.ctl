VERSION 5.00
Begin VB.UserControl jChart 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6525
   ScaleHeight     =   1350
   ScaleWidth      =   6525
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Index           =   0
      Left            =   540
      TabIndex        =   0
      Top             =   600
      Width           =   45
   End
   Begin VB.Shape Item 
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   1860
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Border 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "jChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

Dim Items As chrtItems
Dim UseableWidth As Integer
Dim IndivWidth As Integer
Dim TopValue As Integer
Dim BottomValue As Integer

Private Type chrtItems
    ItemValue(0 To 99) As Single
    ItemLabel(0 To 99) As String
    ItemLabelColor(0 To 99) As OLE_COLOR
    ItemColor(0 To 99) As OLE_COLOR
    ItemCount As Integer
End Type


Private Sub UserControl_Initialize()
    UserControl.BackColor = vbWhite
    TopValue = 100
    BottomValue = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", vbWhite)
End Sub

Private Sub UserControl_Resize()
    With Border
        .Left = 0
        .Top = 0
        .Width = Width
        .Height = Height
    End With
    
    If Items.ItemCount < 1 Then Exit Sub
    UseableWidth = Width - 100
    IndivWidth = UseableWidth / Items.ItemCount
    DrawItems
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", BackColor, vbWhite)
End Sub


Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Get ItemCount() As Integer
    ItemCount = Items.ItemCount
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Function AddItem(Label As String, Value As Integer, Color As OLE_COLOR, LabelColor As OLE_COLOR)
  Dim X As Integer
    X = Items.ItemCount
    Items.ItemValue(X) = Value
    Items.ItemLabel(X) = Label
    Items.ItemColor(X) = Color
    Items.ItemLabelColor(X) = LabelColor
    Items.ItemCount = Items.ItemCount + 1
    UserControl_Resize
End Function

Private Function DrawItems()
    Dim NextSpot As Integer
    Dim X As Integer
    Dim UnitHeight As Integer
    
    UnitHeight = (Height - 100) / 100
    
    On Error Resume Next
    NextSpot = 50
    For X = 0 To Items.ItemCount - 1
        Load Item(X)
        Item(X).Left = NextSpot
        Item(X).Width = IndivWidth
        Item(X).FillColor = Items.ItemColor(X)
        Item(X).Height = Items.ItemValue(X) * UnitHeight
        Item(X).Top = (Height - 50) - Item(X).Height
        Item(X).Visible = True
        
        Load Label(X)
        Label(X).Caption = Items.ItemLabel(X)
        Label(X).Left = Item(X).Left + 30
        If Items.ItemValue(X) > 93 Then
            Label(X).Top = Item(X).Top + 30
            Label(X).ZOrder 0
        Else
            Label(X).Top = Item(X).Top - (Label(X).Height)
        End If
        Label(X).ForeColor = Items.ItemLabelColor(X)
        Label(X).Visible = True
        
        NextSpot = NextSpot + IndivWidth
    Next X
End Function

Public Function Clear()
  Dim X As Integer
    For X = 0 To Items.ItemCount - 1
        With Items
            .ItemCount = 0
            .ItemColor(X) = vbWhite
            .ItemLabel(X) = ""
            .ItemLabelColor(X) = vbWhite
            .ItemValue(X) = 0
            
            Label(X).Visible = False
            Item(X).Visible = False
        End With
    Next X
End Function
