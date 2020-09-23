VERSION 5.00
Begin VB.UserControl FSL 
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4260
   ScaleHeight     =   690
   ScaleWidth      =   4260
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   4305
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   375
      Width           =   1935
   End
   Begin VB.PictureBox picFontsList 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   390
      Width           =   4065
      Begin VB.VScrollBar VScroll1 
         Height          =   2520
         LargeChange     =   5
         Left            =   3750
         Max             =   210
         SmallChange     =   3
         TabIndex        =   6
         Top             =   0
         Width           =   285
      End
      Begin VB.PictureBox picFonts 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2580
         Left            =   -15
         ScaleHeight     =   2580
         ScaleWidth      =   3765
         TabIndex        =   4
         Top             =   0
         Width           =   3765
         Begin VB.Label lblFonts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Blank"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   1725
            TabIndex        =   5
            Top             =   -15
            Width           =   420
         End
      End
   End
   Begin VB.PictureBox picDisplay 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   0
      Width           =   4065
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3705
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   -30
         Width           =   390
      End
      Begin VB.Label lblDisplay 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Font"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         TabIndex        =   2
         Top             =   0
         Width           =   1410
      End
   End
End
Attribute VB_Name = "FSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************
'**                              Font Selector
'**                               Version 1.0.0
'**                               By Ken Foster
'**                                 June 2005
'**                     Freeware--- no copyrights claimed
'**                      Yes this means you can steal it.
'*******************************************************************

'***************** Table of Procedures *************
'   Private Sub UserControl_Initialize
'   Private Sub UserControl_Resize
'   Private Sub UserControl_ReadProperties
'   Private Sub UserControl_WriteProperties
'   Private Sub Fill_Labels
'   Private Sub cmdSelect_Click
'   Private Sub lblFonts_Click
'   Private Sub VScroll1_Change
'   Private Sub VScroll1_Scroll
'   Public Property Get Selected
'   Public Property Let Selected
'   Public Property Get DisplayBkgd
'   Public Property Let DisplayBkgd
'   Public Property Get DisplayFontCol
'   Public Property Let DisplayFontCol
'   Public Property Get ListBkgd
'   Public Property Let ListBkgd
'   Public Property Get ListFontCol
'   Public Property Let ListFontCol
'***************** End of Table ********************

Const m_def_Selected = "Read Only"
Const m_def_DisplayBkgd = vbWhite
Const m_def_DisplayFontCol = vbBlack
Const m_def_FontListBkgd = vbWhite
Const m_def_FontListFontCol = vbBlack

Dim m_DisplayBkgd As OLE_COLOR
Dim m_DisplayFontCol As OLE_COLOR
Dim m_FontListBkgd As OLE_COLOR
Dim m_FontListFontCol As OLE_COLOR
Dim m_Selected As String
Dim PicH As Integer  ' sets height of picFontsList
 Event Click()

Private Sub UserControl_Initialize()
    Dim I As Integer
    Dim x As Integer
    
   'load fonts into listbox and sort
   For I = 0 To Screen.FontCount - 1
      List1.AddItem Screen.Fonts(I)
   Next I
   Fill_Labels  'put fonts in picFontsList
   
    For x = 0 To 8
      PicH = PicH + lblFonts(x).Height + 50  'sets picFontsList Height to first 9 fonts plus a little extra
    Next x
End Sub

Private Sub UserControl_Resize()
   If picFontsList.Height <> PicH Then UserControl.Height = picDisplay.Height
   UserControl.Width = picDisplay.Width
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_Selected = PropBag.ReadProperty("Selected", m_def_Selected)
   m_DisplayBkgd = PropBag.ReadProperty("DisplayBkgd", m_def_DisplayBkgd)
   m_DisplayFontCol = PropBag.ReadProperty("DisplayFontCol", m_def_DisplayFontCol)
   m_FontListBkgd = PropBag.ReadProperty("ListBkgd", m_def_FontListBkgd)
   m_FontListFontCol = PropBag.ReadProperty("ListFontCol", m_def_FontListFontCol)
   
   DisplayFontCol = m_DisplayFontCol
   DisplayBkgd = m_DisplayBkgd
   ListBkgd = m_FontListBkgd
   ListFontCol = m_FontListFontCol
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "DisplayBkgd", m_DisplayBkgd, m_def_DisplayBkgd
   PropBag.WriteProperty "DisplayFontCol", m_DisplayFontCol, m_def_DisplayFontCol
   PropBag.WriteProperty "ListBkgd", m_FontListBkgd, m_def_FontListBkgd
   PropBag.WriteProperty "ListFontCol", m_FontListFontCol, m_def_FontListFontCol
End Sub

Private Sub Fill_Labels()
   Dim x As Integer
   Dim y As Long
   
   'sets first font in list
   lblFonts(0).Caption = List1.List(0)
   lblFonts(0).Font = List1.List(0)
   lblFonts(0).FontSize = 14
   
   'set all other fonts in list
   For x = 1 To List1.ListCount - 1
      Load lblFonts(x)
      lblFonts(x).Visible = True
      lblFonts(x).Height = lblFonts(0).Height
      lblFonts(x).Top = lblFonts(x - 1).Top + lblFonts(0).Height + 50
      lblFonts(x).FontSize = 14
      lblFonts(x).Font = List1.List(x)
      lblFonts(x).Caption = List1.List(x)
      y = y + lblFonts(x).Height + 50
   Next x
   
   picFonts.Height = y + lblFonts(0).Height
   
End Sub

Private Sub cmdSelect_Click()

   'show\hide fonts window
   If picFontsList.Height = PicH Then
      picFontsList.Height = 0
      UserControl.Height = picDisplay.Height
   Else
      picFontsList.Height = PicH
      UserControl.Height = picFontsList.Height + picDisplay.Height
      VScroll1.Height = picFontsList.Height - 20
   End If
   
   List1.SetFocus  'takes focus off cmdSelect button so rectangle does,nt show
   
End Sub

Private Sub lblFonts_Click(Index As Integer)

   'put selected font in display window
   lblDisplay.Caption = lblFonts(Index).Caption
   lblDisplay.Font = lblFonts(Index).Font
   lblDisplay.FontSize = 14
   picFontsList.Height = 0  'hide picFontsList
   Selected = lblDisplay
   
   RaiseEvent Click
   
End Sub

Private Sub VScroll1_Change()
   picFonts.Top = -(VScroll1.Value * (lblFonts(0).Height + 50))
End Sub

Private Sub VScroll1_Scroll()
   picFonts.Top = -(VScroll1.Value * (lblFonts(0).Height + 50))
End Sub

Public Property Get Selected() As String
   Selected = m_Selected
End Property

Public Property Let Selected(NewSelected As String)
   m_Selected = NewSelected
   PropertyChanged "Selected"
End Property

Public Property Get DisplayBkgd() As OLE_COLOR
   DisplayBkgd = m_DisplayBkgd
End Property

Public Property Let DisplayBkgd(NewDisplayBkgd As OLE_COLOR)
   m_DisplayBkgd = NewDisplayBkgd
   picDisplay.BackColor = m_DisplayBkgd
   PropertyChanged "DisplayBkgd"
End Property

Public Property Get DisplayFontCol() As OLE_COLOR
   DisplayFontCol = m_DisplayFontCol
End Property

Public Property Let DisplayFontCol(NewDisplayFontCol As OLE_COLOR)
   m_DisplayFontCol = NewDisplayFontCol
   lblDisplay.ForeColor = m_DisplayFontCol
   PropertyChanged "DisplayFontCol"
End Property

Public Property Get ListBkgd() As OLE_COLOR
   ListBkgd = m_FontListBkgd
End Property

Public Property Let ListBkgd(NewListBkgd As OLE_COLOR)
   m_FontListBkgd = NewListBkgd
   picFonts.BackColor = m_FontListBkgd
   PropertyChanged "ListBkgd"
End Property

Public Property Get ListFontCol() As OLE_COLOR
   ListFontCol = m_FontListFontCol
End Property

Public Property Let ListFontCol(NewListFontCol As OLE_COLOR)
   Dim x As Integer
   
   m_FontListFontCol = NewListFontCol
   For x = 0 To List1.ListCount - 1
      lblFonts(x).ForeColor = m_FontListFontCol
   Next x
   PropertyChanged "ListFontCol"
End Property
