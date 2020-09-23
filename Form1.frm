VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4335
      TabIndex        =   1
      Text            =   "Nothing selected yet."
      Top             =   165
      Width           =   2460
   End
   Begin Project1.FSL FSL1 
      Height          =   405
      Left            =   75
      TabIndex        =   0
      Top             =   180
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   714
      DisplayBkgd     =   12648447
      ListBkgd        =   12648447
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub FSL1_Click()
   Text1.Text = FSL1.Selected
End Sub
