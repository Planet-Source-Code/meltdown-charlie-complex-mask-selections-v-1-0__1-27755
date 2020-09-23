VERSION 5.00
Begin VB.Form fBrushSettings 
   BackColor       =   &H00C5A774&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bursh Mask Settings"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2730
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   360
      Left            =   1500
      TabIndex        =   2
      Top             =   210
      Width           =   1065
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "fBrushSettings.frx":0000
      Left            =   195
      List            =   "fBrushSettings.frx":0031
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   330
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brush width"
      Height          =   195
      Left            =   225
      TabIndex        =   1
      Top             =   75
      Width           =   825
   End
End
Attribute VB_Name = "fBrushSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bLoaded As Boolean

Private Sub Combo1_Click()
    If bLoaded Then BrushWidth = Combo1.List(Combo1.ListIndex - 1)
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    bLoaded = False
    Combo1.ListIndex = BrushWidth - 1
    bLoaded = True
End Sub
