VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Show Board"
      Height          =   270
      Left            =   570
      TabIndex        =   10
      Top             =   1620
      Width           =   1050
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   15
      ScaleHeight     =   720
      ScaleWidth      =   1695
      TabIndex        =   5
      Top             =   825
      Width           =   1725
      Begin VB.TextBox Pb 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   0
         MaxLength       =   8
         TabIndex        =   9
         Text            =   "Blue Player"
         Top             =   0
         Width           =   1005
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Mouse"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1035
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   45
         Width           =   600
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Keyboard"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   735
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   375
         Width           =   915
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Both"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   375
         Value           =   -1  'True
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   30
      ScaleHeight     =   720
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   60
      Width           =   1725
      Begin VB.OptionButton Option3 
         Caption         =   "Both"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   375
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Keyboard"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   735
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   375
         Width           =   915
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mouse"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1050
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   45
         Width           =   600
      End
      Begin VB.TextBox Pr 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   30
         MaxLength       =   8
         TabIndex        =   1
         Text            =   "Red Player"
         Top             =   0
         Width           =   1005
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   645
      Top             =   2040
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   1725
      Picture         =   "Form2.frx":0000
      Top             =   -30
      Width           =   4920
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form1.SetFocus
End Sub

Private Sub Form_Click()
Me.Visible = False
End Sub

Private Sub Form_Resize()
Image1.Left = (Screen.Width / 2) - (Image1.Width / 2)
End Sub

Private Sub Timer1_Timer()
Load Form1
Form1.Show
Timer1.Enabled = False
End Sub
