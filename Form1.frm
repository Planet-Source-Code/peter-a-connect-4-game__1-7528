VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Connect4"
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   DrawMode        =   6  'Mask Pen Not
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "BackGround"
      Height          =   240
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   75
      Width           =   1095
   End
   Begin VB.Timer WINceleb 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   255
      Top             =   2025
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Caption         =   "Well Done!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   945
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6285
      Visible         =   0   'False
      Width           =   5370
   End
   Begin VB.Timer MoveCounter 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   255
      Top             =   1515
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      Height          =   240
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      Width           =   540
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Go"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3390
      TabIndex        =   18
      Top             =   6015
      Width           =   1005
   End
   Begin VB.Label GoName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3375
      TabIndex        =   17
      Top             =   5850
      Width           =   1005
   End
   Begin VB.Label GoRow 
      BackStyle       =   0  'Transparent
      Height          =   750
      Index           =   6
      Left            =   5790
      TabIndex        =   7
      Top             =   330
      Width           =   645
   End
   Begin VB.Label GoRow 
      BackStyle       =   0  'Transparent
      Height          =   750
      Index           =   1
      Left            =   1590
      TabIndex        =   2
      Top             =   330
      Width           =   645
   End
   Begin VB.Label GoRow 
      BackStyle       =   0  'Transparent
      Height          =   750
      Index           =   5
      Left            =   4950
      TabIndex        =   6
      Top             =   330
      Width           =   645
   End
   Begin VB.Label GoRow 
      BackStyle       =   0  'Transparent
      Height          =   750
      Index           =   4
      Left            =   4110
      TabIndex        =   5
      Top             =   330
      Width           =   645
   End
   Begin VB.Label GoRow 
      BackStyle       =   0  'Transparent
      Height          =   750
      Index           =   3
      Left            =   3270
      TabIndex        =   4
      Top             =   330
      Width           =   645
   End
   Begin VB.Label GoRow 
      BackStyle       =   0  'Transparent
      Height          =   750
      Index           =   2
      Left            =   2430
      TabIndex        =   3
      Top             =   330
      Width           =   645
   End
   Begin VB.Label GoRow 
      BackStyle       =   0  'Transparent
      Height          =   750
      Index           =   0
      Left            =   750
      TabIndex        =   0
      Top             =   330
      Width           =   645
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   225
      Left            =   6060
      TabIndex        =   16
      Top             =   570
      Width           =   120
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   225
      Left            =   5205
      TabIndex        =   15
      Top             =   555
      Width           =   120
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   225
      Left            =   4380
      TabIndex        =   14
      Top             =   600
      Width           =   120
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   225
      Left            =   3525
      TabIndex        =   13
      Top             =   570
      Width           =   120
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   225
      Left            =   2715
      TabIndex        =   12
      Top             =   585
      Width           =   120
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   225
      Left            =   1875
      TabIndex        =   11
      Top             =   570
      Width           =   120
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   225
      Left            =   1005
      TabIndex        =   10
      Top             =   555
      Width           =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   7155
      X2              =   7155
      Y1              =   15
      Y2              =   6840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   15
      X2              =   7155
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   15
      X2              =   7155
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   15
      X2              =   15
      Y1              =   15
      Y2              =   6840
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   1
      Left            =   1590
      Shape           =   3  'Circle
      Top             =   330
      Width           =   645
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   3900
      Left            =   6645
      Top             =   2850
      Width           =   225
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   3900
      Left            =   315
      Top             =   2835
      Width           =   225
   End
   Begin VB.Shape Row7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   5
      Left            =   5790
      Shape           =   3  'Circle
      Top             =   1275
      Width           =   645
   End
   Begin VB.Shape Row6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   5
      Left            =   4950
      Shape           =   3  'Circle
      Top             =   1275
      Width           =   645
   End
   Begin VB.Shape Row5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   5
      Left            =   4110
      Shape           =   3  'Circle
      Top             =   1275
      Width           =   645
   End
   Begin VB.Shape Row4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   5
      Left            =   3270
      Shape           =   3  'Circle
      Top             =   1275
      Width           =   645
   End
   Begin VB.Shape Row3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   5
      Left            =   2430
      Shape           =   3  'Circle
      Top             =   1275
      Width           =   645
   End
   Begin VB.Shape Row2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   5
      Left            =   1590
      Shape           =   3  'Circle
      Top             =   1275
      Width           =   645
   End
   Begin VB.Shape Row1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   5
      Left            =   750
      Shape           =   3  'Circle
      Top             =   1275
      Width           =   645
   End
   Begin VB.Shape Row7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   4
      Left            =   5790
      Shape           =   3  'Circle
      Top             =   2010
      Width           =   645
   End
   Begin VB.Shape Row6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   4
      Left            =   4950
      Shape           =   3  'Circle
      Top             =   2010
      Width           =   645
   End
   Begin VB.Shape Row5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   4
      Left            =   4110
      Shape           =   3  'Circle
      Top             =   2010
      Width           =   645
   End
   Begin VB.Shape Row4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   4
      Left            =   3270
      Shape           =   3  'Circle
      Top             =   2010
      Width           =   645
   End
   Begin VB.Shape Row3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   4
      Left            =   2430
      Shape           =   3  'Circle
      Top             =   2010
      Width           =   645
   End
   Begin VB.Shape Row2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   4
      Left            =   1590
      Shape           =   3  'Circle
      Top             =   2010
      Width           =   645
   End
   Begin VB.Shape Row1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   4
      Left            =   750
      Shape           =   3  'Circle
      Top             =   2010
      Width           =   645
   End
   Begin VB.Shape Row7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   3
      Left            =   5790
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   645
   End
   Begin VB.Shape Row6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   3
      Left            =   4950
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   645
   End
   Begin VB.Shape Row5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   3
      Left            =   4110
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   645
   End
   Begin VB.Shape Row4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   3
      Left            =   3270
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   645
   End
   Begin VB.Shape Row3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   3
      Left            =   2430
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   645
   End
   Begin VB.Shape Row2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   3
      Left            =   1590
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   645
   End
   Begin VB.Shape Row1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   3
      Left            =   750
      Shape           =   3  'Circle
      Top             =   2745
      Width           =   645
   End
   Begin VB.Shape Row7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   2
      Left            =   5790
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   645
   End
   Begin VB.Shape Row6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   2
      Left            =   4950
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   645
   End
   Begin VB.Shape Row5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   2
      Left            =   4110
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   645
   End
   Begin VB.Shape Row4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   2
      Left            =   3270
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   645
   End
   Begin VB.Shape Row3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   2
      Left            =   2430
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   645
   End
   Begin VB.Shape Row2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   2
      Left            =   1590
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   645
   End
   Begin VB.Shape Row1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   2
      Left            =   750
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   645
   End
   Begin VB.Shape Row7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   1
      Left            =   5790
      Shape           =   3  'Circle
      Top             =   4215
      Width           =   645
   End
   Begin VB.Shape Row6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   1
      Left            =   4950
      Shape           =   3  'Circle
      Top             =   4215
      Width           =   645
   End
   Begin VB.Shape Row5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   1
      Left            =   4110
      Shape           =   3  'Circle
      Top             =   4215
      Width           =   645
   End
   Begin VB.Shape Row4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   1
      Left            =   3270
      Shape           =   3  'Circle
      Top             =   4215
      Width           =   645
   End
   Begin VB.Shape Row3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   1
      Left            =   2430
      Shape           =   3  'Circle
      Top             =   4215
      Width           =   645
   End
   Begin VB.Shape Row2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   1
      Left            =   1590
      Shape           =   3  'Circle
      Top             =   4215
      Width           =   645
   End
   Begin VB.Shape Row1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   1
      Left            =   750
      Shape           =   3  'Circle
      Top             =   4215
      Width           =   645
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   0
      Left            =   750
      Shape           =   3  'Circle
      Top             =   330
      Width           =   645
   End
   Begin VB.Shape Row7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   0
      Left            =   5790
      Shape           =   3  'Circle
      Top             =   4950
      Width           =   645
   End
   Begin VB.Shape Row6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   0
      Left            =   4950
      Shape           =   3  'Circle
      Top             =   4950
      Width           =   645
   End
   Begin VB.Shape Row5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   0
      Left            =   4110
      Shape           =   3  'Circle
      Top             =   4950
      Width           =   645
   End
   Begin VB.Shape Row4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   0
      Left            =   3270
      Shape           =   3  'Circle
      Top             =   4950
      Width           =   645
   End
   Begin VB.Shape Row3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   0
      Left            =   2430
      Shape           =   3  'Circle
      Top             =   4950
      Width           =   645
   End
   Begin VB.Shape Row2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   0
      Left            =   1590
      Shape           =   3  'Circle
      Top             =   4950
      Width           =   645
   End
   Begin VB.Shape Row1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   0
      Left            =   750
      Shape           =   3  'Circle
      Top             =   4950
      Width           =   645
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      BorderStyle     =   0  'Transparent
      Height          =   4635
      Left            =   540
      Top             =   1170
      Width           =   6105
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   2
      Left            =   2430
      Shape           =   3  'Circle
      Top             =   330
      Width           =   645
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   3
      Left            =   3270
      Shape           =   3  'Circle
      Top             =   330
      Width           =   645
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   4
      Left            =   4110
      Shape           =   3  'Circle
      Top             =   330
      Width           =   645
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   5
      Left            =   4950
      Shape           =   3  'Circle
      Top             =   330
      Width           =   645
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   750
      Index           =   6
      Left            =   5790
      Shape           =   3  'Circle
      Top             =   330
      Width           =   645
   End
   Begin VB.Shape WhoGo 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   435
      Left            =   3300
      Shape           =   2  'Oval
      Top             =   5790
      Width           =   1110
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim CurrRow, NameRow(0 To 5), DN
Dim board(1 To 7) As boardt
Dim flash(1 To 4) As flash, fintmrcnt
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Command3_KeyDown(KeyCode, Shift)
End Sub

Private Sub Command2_Click()
WINceleb.Enabled = False
Form2.Timer1.Enabled = False
Form2.Timer1.Enabled = True
Unload Me
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
Call Command3_KeyDown(KeyCode, Shift)
End Sub

Private Sub Command3_Click()
If Line1.Visible = False Then
    Form2.Visible = False
    Line1.Visible = True
    Line2.Visible = True
    Line3.Visible = True
    Line4.Visible = True
Else
    Form2.Visible = True
    Me.SetFocus
    Line1.Visible = False
    Line2.Visible = False
    Line3.Visible = False
    Line4.Visible = False
End If
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)

If Me.WhoGo.BackColor = &HFF& Then
    If Form2.Option1.Value = True Then Exit Sub
Else
    If Form2.Option6.Value = True Then Exit Sub
End If

If KeyCode = vbKey1 Then
    Call GoRow2(0, "KEY")
ElseIf KeyCode = vbKey2 Then
    Call GoRow2(1, "KEY")
ElseIf KeyCode = vbKey3 Then
    Call GoRow2(2, "KEY")
ElseIf KeyCode = vbKey4 Then
    Call GoRow2(3, "KEY")
ElseIf KeyCode = vbKey5 Then
    Call GoRow2(4, "KEY")
ElseIf KeyCode = vbKey6 Then
    Call GoRow2(5, "KEY")
ElseIf KeyCode = vbKey7 Then
    Call GoRow2(6, "KEY")
Else
    MsgBox "Press 1, 2, 3, 4, 5, 6 or 7 to play.", vbInformation
End If
End Sub

Private Sub Command4_Click()
Form2.Visible = True
End Sub

Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
Call Command3_KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call Command3_KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_Load()
CurrRow = 0
DN = 0
For i = 0 To 5
NameRow(i) = 0
Next
fintmrcnt = 0

Form2.Show
For i = 0 To 5
    NameRow(i) = &HC0C0C0
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Form2.Timer1.Enabled = False Then Unload Form2
End Sub
Sub GoRow2(index, from)
If Me.MoveCounter.Enabled = True Or Me.Command2.Visible = True Then Exit Sub
CurrRow = index + 1
DN = 5
If getcolour(5) <> &HC0C0C0 Then
    MsgBox "Col Full", vbExclamation
    Exit Sub
End If
Shape(index).BackColor = WhoGo.BackColor



MoveCounter.Enabled = True
End Sub
Private Sub GoRow_Click(index As Integer)
If Me.WhoGo.BackColor = &HFF& Then
    If Form2.Option5.Value = True Then Exit Sub
Else
    If Form2.Option2.Value = True Then Exit Sub
End If
Call GoRow2(index, "CLICK")
End Sub

Private Sub MoveCounter_Timer()
If DN = -1 Then 'It has reached the bottom!
    Call GotOne
    Exit Sub
End If
If getcolour(DN) <> &HC0C0C0 Then 'Theres on colour there!
    Call GotOne
    Exit Sub
End If

If DN < 5 Then
    NameRow(DN + 1) = &HC0C0C0
    Call setcolour(DN + 1)
End If

NameRow(DN) = WhoGo.BackColor
Call setcolour(DN)

DN = DN - 1

End Sub
Sub WINceleb_Timer()
If WhoGo.BackColor = &HFF& Then
    WinNed = "RED"
    Me.Command2.Caption = "Well Done " + Form2.Pr.Text + "!"
Else
    WinNed = "BLUE"
    Me.Command2.Caption = "Well Done " + Form2.Pb.Text + "!"
End If

For x = 0 To 5
    For y = 1 To 7
            CurrRow = y
            If getcolour(x) = &HC0C0C0 Then
                NameRow(x) = &HFFFF&
                Call setcolour(x)
            End If
    Next
Next
WINceleb.Enabled = True
fintmrcnt = fintmrcnt + 1
If fintmrcnt > 10 Then
    WINceleb.Enabled = False
    fintmrcnt = 0
    Exit Sub
End If
For i = 1 To 4 'FLASH WINNING ROW
    If NameRow(flash(i).Top) = &HFFFFFF Then
        NameRow(flash(i).Top) = &H0&
    Else
        NameRow(flash(i).Top) = &HFFFFFF
    End If
    CurrRow = flash(i).Left
    Call setcolour(flash(i).Top)
Next
             Me.Command2.BackColor = WhoGo.BackColor
             Me.Command2.Visible = True
End Sub
Sub GotOne()
MoveCounter.Enabled = False
Shape(CurrRow - 1).BackColor = RGB(255, 255, 255)
'****************SET GRID******************************
For ii = 1 To 7
    CurrRow = ii
    For i = 0 To 5
        board(ii).rows(i) = "!!!NONE!!!"
        If getcolour(i) = WhoGo.BackColor Then
            board(ii).rows(i) = "COLOUR"
        End If
    Next
Next

'****************FIND HORZ******************************
For x = 0 To 5
    clr = "COLOUR"
    For v = 1 To 4
        If board(v).rows(x) = clr And board(v + 1).rows(x) = clr And board(v + 2).rows(x) = clr And board(v + 3).rows(x) = clr Then
        For i = 1 To 4
            flash(i).Left = v + (i - 1)
            flash(i).Top = x
        Next
             Call WINceleb_Timer
        End If
    Next
Next

'****************FIND VERT******************************
For x = 1 To 7
    clr = "COLOUR"
    For v = 0 To 2
        If board(x).rows(v) = clr And board(x).rows(v + 1) = clr And board(x).rows(v + 2) = clr And board(x).rows(v + 3) = clr Then
        For i = 1 To 4
            flash(i).Left = x
            flash(i).Top = v + (i - 1)
        Next
             Call WINceleb_Timer
        End If
    Next
Next


'****************FIND T/L -B/R ******************************
For x = 1 To 4
    clr = "COLOUR"
    For v = 3 To 5
        If board(x).rows(v) = clr And board(x + 1).rows(v - 1) = clr And board(x + 2).rows(v - 2) = clr And board(x + 3).rows(v - 3) = clr Then
        For i = 1 To 4
            flash(i).Left = x + (i - 1)
            flash(i).Top = v - (i - 1)
        Next
            Call WINceleb_Timer
        End If
    Next
Next


'****************FIND B/L -T/R ******************************

For x = 1 To 4
    clr = "COLOUR"
    For v = 0 To 2
        If board(x).rows(v) = clr And board(x + 1).rows(v + 1) = clr And board(x + 2).rows(v + 2) = clr And board(x + 3).rows(v + 3) = clr Then
        For i = 1 To 4
            flash(i).Left = x + (i - 1)
            flash(i).Top = v + (i - 1)
        Next
            Call WINceleb_Timer
        End If
    Next
Next

'****************CHANGE GO******************************
If WhoGo.BackColor = RGB(255, 0, 0) Then
    WhoGo.BackColor = RGB(0, 0, 255)
    GoName.Caption = Form2.Pb.Text
Else
    WhoGo.BackColor = RGB(255, 0, 0)
        GoName.Caption = Form2.Pr.Text
End If

End Sub








Function setcolour(i)
If CurrRow = 1 Then
    Row1(i).BackColor = NameRow(i)
ElseIf CurrRow = 2 Then
    Row2(i).BackColor = NameRow(i)
ElseIf CurrRow = 3 Then
    Row3(i).BackColor = NameRow(i)
ElseIf CurrRow = 4 Then
    Row4(i).BackColor = NameRow(i)
ElseIf CurrRow = 5 Then
    Row5(i).BackColor = NameRow(i)
ElseIf CurrRow = 6 Then
    Row6(i).BackColor = NameRow(i)
ElseIf CurrRow = 7 Then
    Row7(i).BackColor = NameRow(i)
End If
End Function

Function getcolour(i)
If CurrRow = 1 Then
    getcolour = Row1(i).BackColor
ElseIf CurrRow = 2 Then
    getcolour = Row2(i).BackColor
ElseIf CurrRow = 3 Then
    getcolour = Row3(i).BackColor
ElseIf CurrRow = 4 Then
    getcolour = Row4(i).BackColor
ElseIf CurrRow = 5 Then
    getcolour = Row5(i).BackColor
ElseIf CurrRow = 6 Then
    getcolour = Row6(i).BackColor
ElseIf CurrRow = 7 Then
    getcolour = Row7(i).BackColor
End If
End Function


