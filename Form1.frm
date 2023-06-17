VERSION 5.00
Begin VB.Form FrmLogin 
   Caption         =   "Login Form"
   ClientHeight    =   8775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14820
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8775
   ScaleWidth      =   14820
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3240
      Top             =   1800
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1635
      ItemData        =   "Form1.frx":3AC08
      Left            =   12720
      List            =   "Form1.frx":3AC1B
      TabIndex        =   7
      Top             =   4800
      Width           =   6015
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H000080FF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton CmdSubmit 
      BackColor       =   &H000080FF&
      Caption         =   "&Submit"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   9120
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   9120
      TabIndex        =   3
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Lbl2 
      BackColor       =   &H000080FF&
      Caption         =   "User Name :"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Lbl3 
      BackColor       =   &H000080FF&
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "WELCOME TO AC COLD DRINK SHOP"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   6735
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Flag As Single
Private Sub CmdSubmit_Click()
If Text1.Text = "bittusharma07" And Text2.Text = "bittu07" Then
MsgBox "Success", vbInformation
FrmLogin.Hide
FrmHome.Show
Else
    If Text1.Text = "" Then
    MsgBox "Please enter user name.", vbInformation
    End If
    If Text2.Text = "" Then
    MsgBox "Please enter password.", vbInformation
    End If
End If
End Sub
Private Sub CmdExit_Click()
ExitProject
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If keyascci = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Timer1_Timer()
If Flag = 0 Then
Lbl1.Left = Lbl1.Left + 100
    If Lbl1.Left > 9700 Then
    Flag = 1
    End If
Else
Lbl1.Left = Lbl1.Left - 100
    If Lbl1.Left < 2550 Then
    Flag = 0
    End If
End If
End Sub
