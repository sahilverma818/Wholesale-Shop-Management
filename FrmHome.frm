VERSION 5.00
Begin VB.Form FrmHome 
   Caption         =   "Home Form"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14355
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   Picture         =   "FrmHome.frx":0000
   ScaleHeight     =   7665
   ScaleWidth      =   14355
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdProductReport 
      BackColor       =   &H00C0C000&
      Caption         =   "Product Report"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CommandButton CmdCustomerReport 
      BackColor       =   &H00C0C000&
      Caption         =   "Customer Report"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   3735
   End
   Begin VB.CommandButton CmdSales 
      BackColor       =   &H00C0C000&
      Caption         =   "&Sales"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   3735
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0C000&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   3735
   End
   Begin VB.CommandButton CmdProduct 
      BackColor       =   &H00C0C000&
      Caption         =   "&Product"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton CmdCustomer 
      BackColor       =   &H00C0C000&
      Caption         =   "&Customer"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   1680
   End
   Begin VB.Label Lbl4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Lbl2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Lbl1 
      BackColor       =   &H00000000&
      Caption         =   "Time :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "FrmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCustomer_Click()
FrmCustomer.Show
Unload Me
End Sub
Private Sub CmdProduct_Click()
FrmProduct.Show
Unload Me
End Sub
Private Sub CmdSales_Click()
FrmSales.Show
Unload Me
End Sub
Private Sub Timer1_Timer()
Lbl2.Caption = Time
Lbl4.Caption = Date
End Sub
Private Sub CmdCustomerReport_Click()
FrmCustomerReport.Show
Unload Me
End Sub
Private Sub CmdProductReport_Click()
FrmProductReport.Show
Unload Me
End Sub
Private Sub CmdExit_Click()
ExitProject
End Sub
