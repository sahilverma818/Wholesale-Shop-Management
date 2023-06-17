VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmCustomerReport 
   Caption         =   "Customer Report"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "FrmCustomerReport.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H000080FF&
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
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton CmdHome 
      BackColor       =   &H000080FF&
      Caption         =   "Go To &Home"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      Width           =   2775
   End
   Begin VB.TextBox TxtCustomerName 
      BackColor       =   &H0080FF80&
      DataField       =   "Cust_Name"
      DataSource      =   "DataCustomer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4200
      TabIndex        =   4
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox TxtCity 
      BackColor       =   &H0080FF80&
      DataField       =   "City"
      DataSource      =   "DataCustomer"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox TxtMobno 
      BackColor       =   &H0080FF80&
      DataField       =   "Mob_No"
      DataSource      =   "DataCustomer"
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
      Left            =   4920
      TabIndex        =   2
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox TxtCustomerID 
      BackColor       =   &H0080FF80&
      DataField       =   "Cust_ID"
      DataSource      =   "DataCustomer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4200
      TabIndex        =   1
      Top             =   0
      Width           =   3855
   End
   Begin VB.Data DataSales 
      Caption         =   "Sales"
      Connect         =   "Access"
      DatabaseName    =   "D:\AC Cold Drink Shop\Acshop..s.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   11040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Sales"
      Top             =   1560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmCustomerReport.frx":2FA9A
      Height          =   3015
      Left            =   120
      OleObjectBlob   =   "FrmCustomerReport.frx":2FAB2
      TabIndex        =   0
      Top             =   2760
      Width           =   11175
   End
   Begin VB.Label LblCity 
      BackColor       =   &H0080FF80&
      Caption         =   "City of Customer :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label LblCustomerID 
      BackColor       =   &H0080FF80&
      Caption         =   "Customer ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label LblMobNo 
      BackColor       =   &H0080FF80&
      Caption         =   "Mob.No. of Customer :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "+91"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
End
Attribute VB_Name = "FrmCustomerReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdExit_Click()
ExitProject
End Sub

Private Sub CmdHome_Click()
FrmHome.Show
Unload Me
End Sub

Private Sub Form_Load()
Dim ID As String
ID = InputBox("Please enter Customer ID", "CUSTOMER ID")
DataSales.RecordSource = "Select * From Sales where Cust_ID ='" & ID & "'"
DataSales.Refresh
End Sub

