VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmProductReport 
   Caption         =   "Product Report"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "FrmProductReport.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   9
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   2775
   End
   Begin VB.TextBox TxtProductName 
      BackColor       =   &H000080FF&
      DataField       =   "Prod_Name"
      DataSource      =   "DataProduct"
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
      Left            =   6360
      TabIndex        =   4
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox TxtRate 
      BackColor       =   &H000080FF&
      DataField       =   "Rate"
      DataSource      =   "DataProduct"
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
      Left            =   5520
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox TxtQOH 
      BackColor       =   &H000080FF&
      DataField       =   "QOH"
      DataSource      =   "DataProduct"
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
      Left            =   5520
      TabIndex        =   2
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox TxtProductID 
      BackColor       =   &H000080FF&
      DataField       =   "Prod_ID"
      DataSource      =   "DataProduct"
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
      Left            =   5520
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Data DataProduct 
      Caption         =   "Product"
      Connect         =   "Access"
      DatabaseName    =   "D:\AC Cold Drink Shop\Acshop..s.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   855
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Product"
      Top             =   8160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmProductReport.frx":217FA
      Height          =   3015
      Left            =   1080
      OleObjectBlob   =   "FrmProductReport.frx":21814
      TabIndex        =   0
      Top             =   2880
      Width           =   8535
   End
   Begin VB.Label LblQoh 
      BackColor       =   &H000080FF&
      Caption         =   "Quantity on Hand :"
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
      Left            =   1320
      TabIndex        =   7
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label LblProductID 
      BackColor       =   &H000080FF&
      Caption         =   "Product Details :"
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
      Left            =   1320
      TabIndex        =   6
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label LblProductRate 
      BackColor       =   &H000080FF&
      Caption         =   "Rate of Product :"
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
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   4095
   End
End
Attribute VB_Name = "FrmProductReport"
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
ID = InputBox("Please enter Product ID", "PRODUCT ID")
DataProduct.RecordSource = "Select * From Product where Prod_ID ='" & ID & "'"
DataProduct.Refresh
End Sub
