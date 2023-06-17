VERSION 5.00
Begin VB.Form FrmProduct 
   Caption         =   "Product Form"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "FrmProduct.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data DataProdID 
      Caption         =   "Product ID"
      Connect         =   "Access"
      DatabaseName    =   "D:\AC Cold Drink Shop\Acshop..s.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   10440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Product"
      Top             =   5880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton CmdSaveRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Save"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton CmdAddNewRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Add New"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton CmdDeleteRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Delete"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6840
      Width           =   2055
   End
   Begin VB.CommandButton CmdCancelRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Cancel"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton CmdUpdateRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Update"
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton CmdSearchRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Search"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Data DataProduct 
      Caption         =   "Product Details"
      Connect         =   "Access"
      DatabaseName    =   "D:\AC Cold Drink Shop\Acshop..s.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   10440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Product"
      Top             =   6840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox TxtProductID 
      BackColor       =   &H000080FF&
      DataField       =   "Prod_ID"
      DataSource      =   "DataProduct"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox TxtQOH 
      BackColor       =   &H000080FF&
      DataField       =   "QOH"
      DataSource      =   "DataProduct"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox TxtRate 
      BackColor       =   &H000080FF&
      DataField       =   "Rate"
      DataSource      =   "DataProduct"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton CmdNextRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Next"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   2055
   End
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton CmdLastRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Last"
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
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton CmdFirstRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&First"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox TxtProductName 
      BackColor       =   &H000080FF&
      DataField       =   "Prod_Name"
      DataSource      =   "DataProduct"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton CmdPreviousRecord 
      BackColor       =   &H000080FF&
      Caption         =   "&Previous"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   2055
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label LblProductRate 
      BackColor       =   &H000080FF&
      Caption         =   "Rate of Product :"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label LblProductID 
      BackColor       =   &H000080FF&
      Caption         =   "Product Details :"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label LblQoh 
      BackColor       =   &H000080FF&
      Caption         =   "Quantity on Hand :"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   4095
   End
End
Attribute VB_Name = "FrmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub MaxFalse()
CmdFirstRecord.Enabled = False
CmdPreviousRecord.Enabled = False
CmdNextRecord.Enabled = False
CmdLastRecord.Enabled = False
CmdAddNewRecord.Enabled = False
CmdSaveRecord.Enabled = True
CmdCancelRecord.Enabled = True
CmdUpdateRecord.Enabled = False
CmdDeleteRecord.Enabled = False
CmdSearchRecord.Enabled = False
CmdExit.Enabled = False
CmdHome.Enabled = False
TxtProductID.Locked = False
TxtProductName.Locked = False
TxtRate.Locked = False
TxtQOH.Locked = False
TxtProductID.SetFocus
End Sub
Public Sub MaxTrue()
CmdFirstRecord.Enabled = True
CmdPreviousRecord.Enabled = True
CmdNextRecord.Enabled = True
CmdLastRecord.Enabled = True
CmdAddNewRecord.Enabled = True
CmdSaveRecord.Enabled = False
CmdCancelRecord.Enabled = False
CmdUpdateRecord.Enabled = True
CmdDeleteRecord.Enabled = True
CmdSearchRecord.Enabled = True
CmdExit.Enabled = True
CmdHome.Enabled = True
TxtProductID.Locked = True
TxtProductName.Locked = True
TxtRate.Locked = True
TxtQOH.Locked = True
End Sub
Private Sub CmdFirstRecord_Click()
DataProduct.Recordset.MoveFirst
End Sub
Private Sub Cmdpreviousrecord_Click()
DataProduct.Recordset.MovePrevious
    If DataProduct.Recordset.BOF = True Then
    DataProduct.Recordset.MoveNext
    MsgBox "This is first record.", vbInformation, "First Record"
    End If
End Sub
Private Sub CmdNextRecord_Click()
DataProduct.Recordset.MoveNext
    If DataProduct.Recordset.EOF = True Then
    DataProduct.Recordset.MovePrevious
    MsgBox "This is last record.", vbInformation, "Last Record"
    End If
End Sub
Private Sub CmdLastRecord_Click()
DataProduct.Recordset.MoveLast
End Sub
Private Sub CmdAddNewRecord_Click()
DataProduct.Recordset.AddNew
Call MaxFalse
End Sub

Private Sub CmdCancelRecord_Click()
DataProduct.Recordset.CancelUpdate
Call MaxTrue
End Sub

Private Sub CmdSaveRecord_Click()
TxtProductID.Text = Trim(TxtProductID.Text)
TxtProductName.Text = Trim(TxtProductName.Text)
TxtRate.Text = Trim(TxtRate.Text)
TxtQOH.Text = Trim(TxtQOH.Text)
If TxtProductID.Text = "" Then
MsgBox "Please enter a valid Product ID", vbCritical, "Blank ID"
TxtProductID.SetFocus
ElseIf TxtProductName.Text = "" Then
MsgBox "Please enter a valid product name.", vbCritical, "Blank Name"
TxtProductName.SetFocus
ElseIf TxtRate.Text = "" Or Val(TxtRate.Text) <= 0 Then
MsgBox "Please enter a valid rate of product.", vbCritical, "Invalid Rate of Product"
TxtRate.Text = ""
TxtRate.SetFocus
ElseIf TxtQOH.Text = "" Or Val(TxtQOH.Text) <= 0 Then
MsgBox "Please enter a valid Quantity on Hand value.", vbCritical, "Invalid Quantity On Hand"
TxtQOH.Text = ""
TxtQOH.SetFocus
ElseIf DataProduct.Recordset.EOF = False Then
DataProduct.Recordset.Update
Call MaxTrue
Else
DataProdID.RecordSource = "Select Prod_ID from product where Prod_ID = '" & TxtProductID.Text & "'"
DataProdID.Refresh
    If DataProdID.Recordset.AbsolutePosition = 0 Then
    MsgBox "The Product ID is already exist. Please enter another product Id.", vbCritical, "Duplicate ID"
    TxtProductID.Text = ""
    TxtProductID.SetFocus
    Else
    DataProduct.Recordset.Update
    Call MaxTrue
    End If
End If
End Sub

Private Sub CmdUpdateRecord_Click()
DataProduct.Recordset.Edit
Call MaxFalse
End Sub
Private Sub CmdDeleteRecord_Click()
Msg = MsgBox("Are you sure want to Delete the Record of" & TxtProductID.Text & ".", vbYesNo + vbExclamation, "Delete Record")
    If Msg = vbYes Then
    DataProduct.Recordset.Delete
    DataProduct.Recordset.MoveNext
    If DataProduct.Recordset.EOF = True Then
    DataProduct.Recordset.MovePrevious
    End If
End If
End Sub
Private Sub CmdSearchRecord_Click()
ID = InputBox("Please enter a Product ID,Which Record you want to Search:->", "ID", "'")
DataProduct.Recordset.FindFirst ("Prod_ID ='" & ID & "'")
End Sub
Private Sub CmdHome_Click()
FrmHome.Show
Unload Me
End Sub
Private Sub CmdExit_Click()
ExitProject
End Sub
