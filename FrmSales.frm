VERSION 5.00
Begin VB.Form FrmSales 
   Caption         =   "Sales Form"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "FrmSales.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7920
      Width           =   2055
   End
   Begin VB.CommandButton CmdAllCustomerRecord 
      BackColor       =   &H000080FF&
      Caption         =   "All Customer"
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7920
      Width           =   2775
   End
   Begin VB.TextBox TxtPrice 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3240
      TabIndex        =   7
      Top             =   3000
      Width           =   3855
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
      Height          =   585
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   13320
      Top             =   1920
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7200
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7200
      Width           =   2775
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7920
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7200
      Width           =   2055
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Data DataProduct 
      Caption         =   "Product Details"
      Connect         =   "Access"
      DatabaseName    =   "D:\AC Cold Drink Shop\Acshop..s.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Product"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data DataCustomer 
      Caption         =   "Customer Details"
      Connect         =   "Access"
      DatabaseName    =   "D:\AC Cold Drink Shop\Acshop..s.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   10320
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customer"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2535
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
      Height          =   585
      Left            =   5280
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
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
      Height          =   615
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox TxtCustomerName 
      BackColor       =   &H000080FF&
      DataField       =   "Cust_Name"
      DataSource      =   "DataCustomer"
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3015
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8640
      Width           =   2775
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6480
      Width           =   2055
   End
   Begin VB.TextBox TxtCustomerID 
      BackColor       =   &H000080FF&
      DataField       =   "Cust_ID"
      DataSource      =   "DataSales"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   735
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
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   2055
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6480
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7920
      Width           =   2055
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6480
      Width           =   2775
   End
   Begin VB.TextBox TxtProductID 
      BackColor       =   &H000080FF&
      DataField       =   "Prod_ID"
      DataSource      =   "DataSales"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox TxtQuantity 
      BackColor       =   &H000080FF&
      DataField       =   "Qty"
      DataSource      =   "DataSales"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3240
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Data DataSales 
      Caption         =   "Sales Details"
      Connect         =   "Access"
      DatabaseName    =   "D:\AC Cold Drink Shop\Acshop..s.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Sales"
      Top             =   6480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label LblPrice 
      BackColor       =   &H000080FF&
      Caption         =   "Total price :"
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
      Left            =   0
      TabIndex        =   25
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label LblRate 
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
      Left            =   0
      TabIndex        =   24
      Top             =   1560
      Width           =   3135
   End
   Begin VB.Label LblShowDate 
      BackColor       =   &H000080FF&
      Caption         =   " "
      DataField       =   "Date"
      DataSource      =   "DataSales"
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
      Left            =   14880
      TabIndex        =   23
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label LblShowTime 
      BackColor       =   &H000080FF&
      Caption         =   " "
      DataField       =   "Time"
      DataSource      =   "DataSales"
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
      Left            =   14880
      TabIndex        =   22
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label LblQuantity 
      BackColor       =   &H000080FF&
      Caption         =   "Quantity :"
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
      Left            =   0
      TabIndex        =   21
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label LblCustomerID 
      BackColor       =   &H000080FF&
      Caption         =   "Customer Detail :"
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
      Left            =   0
      TabIndex        =   20
      Top             =   120
      Width           =   3135
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
      Left            =   0
      TabIndex        =   19
      Top             =   840
      Width           =   3135
   End
End
Attribute VB_Name = "FrmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Quantity As String
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
CmdAllCustomerRecord.Enabled = False
CmdExit.Enabled = False
CmdHome.Enabled = False
TxtCustomerID.Locked = False
TxtCustomerName.Locked = False
TxtCustomerID.SetFocus
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
CmdAllCustomerRecord.Enabled = TrueH
CmdExit.Enabled = True
CmdHome.Enabled = True
TxtCustomerID.Locked = True
TxtCustomerName.Locked = True
End Sub
Private Sub CmdAddNewRecord_Click()
DataSales.Recordset.AddNew
'TxtCustomerName.Text = ""
'TxtProductName.Text = ""
'TxtRate.Text = ""
'TxtQOH.Text = ""
Call MaxFalse
End Sub

Private Sub CmdCancelRecord_Click()
DataSales.Recordset.CancelUpdate
Call MaxTrue
End Sub

Private Sub CmdDeleteRecord_Click()
Msg = MsgBox("Are you sure want to Delete the Record of" & TxtCustomerID.Text & ", and " & TxtProductID.Text, vbYesNo + vbExclamation, "Delete Record")
    If Msg = vbYes Then
    TxtQOH.Text = Val(TxtQOH.Text) + Val(TxtQuantity.Text)
    DataProduct.Recordset.Edit
    DataProduct.Recordset.Update
    DataSales.Recordset.Delete
    DataSales.Recordset.MoveNext
    If DataSales.Recordset.EOF = True Then
    DataSales.Recordset.MovePrevious
    End If
End If
End Sub

Private Sub CmdFirstRecord_Click()
DataSales.Recordset.MoveFirst
End Sub
Private Sub Cmdpreviousrecord_Click()
DataSales.Recordset.MovePrevious
    If DataSales.Recordset.BOF = True Then
    DataSales.Recordset.MoveNext
    MsgBox "This is First Record.", vbInformation, "Frist Record"
    End If
End Sub
Private Sub CmdNextRecord_Click()
DataSales.Recordset.MoveNext
    If DataSales.Recordset.EOF = True Then
    DataSales.Recordset.MovePrevious
    MsgBox "This is Last Record.", vbInformation, "Last Record"
    End If
End Sub
Private Sub CmdLastRecord_Click()
DataSales.Recordset.MoveLast
End Sub
Private Sub CmdHome_Click()
FrmHome.Show
Unload Me
End Sub
Private Sub CmdExit_Click()
ExitProject
End Sub
Private Sub CmdSaveRecord_Click()
TxtCustomerID.Text = Trim(TxtCustomerID.Text)
TxtProductID.Text = Trim(TxtProductID.Text)
If TxtCustomerID.Text = "" Then
MsgBox "Please enter any valid Customer ID.", vbCritical, "Invalid Customer ID"
TxtCustomerID.SetFocus
ElseIf TxtProductID.Text = "" Then
MsgBox "Please enter any valid Product ID.", vbCritical, "Invalid Product ID"
TxtCustomerID.SetFocus
Else
    If DataCustomer.Recordset.AbsolutePosition = -1 Then
    Msg = MsgBox("This Customer Does not exist in our Database. Do you want to Add a new Customer ?.", vbCritical + vbYesNo, "ERROR")
        If Msg = vbYes Then
        FrmCustomer.Show
        Unload Me
        Else
        TxtCustomerID.Text = ""
        TxtCustomerID.SetFocus
        End If
    ElseIf DataProduct.Recordset.AbsolutePosition = -1 Then
    Msg = MsgBox("This Product ID does not exist in out Database. Do you want to add a new Product ?", vbCritical + vbYesNo, "ERROR")
        If Msg = vbYes Then
        FrmProduct.Show
        Unload Me
        Else
        TxtProductID.Text = ""
        TxtProductID.SetFocus
        End If
    Else
    TxtQOH.Text = (Val(TxtQOH.Text) + Val(Quantity)) - Val(TxtQuantity)
    DataSales.Recordset.Update
    DataProduct.Recordset.Edit
    DataProduct.Recordset.Update
    End If
End If
End Sub

Private Sub CmdUpdateRecord_Click()
DataSales.Recordset.Edit
If DataCustomer.Recordset.EOF = False Then
Quantity = TxtQuantity.Text
End If
Call MaxTrue
End Sub
Private Sub Timer1_Timer()
If LblShowTime = "" Then
LblShowTime.Caption = Time
LblShowDate.Caption = Date
End If
End Sub

Private Sub TxtCustomerID_Change()
If Len(TxtCustomerID.Text) = 3 Then
DataCustomer.Refresh
DataCustomer.Recordset.FindFirst ("Cust_ID = '" & TxtCustomerID.Text & "'")
Else
TxtCustomerName.Text = ""
TxtProductName.Text = ""
TxtRate.Text = ""
TxtQOH.Text = ""
End If
End Sub

Private Sub TxtProductID_Change()
If Len(TxtProductID.Text) = 3 Then
DataProduct.Refresh
DataProduct.Recordset.FindFirst ("Prod_ID = '" & TxtProductID.Text & "'")
Else
TxtCustomerName.Text = ""
TxtProductName.Text = ""
TxtRate.Text = ""
TxtQOH.Text = ""
End If
End Sub

Private Sub TxtQuantity_Change()
If TxtRate.Text <> "" And TxtQuantity <> "" Then
TxtPrice.Text = ""
TxtPrice.Text = Val(TxtRate.Text) * Val(TxtQuantity.Text)
End If
End Sub
